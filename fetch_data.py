import os, requests, json, base64
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime, date

CLIENT_ID = os.environ["CLIENT_ID"]
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
REFRESH_TOKEN = os.environ["REFRESH_TOKEN"]
FILE_URL = os.environ["SHAREPOINT_FILE_URL"]
GH_PAT = os.environ.get("GH_PAT", "")
GH_REPO = os.environ.get("GITHUB_REPOSITORY", "")

MESES_VALIDOS = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                 "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

def get_token():
    resp = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "refresh_token": REFRESH_TOKEN,
            "grant_type": "refresh_token",
            "scope": "Files.Read offline_access"
        }
    )
    data = resp.json()
    if "access_token" not in data:
        raise Exception(f"Error autenticando: {data}")
    new_refresh = data.get("refresh_token")
    if new_refresh and new_refresh != REFRESH_TOKEN and GH_PAT and GH_REPO:
        try:
            update_github_secret("REFRESH_TOKEN", new_refresh)
            print("✓ Refresh token renovado automáticamente")
        except Exception as e:
            print(f"⚠ No se pudo renovar el token: {e}")
    return data["access_token"]

def update_github_secret(secret_name, secret_value):
    headers = {"Authorization": f"token {GH_PAT}", "Accept": "application/vnd.github.v3+json"}
    r = requests.get(f"https://api.github.com/repos/{GH_REPO}/actions/secrets/public-key", headers=headers)
    r.raise_for_status()
    key_data = r.json()
    from base64 import b64encode
    from nacl import encoding, public
    pk = public.PublicKey(key_data["key"].encode(), encoding.Base64Encoder)
    box = public.SealedBox(pk)
    encrypted = b64encode(box.encrypt(secret_value.encode())).decode()
    r2 = requests.put(
        f"https://api.github.com/repos/{GH_REPO}/actions/secrets/{secret_name}",
        headers=headers,
        json={"encrypted_value": encrypted, "key_id": key_data["key_id"]}
    )
    r2.raise_for_status()

def get_file_content(token):
    headers = {"Authorization": f"Bearer {token}"}
    b64 = base64.urlsafe_b64encode(FILE_URL.encode()).decode().rstrip("=")
    share_token = "u!" + b64
    url = f"https://graph.microsoft.com/v1.0/shares/{share_token}/driveItem/content"
    resp = requests.get(url, headers=headers, allow_redirects=True)
    if resp.status_code != 200:
        raise Exception(f"Error descargando: {resp.status_code} - {resp.text[:300]}")
    return resp.content

def parse_sociosanitarios(ws):
    """Lee la hoja DATOS y extrae pacientes con FIJO en columna A (casos sociosanitarios activos)"""
    casos = []
    hoy = date.today()
    header_found = False

    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue

        # Detectar fila de encabezado
        if str(row[0]).strip().upper() == "OTROS":
            header_found = True
            continue

        if not header_found:
            continue

        otros = str(row[0]).strip().upper() if row[0] else ""
        if otros != "FIJO":
            continue

        # Extraer datos del paciente
        def sv(v):
            return str(v).strip() if v is not None else ""

        ap_paterno = sv(row[1])
        ap_materno = sv(row[2])
        nombre = sv(row[3])
        nombre_completo = f"{nombre} {ap_paterno} {ap_materno}".strip()

        rut = sv(row[5])
        edad = sv(row[6])
        procedencia = sv(row[8])
        prevision = sv(row[9])
        sala = sv(row[12])
        cama = sv(row[13])

        # Fecha ingreso
        fecha_ing_raw = row[14]
        fecha_ing_str = ""
        dias_hospitalizados = None

        if fecha_ing_raw:
            if isinstance(fecha_ing_raw, datetime):
                fi = fecha_ing_raw.date()
            else:
                try:
                    fi = datetime.strptime(str(fecha_ing_raw).strip(), "%d/%m/%Y").date()
                except:
                    fi = None
            if fi:
                fecha_ing_str = fi.strftime("%d/%m/%Y")
                dias_hospitalizados = (hoy - fi).days

        if nombre_completo:
            casos.append({
                "nombre": nombre_completo,
                "rut": rut,
                "edad": edad,
                "fecha_ingreso": fecha_ing_str,
                "dias_hospitalizados": dias_hospitalizados,
                "sala": sala,
                "cama": cama,
                "procedencia": procedencia,
                "prevision": prevision,
            })

    # Ordenar por días hospitalizados descendente
    casos.sort(key=lambda x: x["dias_hospitalizados"] or 0, reverse=True)
    return casos

def parse_mes(ws):
    datos_diarios = []
    dotacion = 29

    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell and "DOTACION" in str(cell).upper():
                idx = list(row).index(cell)
                if idx + 1 < len(row) and row[idx+1]:
                    try:
                        dotacion = int(row[idx+1])
                    except:
                        pass

        first = row[0]
        if first is None:
            continue
        fecha = None
        if isinstance(first, datetime):
            fecha = first
        else:
            try:
                fecha = datetime.strptime(str(first).strip(), "%d/%m/%Y")
            except:
                pass
        if fecha is None:
            continue

        def safe_int(v):
            try:
                return int(v) if v is not None else 0
            except:
                return 0

        fila = list(row)
        datos_diarios.append({
            "fecha": fecha.strftime("%Y-%m-%d"),
            "existencia": safe_int(fila[1]),
            "ing_urgencia": safe_int(fila[2]),
            "ing_aps": safe_int(fila[3]),
            "ing_cae": safe_int(fila[4]),
            "ing_otros_hosp": safe_int(fila[5]),
            "ing_otra_proc": safe_int(fila[6]),
            "ing_mismo_serv": safe_int(fila[7]),
            "total_ingresos": safe_int(fila[8]),
            "egr_alta": safe_int(fila[9]),
            "egr_traslado": safe_int(fila[10]),
            "egr_fallecidos": safe_int(fila[11]),
            "total_egresos": safe_int(fila[12]),
            "mismo_dia": safe_int(fila[13]),
            "camas_disponibles": safe_int(fila[14]),
            "camas_ocupadas": safe_int(fila[15]),
            "dias_estada": safe_int(fila[16]),
        })

    if not datos_diarios:
        return None

    total_ingresos = sum(d["total_ingresos"] for d in datos_diarios)
    total_egresos = sum(d["total_egresos"] for d in datos_diarios)
    total_fallecidos = sum(d["egr_fallecidos"] for d in datos_diarios)
    total_dias_estada = sum(d["dias_estada"] for d in datos_diarios)
    dias = len(datos_diarios)
    ocupacion_prom = round(sum(d["camas_ocupadas"] for d in datos_diarios) / dias, 1) if dias > 0 else 0
    pct_ocupacion = round((ocupacion_prom / dotacion * 100), 1) if dotacion > 0 else 0
    estada_prom = round(total_dias_estada / total_egresos, 1) if total_egresos > 0 else 0

    return {
        "dias": datos_diarios,
        "resumen": {
            "dotacion": dotacion,
            "total_ingresos": total_ingresos,
            "total_egresos": total_egresos,
            "total_fallecidos": total_fallecidos,
            "ocupacion_promedio": ocupacion_prom,
            "porcentaje_ocupacion": pct_ocupacion,
            "estada_promedio": estada_prom,
            "ing_urgencia": sum(d["ing_urgencia"] for d in datos_diarios),
            "ing_aps": sum(d["ing_aps"] for d in datos_diarios),
            "ing_otros_hosp": sum(d["ing_otros_hosp"] for d in datos_diarios),
        }
    }

def main():
    print("Obteniendo token...")
    token = get_token()
    print("Descargando Excel desde SharePoint...")
    content = get_file_content(token)
    print("Procesando Excel...")
    wb = load_workbook(BytesIO(content), data_only=True)

    resultado = {
        "actualizado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "meses": {},
        "sociosanitarios": []
    }

    for sheet_name in wb.sheetnames:
        nombre_upper = sheet_name.strip().upper()

        if nombre_upper in MESES_VALIDOS:
            print(f"  Procesando mes: {sheet_name}")
            ws = wb[sheet_name]
            datos = parse_mes(ws)
            if datos:
                resultado["meses"][nombre_upper] = datos

        elif nombre_upper == "DATOS":
            print(f"  Procesando casos sociosanitarios...")
            ws = wb[sheet_name]
            casos = parse_sociosanitarios(ws)
            resultado["sociosanitarios"] = casos
            print(f"  → {len(casos)} casos FIJO encontrados")

    os.makedirs("data", exist_ok=True)
    with open("data/censo.json", "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f"✓ Listo. Meses: {list(resultado['meses'].keys())}")
    print(f"✓ Casos sociosanitarios: {len(resultado['sociosanitarios'])}")

if __name__ == "__main__":
    main()
