import os, requests, json
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

CLIENT_ID = os.environ["CLIENT_ID"]
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
REFRESH_TOKEN = os.environ["REFRESH_TOKEN"]
FILE_URL = os.environ["SHAREPOINT_FILE_URL"]

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
    return data["access_token"]

def get_file_content(token):
    headers = {"Authorization": f"Bearer {token}"}
    import base64
    b64 = base64.urlsafe_b64encode(FILE_URL.encode()).decode().rstrip("=")
    share_token = "u!" + b64
    url = f"https://graph.microsoft.com/v1.0/shares/{share_token}/driveItem/content"
    resp = requests.get(url, headers=headers, allow_redirects=True)
    if resp.status_code != 200:
        raise Exception(f"Error descargando: {resp.status_code} - {resp.text[:300]}")
    return resp.content

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
        "meses": {}
    }

    for sheet_name in wb.sheetnames:
        nombre_upper = sheet_name.strip().upper()
        if nombre_upper in MESES_VALIDOS:
            print(f"  Procesando: {sheet_name}")
            ws = wb[sheet_name]
            datos = parse_mes(ws)
            if datos:
                resultado["meses"][nombre_upper] = datos

    os.makedirs("data", exist_ok=True)
    with open("data/censo.json", "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f"✓ Listo. Meses: {list(resultado['meses'].keys())}")

if __name__ == "__main__":
    main()
