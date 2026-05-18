import os, requests, json, base64
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime, date

CLIENT_ID       = os.environ["CLIENT_ID"]
TENANT_ID       = os.environ["TENANT_ID"]
CLIENT_SECRET   = os.environ["CLIENT_SECRET"]
REFRESH_TOKEN   = os.environ["REFRESH_TOKEN"]
FILE_URL        = os.environ["SHAREPOINT_FILE_URL"]
G_CLIENT_ID     = os.environ["GOOGLE_CLIENT_ID"]
G_CLIENT_SECRET = os.environ["GOOGLE_CLIENT_SECRET"]
G_REFRESH_TOKEN = os.environ["GOOGLE_REFRESH_TOKEN"]
CUDYR_2026_FOLDER = "1DdnURFrLf9pSX67J1WzsCDLmOUIQu0D5"
GH_PAT  = os.environ.get("GH_PAT", "")
GH_REPO = os.environ.get("GITHUB_REPOSITORY", "")

MESES_VALIDOS = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                 "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
MESES_ES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12
}

def get_ms_token():
    resp = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"client_id":CLIENT_ID,"client_secret":CLIENT_SECRET,
              "refresh_token":REFRESH_TOKEN,"grant_type":"refresh_token",
              "scope":"Files.Read offline_access"}
    )
    data = resp.json()
    if "access_token" not in data:
        raise Exception(f"Error MS auth: {data}")
    new_rt = data.get("refresh_token")
    if new_rt and new_rt != REFRESH_TOKEN and GH_PAT and GH_REPO:
        try:
            update_github_secret("REFRESH_TOKEN", new_rt)
            print("✓ MS refresh token renovado")
        except Exception as e:
            print(f"⚠ No se pudo renovar MS token: {e}")
    return data["access_token"]

def get_google_token():
    resp = requests.post(
        "https://oauth2.googleapis.com/token",
        data={"client_id":G_CLIENT_ID,"client_secret":G_CLIENT_SECRET,
              "refresh_token":G_REFRESH_TOKEN,"grant_type":"refresh_token"}
    )
    data = resp.json()
    if "access_token" not in data:
        raise Exception(f"Error Google auth: {data}")
    print(f"  Google token OK, scope: {data.get('scope','?')}")
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
    requests.put(
        f"https://api.github.com/repos/{GH_REPO}/actions/secrets/{secret_name}",
        headers=headers,
        json={"encrypted_value": encrypted, "key_id": key_data["key_id"]}
    ).raise_for_status()

def get_sharepoint_file(token):
    headers = {"Authorization": f"Bearer {token}"}
    b64 = base64.urlsafe_b64encode(FILE_URL.encode()).decode().rstrip("=")
    url = f"https://graph.microsoft.com/v1.0/shares/u!{b64}/driveItem/content"
    resp = requests.get(url, headers=headers, allow_redirects=True)
    if resp.status_code != 200:
        raise Exception(f"Error descargando Excel: {resp.status_code} - {resp.text[:300]}")
    return resp.content

def list_drive_files(token, folder_id):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://www.googleapis.com/drive/v3/files?q='{folder_id}'+in+parents+and+mimeType='application/vnd.google-apps.spreadsheet'&fields=files(id,name)&pageSize=50"
    resp = requests.get(url, headers=headers)
    print(f"  Drive list status: {resp.status_code}")
    if resp.status_code != 200:
        print(f"  Drive error: {resp.text[:300]}")
    resp.raise_for_status()
    return resp.json().get("files", [])

def get_cudyr_via_sheets_api(token, file_id):
    headers = {"Authorization": f"Bearer {token}"}

    # Obtener el gid de la hoja TOTAL MENSUAL
    meta_url = f"https://sheets.googleapis.com/v4/spreadsheets/{file_id}?fields=sheets.properties"
    meta_resp = requests.get(meta_url, headers=headers)
    if meta_resp.status_code != 200:
        print(f"    Meta error {meta_resp.status_code}: {meta_resp.text[:200]}")
        return None

    sheets = meta_resp.json().get("sheets", [])
    gid = None
    for s in sheets:
        if "TOTAL" in s["properties"]["title"].upper():
            gid = s["properties"]["sheetId"]
            break

    if gid is None:
        print(f"    ⚠ No hay hoja TOTAL MENSUAL")
        return None

    # Exportar como CSV (más confiable que la Sheets API para leer valores)
    export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv&gid={gid}"
    resp = requests.get(export_url, headers=headers)
    if resp.status_code != 200:
        print(f"    Export error {resp.status_code}")
        return None

    # Parsear CSV línea por línea
    try:
        text = resp.content.decode('utf-8')
    except:
        text = resp.content.decode('latin-1')

    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')

    total_criticos = total_medios = total_basicos = total_d = 0
    bloques = 0

    for i, line in enumerate(lines):
        line_up = line.upper()
        if 'RESUMEN DIARIO' in line_up and i + 1 < len(lines):
            next_line = lines[i + 1]
            cols = next_line.split(',')
            nums = []
            for cell in cols:
                cell = cell.strip().strip('"')
                try:
                    nums.append(int(cell))
                except:
                    pass
            if len(nums) >= 9:
                total_criticos += nums[0] + nums[1] + nums[2]
                total_medios   += nums[3] + nums[4] + nums[5]
                total_basicos  += nums[6] + nums[7] + nums[8]
                if len(nums) >= 12:
                    total_d += nums[9] + nums[10] + nums[11]
                bloques += 1

    print(f"    Bloques: {bloques} | C={total_criticos} M={total_medios} B={total_basicos} D={total_d}")

    total_pacientes = total_criticos + total_medios + total_basicos + total_d
    if total_pacientes == 0:
        print(f"    ⚠ Sin datos suficientes")
        return None

    return {
        "criticos_pct":  round(total_criticos / total_pacientes * 100, 1),
        "medios_pct":    round(total_medios   / total_pacientes * 100, 1),
        "basicos_pct":   round(total_basicos  / total_pacientes * 100, 1),
        "total_dias":    total_pacientes,
        "criticos_dias": total_criticos,
        "medios_dias":   total_medios,
        "basicos_dias":  total_basicos,
    }

def extract_month_from_name(name):
    name_lower = name.lower()
    for mes in MESES_ES:
        if mes in name_lower:
            return mes.upper(), MESES_ES[mes]
    return None, None

def fetch_cudyr_data(g_token):
    print("  Listando archivos CUDYR 2026...")
    files = list_drive_files(g_token, CUDYR_2026_FOLDER)
    print(f"  → {len(files)} archivos encontrados")
    cudyr = {}
    for f in files:
        mes_nombre, _ = extract_month_from_name(f["name"])
        if not mes_nombre:
            continue
        print(f"  Procesando CUDYR {mes_nombre} ({f['name']})...")
        datos = get_cudyr_via_sheets_api(g_token, f["id"])
        if datos:
            cudyr[mes_nombre] = datos
            print(f"  → C:{datos['criticos_pct']}% M:{datos['medios_pct']}% B:{datos['basicos_pct']}%")
        else:
            print(f"  ⚠ Sin datos: {f['name']}")
    return cudyr

def parse_mes(ws):
    datos_diarios = []
    dotacion = 29
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell and "DOTACION" in str(cell).upper():
                idx = list(row).index(cell)
                if idx + 1 < len(row) and row[idx+1]:
                    try: dotacion = int(row[idx+1])
                    except: pass
        first = row[0]
        if first is None:
            continue
        fecha = None
        if isinstance(first, datetime):
            fecha = first
        else:
            try: fecha = datetime.strptime(str(first).strip(), "%d/%m/%Y")
            except: pass
        if fecha is None:
            continue
        def si(v):
            try: return int(v) if v is not None else 0
            except: return 0
        fila = list(row)
        datos_diarios.append({
            "fecha": fecha.strftime("%Y-%m-%d"),
            "existencia": si(fila[1]), "ing_urgencia": si(fila[2]),
            "ing_aps": si(fila[3]), "ing_cae": si(fila[4]),
            "ing_otros_hosp": si(fila[5]), "ing_otra_proc": si(fila[6]),
            "ing_mismo_serv": si(fila[7]), "total_ingresos": si(fila[8]),
            "egr_alta": si(fila[9]), "egr_traslado": si(fila[10]),
            "egr_fallecidos": si(fila[11]), "total_egresos": si(fila[12]),
            "mismo_dia": si(fila[13]), "camas_disponibles": si(fila[14]),
            "camas_ocupadas": si(fila[15]), "dias_estada": si(fila[16]),
        })
    if not datos_diarios:
        return None
    total_ingresos   = sum(d["total_ingresos"] for d in datos_diarios)
    total_egresos    = sum(d["total_egresos"] for d in datos_diarios)
    total_fallecidos = sum(d["egr_fallecidos"] for d in datos_diarios)
    total_dias_estada= sum(d["dias_estada"] for d in datos_diarios)
    dias = len(datos_diarios)
    ocup_prom = round(sum(d["camas_ocupadas"] for d in datos_diarios)/dias, 1) if dias else 0
    pct_ocup  = round(ocup_prom/dotacion*100, 1) if dotacion else 0
    estada    = round(total_dias_estada/total_egresos, 1) if total_egresos else 0
    return {
        "dias": datos_diarios,
        "resumen": {
            "dotacion": dotacion, "total_ingresos": total_ingresos,
            "total_egresos": total_egresos, "total_fallecidos": total_fallecidos,
            "ocupacion_promedio": ocup_prom, "porcentaje_ocupacion": pct_ocup,
            "estada_promedio": estada,
            "ing_urgencia": sum(d["ing_urgencia"] for d in datos_diarios),
            "ing_aps": sum(d["ing_aps"] for d in datos_diarios),
            "ing_otros_hosp": sum(d["ing_otros_hosp"] for d in datos_diarios),
        }
    }

def parse_sociosanitarios(ws):
    casos = []
    hoy = date.today()
    header_found = False
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        if str(row[0]).strip().upper() == "OTROS":
            header_found = True
            continue
        if not header_found:
            continue
        if str(row[0]).strip().upper() != "FIJO":
            continue
        def sv(v): return str(v).strip() if v is not None else ""
        nombre_completo = f"{sv(row[3])} {sv(row[1])} {sv(row[2])}".strip()
        fecha_ing_raw = row[14]
        fecha_ing_str = ""
        dias_hosp = None
        if fecha_ing_raw:
            if isinstance(fecha_ing_raw, datetime):
                fi = fecha_ing_raw.date()
            else:
                try: fi = datetime.strptime(str(fecha_ing_raw).strip(), "%d/%m/%Y").date()
                except: fi = None
            if fi:
                fecha_ing_str = fi.strftime("%d/%m/%Y")
                dias_hosp = (hoy - fi).days
        if nombre_completo:
            casos.append({
                "nombre": nombre_completo, "rut": sv(row[5]), "edad": sv(row[6]),
                "fecha_ingreso": fecha_ing_str, "dias_hospitalizados": dias_hosp,
                "sala": sv(row[12]), "cama": sv(row[13]),
                "procedencia": sv(row[8]), "prevision": sv(row[9]),
            })
    casos.sort(key=lambda x: x["dias_hospitalizados"] or 0, reverse=True)
    return casos

def main():
    print("─── Microsoft / SharePoint ───")
    ms_token = get_ms_token()
    print("Descargando Excel censo...")
    content = get_sharepoint_file(ms_token)
    print("Procesando Excel...")
    wb = load_workbook(BytesIO(content), data_only=True)

    resultado = {
        "actualizado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "meses": {}, "sociosanitarios": [], "cudyr": {}
    }

    for sheet_name in wb.sheetnames:
        nombre_upper = sheet_name.strip().upper()
        if nombre_upper in MESES_VALIDOS:
            print(f"  Procesando mes: {sheet_name}")
            datos = parse_mes(wb[sheet_name])
            if datos:
                resultado["meses"][nombre_upper] = datos
        elif nombre_upper == "DATOS":
            print("  Procesando casos sociosanitarios...")
            casos = parse_sociosanitarios(wb[sheet_name])
            resultado["sociosanitarios"] = casos
            print(f"  → {len(casos)} casos FIJO")

    print("\n─── Google Drive / CUDYR ───")
    try:
        g_token = get_google_token()
        cudyr = fetch_cudyr_data(g_token)
        resultado["cudyr"] = cudyr
        print(f"✓ CUDYR: {list(cudyr.keys())}")
    except Exception as e:
        print(f"⚠ Error CUDYR: {e}")
        resultado["cudyr"] = {}

    os.makedirs("data", exist_ok=True)
    with open("data/censo.json", "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f"\n✓ Meses censo: {list(resultado['meses'].keys())}")
    print(f"✓ Sociosanitarios: {len(resultado['sociosanitarios'])}")
    print(f"✓ CUDYR meses: {list(resultado['cudyr'].keys())}")

if __name__ == "__main__":
    main()
