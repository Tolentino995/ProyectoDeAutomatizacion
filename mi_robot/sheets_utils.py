from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
KEY = "key_sheets.json"
SPREADSHEET_ID = "1o0SG78GLAPsTK-jGKVkNf2iOHyX9ycJMP-9PzlZYMro"
HOJA = "Hoja 1"


def leer_google_sheet_columna(columna_buscada):
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    valores = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{HOJA}!A:Z"
    ).execute().get('values', [])

    encabezados = valores[0]
    filas = valores[1:]

    idx = encabezados.index(columna_buscada)
    return [fila[idx] for fila in filas if len(fila) > idx and fila[idx].strip()]


def escribir_link_en_sheet(cliente, link):
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    valores = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{HOJA}!A:B"
    ).execute().get('values', [])

    for i, fila in enumerate(valores):
        if fila and fila[0] == cliente:
            fila_destino = i + 1
            break
    else:
        return

    rango = f"{HOJA}!B{fila_destino}"
    body = {"values": [[f'=HYPERLINK("{link}"; "Ver PDF")']]}

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rango,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()


def escribir_estado_deuda(cliente, estado):
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    valores = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{HOJA}!A:Z"
    ).execute().get('values', [])

    encabezados = valores[0]
    filas = valores[1:]

    try:
        col = encabezados.index("Estado de Cuenta")
    except:
        return

    for i, fila in enumerate(filas, start=2):
        if fila and fila[0] == cliente:
            fila_destino = i
            break
    else:
        return

    letra = chr(ord('A') + col)
    rango = f"{HOJA}!{letra}{fila_destino}"

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rango,
        valueInputOption="RAW",
        body={"values": [[estado]]}
    ).execute()
