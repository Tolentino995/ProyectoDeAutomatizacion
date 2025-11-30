import os
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

CREDENCIALES = "key_drive.json"
ID_CARPETA_DRIVE = "1BCb9JXw17Eix4Ux2BIzN1ok50lgFLb_H"


def login_drive():
    gauth = GoogleAuth()
    GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = CREDENCIALES

    gauth.LoadCredentialsFile(CREDENCIALES)

    if gauth.credentials is None:
        gauth.LocalWebserverAuth(port_numbers=[8092])
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()

    gauth.SaveCredentialsFile(CREDENCIALES)
    return GoogleDrive(gauth)


def subir_a_drive(ruta_archivo, id_folder=ID_CARPETA_DRIVE):
    drive = login_drive()
    archivo_drive = drive.CreateFile({
        "parents": [{"id": id_folder}],
        "title": os.path.basename(ruta_archivo)
    })
    archivo_drive.SetContentFile(ruta_archivo)
    archivo_drive.Upload()

    archivo_drive.InsertPermission({
        'type': 'anyone',
        'role': 'reader'
    })

    return archivo_drive['alternateLink']
