from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import json
import sys

orchestrator_connection = OrchestratorConnection("Slettebot Sharepoint", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
RobotCredentials = orchestrator_connection.get_credential("Robot365User")
username = RobotCredentials.username
password = RobotCredentials.password


def sharepoint_client(username: str, password: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"✅ Authenticated to SharePoint. Site Title: {web.properties['Title']}")
    return ctx


def delete_sharepoint_folder(folder_path: str, ctx: ClientContext, orchestrator_connection: OrchestratorConnection):
    """
    Recursively deletes a SharePoint folder and all its contents.
    """
    print(f"🗑 Deleting folder: {folder_path}")
    try:
        target_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        ctx.load(target_folder)
        ctx.execute_query()

        files = target_folder.files
        ctx.load(files)
        ctx.execute_query()
        for file in files:
            print(f"  Deleting file: {file.serverRelativeUrl}")
            file.delete_object()
        ctx.execute_query()

        subfolders = target_folder.folders
        ctx.load(subfolders)
        ctx.execute_query()
        for subfolder in subfolders:
            delete_sharepoint_folder(subfolder.serverRelativeUrl, ctx, orchestrator_connection)

        target_folder.delete_object()
        ctx.execute_query()
        orchestrator_connection.log_info(f"✅ Folder deleted: {folder_path}")
    except Exception as e:
        orchestrator_connection.log_info(f'An exception occurred: {e}')

#Hent mappenavn til sletning

# queue_json = json.loads(queue_element.data)
# mappenavn = queue_json.get('SharepointMappeNavn')

# Hent mappenavn til sletning
try:
    # queue_json = json.loads(queue_element.data)
    mappenavn = ""
except Exception as e:
    orchestrator_connection.log_info(f"❌ Fejl ved indlæsning af køelementets JSON: {e}")
    sys.exit()

# Beskyt mod farlige eller tomme navne
forbidden_names = ["", "Dokumentlister", "Aktindsigter", None]
if not mappenavn or mappenavn.strip() in forbidden_names:
    orchestrator_connection.log_info(f"❌ Mappenavn '{mappenavn}' er ugyldigt eller forbudt – sletning afbrydes.")
    sys.exit()

#Hent sharepoint site

# sharepoint_site_url = orchestrator_connection.get_constant("AktbobSharePointURL").value
sharepoint_site_url = "https://aarhuskommune.sharepoint.com/Teams/tea-teamsite11819"

#Definer mapper indenfor sharepointsite (Dokumentlistemappe og aktindsigtsmappe)

# folder_relative_url_aktliste = "/Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter"
# folder_relative_url_dokumentliste = "/Teams/tea-teamsite10506/Delte dokumenter/Dokumentlister/"
folder_relative_url_aktliste = f"/Teams/tea-teamsite11819/Delte dokumenter/{mappenavn}"
folder_relative_url_dokumentliste= f"/Teams/tea-teamsite11819/Delte dokumenter/{mappenavn}"

# Opret forbindelse
ctx = sharepoint_client(username, password, sharepoint_site_url, orchestrator_connection)

# # Slet dem
orchestrator_connection.log_info('Deleting aktliste folder')
delete_sharepoint_folder(folder_relative_url_aktliste, ctx= ctx, orchestrator_connection= orchestrator_connection)
orchestrator_connection.log_info('Deleting dokumentliste folder')
delete_sharepoint_folder(folder_relative_url_dokumentliste, ctx= ctx, orchestrator_connection= orchestrator_connection)