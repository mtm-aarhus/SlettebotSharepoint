from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import json
import pyodbc

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    RobotCredentials = orchestrator_connection.get_credential("Robot365User")
    server = orchestrator_connection.get_constant('AktbobServer').value
    database = orchestrator_connection.get_constant('AktbobDatabase').value
    databasebruger = orchestrator_connection.get_credential('AktbobDatabaseBruger')
    username = RobotCredentials.username
    password = RobotCredentials.password
    connection_string = (
        "Driver={ODBC Driver 17 for SQL Server};"
        f"Server=tcp:{server}.database.windows.net,1433;"
        f"Database={database};"
        f"Uid={databasebruger.username};"
        f"Pwd={databasebruger.password};"
        "Encrypt=yes;"
        "TrustServerCertificate=no;"
    )
    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()

    def mark_as_deleted(deskpro_id, cursor, conn):
        cursor.execute("""
            UPDATE dbo.Tickets
            SET SlettetSharepoint = 1
            WHERE DeskproId = ?
        """, deskpro_id)

        conn.commit()



    def sharepoint_client(username: str, password: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
        """
        Creates and returns a SharePoint client context.
        """
        ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        orchestrator_connection.log_info(f"✅ Authenticated to SharePoint. Site Title: {web.properties['Title']}")
        return ctx


    def delete_sharepoint_folder(folder_path: str, ctx: ClientContext, orchestrator_connection: OrchestratorConnection, deskpro_id, cursor, conn):
        """
        Recursively deletes a SharePoint folder and all its contents.
        """
        orchestrator_connection.log_info(f"🗑 Deleting folder: {folder_path}")
        try:
            target_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
            ctx.load(target_folder)
            ctx.execute_query()

            files = target_folder.files
            ctx.load(files)
            ctx.execute_query()
            for file in files:
                orchestrator_connection.log_info(f"  Deleting file: {file.serverRelativeUrl}")
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
            mark_as_deleted(deskpro_id, cursor, conn)
        except Exception as e:
            orchestrator_connection.log_info(f'An exception occurred: {e}')

    #Hent mappenavn til sletning
    # Hent mappenavn til sletning
    try:
        queue_json = json.loads(queue_element.data)
        deskpro_id = queue_json.get('DeskproId')
        mappenavn = queue_json.get('SharepointMappeNavn')
    except Exception as e:
        orchestrator_connection.log_info(f"❌ Fejl ved indlæsning af køelementets JSON: {e}")
        return

    # Beskyt mod farlige eller tomme navne
    forbidden_names = ["", "Dokumentlister", "Aktindsigter", None]
    if not mappenavn or mappenavn.strip() in forbidden_names:
        orchestrator_connection.log_info(f"❌ Mappenavn '{mappenavn}' er ugyldigt eller forbudt – sletning afbrydes.")
        return
    #Hent sharepoint site

    sharepoint_site_url = orchestrator_connection.get_constant("AktbobSharePointURL").value

    #Definer mapper indenfor sharepointsite (Dokumentlistemappe og aktindsigtsmappe)

    folder_relative_url_aktliste = f"/Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter/{mappenavn}"
    folder_relative_url_dokumentliste = f"/Teams/tea-teamsite10506/Delte dokumenter/Dokumentlister/{mappenavn}"

    # Opret forbindelse
    ctx = sharepoint_client(username, password, sharepoint_site_url, orchestrator_connection)

    # # Slet dem
    orchestrator_connection.log_info('Deleting aktliste folder')
    delete_sharepoint_folder(folder_relative_url_aktliste, ctx= ctx, orchestrator_connection= orchestrator_connection, deskpro_id, cursor, conn)
    orchestrator_connection.log_info('Deleting dokumentliste folder')
    delete_sharepoint_folder(folder_relative_url_dokumentliste, ctx= ctx, orchestrator_connection= orchestrator_connection, deskpro_id, cursor, conn)