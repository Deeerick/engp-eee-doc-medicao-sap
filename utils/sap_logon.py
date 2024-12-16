import win32com.client

from time import sleep
from utils.config import STRING_SAP


def login_sap():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnectionByConnectionString(
        STRING_SAP, True)
    session = connection.Children(0)

    return session
