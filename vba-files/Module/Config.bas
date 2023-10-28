Attribute VB_Name = "Config"
Option Explicit

' �V�[�g/�e�[�u������
Public Const CONFIG_SHEET = "Config"
Public Const CONFIG_TABLE = "config"
Public Const CONFIG_COL_KEY = "Key"
Public Const CONFIG_COL_ITEM = "Item"

Public Const SQL_SHEET = "SQL"
Public Const SQL_TABLE = "sql"
Public Const SQL_COL_NAME = "Name"
Public Const SQL_COL_SHEET = "Sheet"
Public Const SQL_COL_TABLE = "Table"
Public Const SQL_COL_SQL = "SQL"

' DB�ڑ����@
Public Const DB_CONNECTION_MODE = "OracleOra"

' Oracle�ڑ��ݒ�
Public ORA_DATA_SOURCE As String
Public ORA_USER_ID As String
Public ORA_PASSWORD As String


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : Excel�V�[�g�Őݒ肵��Config�f�[�^�𔽉f
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub InitConfig()
    
    Dim i As Long
    
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    
    ' Config�e�[�u�����i�[
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(CONFIG_SHEET).ListObjects(CONFIG_TABLE)
    
    ' Config�e�[�u����Key��Item�������Ɋi�[
    Dim key As String, item As String
    
    For i = 1 To table.ListRows.Count
        key = table.ListColumns(CONFIG_COL_KEY).DataBodyRange(i).Value
        item = table.ListColumns(CONFIG_COL_ITEM).DataBodyRange(i).Value
    
        dictionary.Add key, item
    
        Debug.Print "Key:" & key & " Item:" & item & " �i�["
    
    Next i
    
    ' Public�ϐ��Ɋi�[
    ORA_DATA_SOURCE = dictionary.item("ORA_DATA_SOURCE")
    ORA_USER_ID = dictionary.item("ORA_USER_ID")
    ORA_PASSWORD = dictionary.item("ORA_PASSWORD")

End Sub

