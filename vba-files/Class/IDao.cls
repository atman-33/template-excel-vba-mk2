VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' インターフェース：IDao
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Public Sub OpenConnection(servicename As String, username As String, password As String)
End Sub

Public Sub Query(sql As String)
End Sub

Public Function Execute(sql As String) As Long
End Function


Public Sub Save(insert As String, update As String)
End Sub

Public Sub SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                      updateColumns As Collection, conditions As Collection)
End Sub

Public Sub DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                             conditions As Collection)
End Sub


Public Sub PasteRecordsetToTable(table As ListObject)
End Sub

Public Sub CloseRecordset()
End Sub

Public Sub CloseConnection()
End Sub

