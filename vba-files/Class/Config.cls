VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' �N���X�FConfig
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- �萔�ݒ� ---- '
Const CONFIG_SHEET = "Config"
Const CONFIG_TABLE = "ConfigTable"
Const CONFIG_COL_KEY = "Key"
Const CONFIG_COL_ITEM = "Item"
' ------------------ '

Private dictionary_ As Object

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �R���X�g���N�^
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
    
    Set dictionary_ = CreateObject("Scripting.Dictionary")
    
    ' Config�e�[�u�����i�[
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(CONFIG_SHEET).ListObjects(CONFIG_TABLE)
    
    ' Config�e�[�u����Key��Item�������Ɋi�[
    Dim key As String, Item As String
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
        key = table.ListColumns(CONFIG_COL_KEY).DataBodyRange(i).Value
        Item = table.ListColumns(CONFIG_COL_ITEM).DataBodyRange(i).Value
    
        Call dictionary_.Add(key, Item)
    Next i
    
    ' ---- Debug ---- '
    Dim varItem As Variant
    Dim str As String
    For Each varItem In dictionary_
        str = str & varItem & ":" & dictionary_.Item(varItem) & vbCrLf
    Next
    
    Debug.Print str
'    Debug.Print dictionary_.Item("ORA_DATA_SOURCE")
    
    ' --------------- '
    
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �w�肵��Key�̃A�C�e�����擾
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Get Item(key As String) As Variant
    Item = dictionary_.Item(key)
End Property
