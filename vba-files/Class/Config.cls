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
' クラス：Config
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- 定数設定 ---- '
Const CONFIG_SHEET = "Config"
Const CONFIG_TABLE = "Config_tbl"
Const CONFIG_COL_KEY = "Key"
Const CONFIG_COL_ITEM = "Item"
' ------------------ '

Private dictionary_ As Object

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コンストラクタ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
    
    Set dictionary_ = CreateObject("Scripting.Dictionary")
    
    ' Configテーブルを格納
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(CONFIG_SHEET).ListObjects(CONFIG_TABLE)
    
    ' ConfigテーブルのKeyとItemを辞書に格納
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
' Summary : 指定したKeyのアイテムを取得
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Get Item(key As String) As Variant
    Item = dictionary_.Item(key)
End Property

