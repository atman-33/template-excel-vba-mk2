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
' Class: Config
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- 定数設定 ---- '
Const CONFIG_SHEET = "Config"
Const CONFIG_TABLE = "Config_tbl"
Const CONFIG_COL_KEY = "Key"
Const CONFIG_COL_ITEM = "Item"
' ------------------ '

Private m_dictionary As Object

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コンストラクタ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
    
    Set m_dictionary = CreateObject("Scripting.Dictionary")
    
    ' Configテーブルを格納
    Dim varTable As ListObject
    Set varTable = ThisWorkbook.Worksheets(CONFIG_SHEET).ListObjects(CONFIG_TABLE)
    
    ' ConfigテーブルのKeyとItemを辞書に格納
    Dim varKey As String, varItem As String
    
    Dim i As Long
    For i = 1 To varTable.ListRows.Count
        varKey = varTable.ListColumns(CONFIG_COL_KEY).DataBodyRange(i).Value
        varItem = varTable.ListColumns(CONFIG_COL_ITEM).DataBodyRange(i).Value
    
        Call m_dictionary.Add(varKey, varItem)
    Next i
    
    ' ---- Debug ---- '
'    Dim var As Variant
'    Dim str As String
'    For Each var In m_dictionary
'        str = str & varItem & ":" & m_dictionary.Item(varItem) & vbCrLf
'    Next
'
'    Debug.Print str
    ' --------------- '
    
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 指定したKeyのアイテムを取得
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Get Item(key As String) As Variant
    Item = m_dictionary.Item(key)
End Property

