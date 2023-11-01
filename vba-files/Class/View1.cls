VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "View1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const SHEET_NAME = "Sample1"
Private Const TABLE_NAME = "sample1_tbl1"

Private Sub CommandButton1_Click()
    MsgBox "test"
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 各テーブルのデータを更新
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub GetDataButton_Click()

    Dim tableCls As New TableSample
    Call tableCls.LoadData
    
    MsgBox "データを更新しました。"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 昇順に並び替え
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SampleIdOrderByButton_Click()

    Dim tableCls As New TableSample
    Call tableCls.OrderBy("ID")

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 降順に並び替え
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SampleIdOrderByDescendingButton_Click()

    Dim tableCls As New TableSample
    Call tableCls.OrderByDescending("ID")

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルをフィルター
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub FilterButton_Click()

    Dim keywords As Collection
    Set keywords = New Collection

    keywords.Add ("aaa")
    keywords.Add ("ccc")

    Dim tableCls As New TableSample
    Call tableCls.Filter("TEXT", keywords)
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルのフィルターを解除
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub FilterClearButton_Click()

    Dim tableCls As New TableSample
    Call tableCls.FilterClear

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 全レコードのデータ保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SaveAllButton_Click()

    Dim tableCls As New TableSample
    Call tableCls.SaveAll
    
    MsgBox "全レコードのデータを保存しました。"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 1レコードのデータ保存もしくは削除
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Debug.Print "Double Clicked!"
    
    Dim tableCls As New TableSample
    Call tableCls.SaveOrDeleteOne(Target, Cancel)

End Sub

