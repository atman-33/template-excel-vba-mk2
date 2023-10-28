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

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : 各テーブルのデータを更新
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub GetDataButton_Click()

    Call ExecuteSelectSqls
    
    MsgBox "データを更新しました。"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : 昇順に並び替え
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SampleIdOrderByButton_Click()

    Dim table As ExListObject
    Set table = New ExListObject
    
    table.ListObject = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    table.OrderBy "SAMPLE_ID"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : 降順に並び替え
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SampleIdOrderByDescendingButton_Click()

    Dim table As ExListObject
    Set table = New ExListObject
    
    table.ListObject = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    table.OrderByDescending "SAMPLE_ID"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : テーブルをフィルター
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub FilterButton_Click()

    Dim table As ExListObject
    Set table = New ExListObject
    
    table.ListObject = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)

    Dim conditions As Collection
    Set conditions = New Collection
    
    conditions.Add ("aaa")
    conditions.Add ("ccc")
    
    table.Filter "SAMPLE_TEXT", conditions
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : テーブルのフィルターを解除
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub FilterClearButton_Click()

    Dim table As ExListObject
    Set table = New ExListObject
    
    table.ListObject = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    table.FilterClear

End Sub
