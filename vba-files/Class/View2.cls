VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "View2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const SHEET_NAME = "Sample2"
Private Const TABLE_NAME = "sample2_tbl1"

Private Const SAVE_COL_NAME = "データ保存"
Private Const DELETE_COL_NAME = "データ削除"

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : 全レコードのデータ保存
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SaveAllButton_Click()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)

    Dim updateColumns As Collection, conditions As Collection

    Set updateColumns = New Collection
    Set conditions = New Collection
            
    updateColumns.Add ("SAMPLE_TEXT")
    updateColumns.Add ("SAMPLE_COMBOBOX_TEXT")
    updateColumns.Add ("SAMPLE_VALUE")
    updateColumns.Add ("SAMPLE_DATE")
    updateColumns.Add ("SAMPLE_FLAG")
    updateColumns.Add ("SAMPLE_CODE")
    
    conditions.Add ("SAMPLE_ID")
    
    Call SaveRecords(table, "sample_table", updateColumns, conditions)
    
    MsgBox "全レコードのデータを保存しました。"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : 1レコードのデータ保存もしくは削除
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    ' Debug.Print table.ListColumns(SAVE_COL_NAME).DataBodyRange.Address
    
    ' 行インデックスを取得
    Dim indexRow As Long
    indexRow = ActiveCell.Row - table.DataBodyRange(1).Row + 1
    
    ' 見出し行、集計行の選択チェック
    If indexRow = 0 Or table.ListRows.Count < indexRow Then Exit Sub
    
    
    Dim updateColumns As Collection, conditions As Collection
    
    ' ---- 1レコードのデータを保存 ----
    Dim saveRange As Range
    Set saveRange = table.ListColumns(SAVE_COL_NAME).DataBodyRange

    If Not Intersect(Target, saveRange) Is Nothing Then '対象範囲と重なるところが「なし」ではない場合
        
        Cancel = True ' アクティブを解除
                
        Set updateColumns = New Collection
        Set conditions = New Collection
                
        updateColumns.Add ("SAMPLE_TEXT")
        updateColumns.Add ("SAMPLE_COMBOBOX_TEXT")
        updateColumns.Add ("SAMPLE_VALUE")
        updateColumns.Add ("SAMPLE_DATE")
        updateColumns.Add ("SAMPLE_FLAG")
        updateColumns.Add ("SAMPLE_CODE")
        
        conditions.Add ("SAMPLE_ID")
        
        Call SaveRecord(table, indexRow, "sample_table", updateColumns, conditions)
        
        MsgBox "クリックしたレコードのデータを保存しました。"
        Exit Sub
        
    End If
    
    ' ---- 1レコードのデータを削除 ----
    Dim deleteRange As Range
    Set deleteRange = table.ListColumns(DELETE_COL_NAME).DataBodyRange
    
    If Not Intersect(Target, deleteRange) Is Nothing Then '対象範囲と重なるところが「なし」ではない場合
        
        Cancel = True ' アクティブを解除
                
        Set conditions = New Collection
                        
        conditions.Add ("SAMPLE_ID")
        
        Call DeleteRecord(table, indexRow, "sample_table", conditions)
        Call ExecuteSelectSql("サンプル2-1")
        
        MsgBox "クリックしたレコードのデータを削除しました。"
        Exit Sub
        
    End If
End Sub
