VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' クラス：TableSample
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- 定数設定 ---- '
Const SHEET_NAME = "Sample1"
Const TABLE_NAME = "SampleTable_tbl"

Const COL_ID = "ID"
Const COL_TEXT = "TEXT"
Const COL_VALUE = "VALUE"
Const COL_CREATED_AT = "CREATED_AT"

' DBテーブル名
Const DB_TABLE = "SAMPLE_TBL"

' DB CRUD処理用のカラム定義（複数カラムは半角カンマ区切り）
Const DB_KEY_COLS = "ID"
Const DB_UPDATE_COLS = "TEXT,VALUE,CREATED_AT"

' データ保存/削除ボタン列
Const SAVE_COL_NAME = "保存"
Const DELETE_COL_NAME = "削除"
' ------------------ '

Private listObject_ As ListObject
Private repository_ As Repository

Private dbKeyCols As Collection
Private dbUpdateCols As Collection

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コンストラクタ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
        
    ' Config設定
    Dim conf As New Config
        
    ' テーブル設定
    Set listObject_ = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    ' DAO設定
    Dim dao As New DaoOracleOra
    Call dao.Init(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    
    ' リポジトリ設定
    Set repository_ = New Repository
    Call repository_.Init(dao)
    
    ' DB CRUD処理用コレクション設定
    Dim varItem As Variant
    
    ' キーカラムを格納（配列をコレクションに変換）
    Set dbKeyCols = New Collection
    Dim keyColArray As Variant: keyColArray = Split(DB_KEY_COLS, ",")
        For Each varItem In keyColArray
        Call dbKeyCols.Add(varItem)
    Next

    ' キーでないカラムを格納（配列をコレクションに変換）
    Set dbUpdateCols = New Collection
    Dim updateColArray As Variant: updateColArray = Split(DB_UPDATE_COLS, ",")
    For Each varItem In updateColArray
        Call dbUpdateCols.Add(varItem)
    Next
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : データ取得
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub LoadData()
    Call repository_.ExecuteSelectSql(TABLE_NAME)
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 昇順にソート
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub OrderBy(varColumn As String)
    
    Dim exTable As New ExListObject
    exTable.ListObject = listObject_
    
    Call exTable.OrderBy(varColumn)

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 降順にソート
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub OrderByDescending(varColumn As String)
    
    Dim exTable As New ExListObject
    exTable.ListObject = listObject_
    
    Call exTable.OrderByDescending(varColumn)

End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルをフィルター
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Filter(varColumn As String, keywords As Collection)
    
    Dim exTable As New ExListObject
    exTable.ListObject = listObject_
    
    Call exTable.Filter(varColumn, keywords)
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルのフィルターを解除
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub FilterClear()
    
    Dim exTable As New ExListObject
    exTable.ListObject = listObject_
        
    Call exTable.FilterClear
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの全データを保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveAll()

    Dim varItem As Variant

    Dim keyCols As New Collection
    Dim updateCols As New Collection
            
    Call repository_.SaveRecords(listObject_, DB_TABLE, dbKeyCols, dbUpdateCols)
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 1レコードのデータ保存もしくは削除
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveOrDeleteOne(ByVal Target As Range, Cancel As Boolean)
    
    ' 行インデックスを取得
    Dim indexRow As Long
    indexRow = ActiveCell.Row - listObject_.DataBodyRange(1).Row + 1
    
    ' 見出し行、集計行の選択チェック
    If indexRow = 0 Or listObject_.ListRows.Count < indexRow Then Exit Sub
    
    Dim keyCols As New Collection
    Dim updateCols As New Collection
    
    ' ---- 1レコードのデータを保存 ----
    Dim saveRange As Range
    Set saveRange = listObject_.ListColumns(SAVE_COL_NAME).DataBodyRange

    If Not Intersect(Target, saveRange) Is Nothing Then '対象範囲と重なるところが「なし」ではない場合
        
        Cancel = True ' アクティブを解除
            
        Call repository_.SaveRecord(listObject_, indexRow, DB_TABLE, dbKeyCols, dbUpdateCols)
                
        MsgBox "クリックしたレコードのデータを保存しました。"
        Exit Sub
        
    End If
    
    ' ---- 1レコードのデータを削除 ----
    Dim deleteRange As Range
    Set deleteRange = listObject_.ListColumns(DELETE_COL_NAME).DataBodyRange
    
    If Not Intersect(Target, deleteRange) Is Nothing Then '対象範囲と重なるところが「なし」ではない場合
        
        Cancel = True ' アクティブを解除
        

        ' ダイアログを表示
        Dim response As VbMsgBoxResult
        response = MsgBox("クリックしたレコードを削除しますか？", vbQuestion + vbYesNo, "削除の確認")
    
        ' ユーザーの選択に応じて処理を実行または終了
        If response <> vbYes Then
            ' キャンセルされた場合の処理
            Exit Sub
        End If
                        
        Call repository_.DeleteRecord(listObject_, indexRow, DB_TABLE, dbKeyCols)
        Call repository_.ExecuteSelectSql(TABLE_NAME)
        
        MsgBox "クリックしたレコードのデータを削除しました。"
        Exit Sub
        
    End If
    
End Sub
