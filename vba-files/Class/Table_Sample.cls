VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table_Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' クラス：Table_Sample
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- 定数設定 ---- '
Const SHEET_NAME = "Sample1"
Const TABLE_NAME = "SampleTable_tbl"

' DBテーブル名
Const DB_TABLE = "SAMPLE_TBL"

' DB CRUD処理用のカラム定義（複数カラムは半角カンマ区切り）
Const DB_KEY_COLS = "ID"
Const DB_UPDATE_COLS = "TEXT,VALUE,CREATED_AT"

' データ保存/削除ボタン列
Const SAVE_COL_NAME = "保存"
Const DELETE_COL_NAME = "削除"
' ------------------ '

Private m_listObject As ListObject
Private m_repository As Repository
Private m_base As TableBase

Private m_dbKeyCols As Collection
Private m_dbUpdateCols As Collection

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コンストラクタ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
        
    ' Config設定
    Dim conf As New Config
        
    ' テーブル設定
    Set m_listObject = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    ' ベースクラス設定
    Set m_base = New TableBase
    Call m_base.Init(m_listObject)
    
    ' DAO設定
    Dim dao As New Dao_OracleOra
    Call dao.Init(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    
    ' リポジトリ設定
    Set m_repository = New Repository
    Call m_repository.Init(dao)
    
    ' DB CRUD処理用コレクション設定
    Dim varItem As Variant
    
    ' キーカラムを格納（配列をコレクションに変換）
    Set m_dbKeyCols = New Collection
    Dim keyColArray As Variant: keyColArray = Split(DB_KEY_COLS, ",")
        For Each varItem In keyColArray
        Call m_dbKeyCols.Add(varItem)
    Next

    ' キーでないカラムを格納（配列をコレクションに変換）
    Set m_dbUpdateCols = New Collection
    Dim updateColArray As Variant: updateColArray = Split(DB_UPDATE_COLS, ",")
    For Each varItem In updateColArray
        Call m_dbUpdateCols.Add(varItem)
    Next
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ListObjectプロパティ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Get ListObject() As ListObject
    Set ListObject = m_listObject
End Property

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : TableBaseプロパティ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Get Base() As TableBase
    Set Base = m_base
End Property

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : データ取得
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub LoadData()
    Call m_repository.ExecuteSelectSql(TABLE_NAME)
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの全データを保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveAll()

    Dim varItem As Variant

    Dim keyCols As New Collection
    Dim updateCols As New Collection
            
    Call m_repository.SaveRecords(m_listObject, DB_TABLE, m_dbKeyCols, m_dbUpdateCols)
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 1レコードのデータ保存もしくは削除
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveOrDeleteOne(ByVal Target As Range, Cancel As Boolean)
    
    ' 行インデックスを取得
    Dim indexRow As Long
    indexRow = ActiveCell.Row - m_listObject.DataBodyRange(1).Row + 1
    
    ' 見出し行、集計行の選択チェック
    If indexRow = 0 Or m_listObject.ListRows.Count < indexRow Then Exit Sub
    
    Dim keyCols As New Collection
    Dim updateCols As New Collection
    
    ' ---- 1レコードのデータを保存 ----
    Dim saveRange As Range
    Set saveRange = m_listObject.ListColumns(SAVE_COL_NAME).DataBodyRange

    If Not Intersect(Target, saveRange) Is Nothing Then '対象範囲と重なるところが「なし」ではない場合
        
        Cancel = True ' アクティブを解除
            
        Call m_repository.SaveRecord(m_listObject, indexRow, DB_TABLE, m_dbKeyCols, m_dbUpdateCols)
                
        MsgBox "クリックしたレコードのデータを保存しました。"
        Exit Sub
        
    End If
    
    ' ---- 1レコードのデータを削除 ----
    Dim deleteRange As Range
    Set deleteRange = m_listObject.ListColumns(DELETE_COL_NAME).DataBodyRange
    
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
                        
        Call m_repository.DeleteRecord(m_listObject, indexRow, DB_TABLE, m_dbKeyCols)
        Call m_repository.ExecuteSelectSql(TABLE_NAME)
        
        MsgBox "クリックしたレコードのデータを削除しました。"
        Exit Sub
        
    End If
    
End Sub
