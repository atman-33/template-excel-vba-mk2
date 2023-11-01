VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' クラス：Repository
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- 定数設定 ---- '
Const SQL_SHEET = "SQL"
Const SQL_TABLE = "SQL_tbl"
Const SQL_COL_NAME = "Name"
Const SQL_COL_SHEET = "Sheet"
Const SQL_COL_TABLE = "Table"
Const SQL_COL_SQL = "SQL"
' ------------------ '

' DAO
Private dao_ As IDao

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 初期化
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Init(dao As IDao)
    
    Set dao_ = dao
        
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : シートに登録されている各SELECT文を実行
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()
                    
    ' DB接続
    Call dao_.OpenConnection
    
    ' SQLテーブルを格納
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL文を繰り返し実行
    Dim sql As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
    
        sql = table.ListColumns(SQL_COL_SQL).DataBodyRange(i).Value
        sheet = table.ListColumns(SQL_COL_SHEET).DataBodyRange(i).Value
        sheetTable = table.ListColumns(SQL_COL_TABLE).DataBodyRange(i).Value
        
        If Not TableExists(sheetTable) Then
            MsgBox "テーブル " & sheetTable & " が存在しません。テーブル名称が正しいか確認して下さい。"
            Exit Sub
        End If
        
        Set sqlTable = ThisWorkbook.Worksheets(sheet).ListObjects(sheetTable)
        
        ' クエリ実行
        Call dao_.Query(sql)
        Call dao_.PasteRecordsetToTable(sqlTable)
        
        ' レコードセットを切断
        Call dao_.CloseRecordset
    
    Next i
        
    ' Oracle切断
    Call dao_.CloseConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : シートに登録されているSELECT文を実行
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSql(tableName As String)
                    
    ' DB接続
    Call dao_.OpenConnection
    
    ' SQLテーブルを格納
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL文を繰り返し実行
    Dim sql As String, sqlName As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
    
        sqlName = table.ListColumns(SQL_COL_NAME).DataBodyRange(i).Value
        sql = table.ListColumns(SQL_COL_SQL).DataBodyRange(i).Value
        sheet = table.ListColumns(SQL_COL_SHEET).DataBodyRange(i).Value
        sheetTable = table.ListColumns(SQL_COL_TABLE).DataBodyRange(i).Value
        
        ' 指定したnameのSQLのみ実行
        If sheetTable = tableName Then
            
            If Not TableExists(sheetTable) Then
                MsgBox "テーブル " & sheetTable & " が存在しません。テーブル名称が正しいか確認して下さい。"
                Exit Sub
            End If
            
            Set sqlTable = ThisWorkbook.Worksheets(sheet).ListObjects(sheetTable)
            
            ' クエリ実行
            Call dao_.Query(sql)
            Call dao_.PasteRecordsetToTable(sqlTable)
            
            ' レコードセットを切断
            Call dao_.CloseRecordset
            
        End If
    
    Next i
        
    ' Oracle切断
    Call dao_.CloseConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの全レコードを保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecords(table As ListObject, dbTable As String, _
                       keyColumns As Collection, updateColumns As Collection)
                        
    ' DB接続
    Call dao_.OpenConnection
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
        Call dao_.SaveRecord(table, i, dbTable, keyColumns, updateColumns)
    Next

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの1レコードを保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                       keyColumns As Collection, updateColumns As Collection)
    
    ' DB接続
    Call dao_.OpenConnection
    Call dao_.SaveRecord(table, rowIndex, dbTable, keyColumns, updateColumns)

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの1レコードを削除
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                        keyColumns As Collection)
    
    ' DB接続
    Call dao_.OpenConnection
    Call dao_.DeleteRecord(table, rowIndex, dbTable, keyColumns)

End Sub


