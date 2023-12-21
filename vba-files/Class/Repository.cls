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
' Class: Repository
'
' - DAOのオープン/クローズは、Repositoryクラスから実施すること
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
Private m_dao As IDao

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 初期化
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Init(dao As IDao)
    
    Set m_dao = dao
        
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 接続オープン
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub OpenConnection()

    Call m_dao.OpenConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 接続クローズ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub CloseConnection()

    Call m_dao.CloseConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : シートに登録されている各SELECT文を実行
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()
                        
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
        Call m_dao.Query(sql)
        Call m_dao.PasteRecordsetToTable(sqlTable)
        
        ' レコードセットを切断
        Call m_dao.CloseRecordset
    
    Next i
        
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : シートに登録されているSELECT文を実行
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSql(tableName As String)
                        
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
            Call m_dao.Query(sql)
            Call m_dao.PasteRecordsetToTable(sqlTable)
            
            ' レコードセットを切断
            Call m_dao.CloseRecordset
            
        End If
    
    Next i
        
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの全レコードを保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecords(table As ListObject, dbTable As String, _
                       keyColumns As Collection, updateColumns As Collection)
                            
    Dim i As Long
    For i = 1 To table.ListRows.Count
        Call m_dao.SaveRecord(table, i, dbTable, keyColumns, updateColumns)
    Next

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの1レコードを保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                       keyColumns As Collection, updateColumns As Collection)
    
    Call m_dao.SaveRecord(table, rowIndex, dbTable, keyColumns, updateColumns)

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルの1レコードを削除
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                        keyColumns As Collection)
    
    Call m_dao.DeleteRecord(table, rowIndex, dbTable, keyColumns)

End Sub


