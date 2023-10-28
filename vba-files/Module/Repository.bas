Attribute VB_Name = "Repository"
Option Explicit

Dim dao As IDao

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : シートに登録されている各SELECT文を実行
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()

    Dim i As Long

    ' Config情報を初期化
    Call InitConfig
                    
    ' DB接続
    Call OpenConnection(DB_CONNECTION_MODE)
    
    ' SQLテーブルを格納
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL文を繰り返し実行
    Dim sql As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
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
        dao.Query sql
        dao.PasteRecordsetToTable sqlTable
        
        ' レコードセットを切断
        dao.CloseRecordset
    
    Next i
        
    ' Oracle切断
    dao.CloseConnection
    Set dao = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : シートに登録されているSELECT文を実行
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSql(name As String)

    Dim i As Long

    ' Config情報を初期化
    Call InitConfig
                    
    ' DB接続
    Call OpenConnection(DB_CONNECTION_MODE)
    
    ' SQLテーブルを格納
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL文を繰り返し実行
    Dim sql As String, sqlName As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
    For i = 1 To table.ListRows.Count
    
        sqlName = table.ListColumns(SQL_COL_NAME).DataBodyRange(i).Value
        
        ' 指定したnameのSQLのみ実行
        If sqlName = name Then
            sql = table.ListColumns(SQL_COL_SQL).DataBodyRange(i).Value
            sheet = table.ListColumns(SQL_COL_SHEET).DataBodyRange(i).Value
            sheetTable = table.ListColumns(SQL_COL_TABLE).DataBodyRange(i).Value
            
            If Not TableExists(sheetTable) Then
                MsgBox "テーブル " & sheetTable & " が存在しません。テーブル名称が正しいか確認して下さい。"
                Exit Sub
            End If
            
            Set sqlTable = ThisWorkbook.Worksheets(sheet).ListObjects(sheetTable)
            
            ' クエリ実行
            dao.Query sql
            dao.PasteRecordsetToTable sqlTable
            
            ' レコードセットを切断
            dao.CloseRecordset
            
        End If
    
    Next i
        
    ' Oracle切断
    dao.CloseConnection
    Set dao = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : テーブルの全レコードを保存
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecords(table As ListObject, dbTable As String, _
                       updateColumns As Collection, conditions As Collection)
    
    ' Config情報を初期化
    Call InitConfig
                    
    ' Config情報を初期化
    Call InitConfig
                    
    ' DB接続
    Call OpenConnection(DB_CONNECTION_MODE)
    
    Dim i As Long
    
    For i = 1 To table.ListRows.Count
        dao.SaveRecord table, i, dbTable, updateColumns, conditions
    Next

End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : テーブルの1レコードを保存
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                      updateColumns As Collection, conditions As Collection)
    
    ' Config情報を初期化
    Call InitConfig
                    
    ' Config情報を初期化
    Call InitConfig
                    
    ' DB接続
    Call OpenConnection(DB_CONNECTION_MODE)
    
    dao.SaveRecord table, rowIndex, dbTable, updateColumns, conditions

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : テーブルの1レコードを削除
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                        conditions As Collection)
    
    ' Config情報を初期化
    Call InitConfig
                    
    ' Config情報を初期化
    Call InitConfig
                    
    ' DB接続
    Call OpenConnection(DB_CONNECTION_MODE)
    
    dao.DeleteRecord table, rowIndex, dbTable, conditions

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : DB接続
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub OpenConnection(connectionMode As String)
                    
    If connectionMode = "OracleOra" Then
    
        ' Oracle接続
        Set dao = New DaoOracleOra
        dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD
        Exit Sub
    
    End If
                    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : Oracle接続テスト
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOpenConnectionOracleOra()
        
    ' Config情報を初期化
    Call InitConfig
                    
    ' Oracle接続
    Call OpenConnection("OracleOra")
    
    MsgBox "Oracle接続テスト 成功"

End Sub
