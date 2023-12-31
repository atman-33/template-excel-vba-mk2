VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dao_OracleOra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDao

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Class: Dao_OracleOra
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Private m_connection As Object    ' Connection
Private m_recordset As Object     ' Recordset

' Oracle接続情報
Private m_oraDataSource As String
Private m_oraUserId As String
Private m_oraPassword As String

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コンストラクタ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 初期化
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Init(dataSource As String, username As String, password As String)
    
    m_oraDataSource = dataSource
    m_oraUserId = username
    m_oraPassword = password
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : オラクルへの接続処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_OpenConnection()

    ' 既にConnectionが作成済みであれば処理を抜ける
    If Not (m_connection Is Nothing) Then
        Exit Sub
    End If
    
    Set m_connection = CreateObject("Adodb.Connection")
    Set m_recordset = CreateObject("Adodb.Recordset")
    
    Dim constr As String
    
    constr = "Provider=" & "OraOLEDB.Oracle" _
                & ";Data Source=" & m_oraDataSource _
                & ";User ID=" & m_oraUserId _
                & ";Password=" & m_oraPassword & ";"
    
    Debug.Print (constr)
    m_connection.ConnectionString = constr
    m_connection.Open
    Debug.Print "オラクルへの接続完了"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : トランザクション開始処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub BeginTrans()

    m_connection.BeginTrans
    Debug.Print "トランザクション開始"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コミット処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub CommitTrans()
    
    m_connection.CommitTrans
    Debug.Print "コミット処理実施"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ロールバック処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub RollbackTrans()
    
    m_connection.RollbackTrans
    Debug.Print "ロールバック処理実施"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : DB切断処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_CloseConnection()

    ' 既にConnectionが破棄済みであれば処理を抜ける
    If m_connection Is Nothing Then
        Exit Sub
    End If

    ' ---- エラー処理の命令取り消し ---- '
    On Error Resume Next

    m_connection.Close
    m_recordset.Close
    
    Set m_connection = Nothing
    Set m_recordset = Nothing
    
    On Error GoTo 0
    ' ---------------------------------- '
    
    Debug.Print "DBへの切断完了"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : レコードセットのクローズ
'           SQL実行でレコードセットにデータが格納された後はクローズ必要
'           （連続でSQLを実行してレコードセットをOpenすることはできない）
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_CloseRecordset()
    m_recordset.Close
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQLの実行（SELECT）
'           SELECT文の実行後は、レコードセットにデータが格納される。
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_Query(sql As String)
        
    Debug.Print sql & " を実行"
    m_recordset.Open sql, m_connection
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQLの実行（INSERT,UPDATE,DELETE）
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Function IDao_Execute(sql As String) As Long

    Dim recordsAffected  As Long
    recordsAffected = 0
    
    Call BeginTrans
    
On Error GoTo ErrorHandler

    Debug.Print sql & " を実行"
    m_connection.Execute sql, recordsAffected
    Call CommitTrans
    
    'Finally:へ飛ぶ
    GoTo Finally
    
'例外処理
ErrorHandler:
    
    Call RollbackTrans
    
    'エラーメッセージを表示する
    Debug.Print "[No:" & Err.Number & "]" & Err.Description
    MsgBox "[No:" & Err.Number & "]" & Err.Description, vbCritical & vbOKOnly, "エラー"
    MsgBox sql
    MsgBox "エラー発生のためマクロを終了します。", vbCritical & vbOKOnly, "エラー"
    End
        
'最終処理
Finally:
        
    Debug.Print "処理レコード数:" & CStr(recordsAffected)
    IDao_Execute = recordsAffected

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : UPDATEを実行し、対象レコードが無ければINSERT
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_Save(insert As String, update As String)
    
    Dim recordsAffected  As Long
    recordsAffected = IDao_Execute(update)
    Debug.Print "UPDATEレコード数:" & CStr(recordsAffected)
    
    If recordsAffected = 0 Then
    
        recordsAffected = IDao_Execute(insert)
        Debug.Print "INSERTレコード数:" & CStr(recordsAffected)
            
    End If

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルのレコードをSave
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                           keyColumns As Collection, updateColumns As Collection)

    Dim insert As String, update As String
    Dim col As Variant, val As String
    
    insert = "INSERT INTO __Table__ (__InsertColumns__) VALUES (__InsertValues__)"
    update = "UPDATE __Table__ SET __UpdateColumns__ WHERE __Conditions__"

    ' INSERT文を生成（INSERT文は、 keyColumns と updateColumns を合わせた文）
    Dim insertColumnsParts As String, insertValuesParts As String
    
    insertColumnsParts = ""
    insertValuesParts = ""
    
    For Each col In keyColumns
        
        ' カラム
        If insertColumnsParts = "" Then
            insertColumnsParts = col
        Else
            insertColumnsParts = insertColumnsParts + "," + col
        End If
        
        ' 値
        val = table.ListColumns(col).DataBodyRange(rowIndex)
        If insertValuesParts = "" Then
            insertValuesParts = "'" + val + "'"
        Else
            insertValuesParts = insertValuesParts + ",'" + val + "'"
        End If
    
    Next
    
    For Each col In updateColumns
        
        ' カラム
        If insertColumnsParts = "" Then
            insertColumnsParts = col
        Else
            insertColumnsParts = insertColumnsParts + "," + col
        End If
        
        ' 値
        val = table.ListColumns(col).DataBodyRange(rowIndex)
        If insertValuesParts = "" Then
            insertValuesParts = "'" + val + "'"
        Else
            insertValuesParts = insertValuesParts + ",'" + val + "'"
        End If
    
    Next
    
    insert = Replace(insert, "__Table__", dbTable)
    insert = Replace(insert, "__InsertColumns__", insertColumnsParts)
    insert = Replace(insert, "__InsertValues__", insertValuesParts)
    ' Debug.Print insert
    
    ' UPDATE文を生成
    Dim updateColumnsParts As String, conditionsParts As String
    updateColumnsParts = ""
    conditionsParts = ""

    For Each col In updateColumns
        
        ' 値
        val = table.ListColumns(col).DataBodyRange(rowIndex)
        
        If updateColumnsParts = "" Then
            updateColumnsParts = col + "=" + "'" + val + "'"
        Else
            updateColumnsParts = updateColumnsParts + "," + col + "=" + "'" + val + "'"
        End If
        
    Next
    
    For Each col In keyColumns
        
        ' 値
        val = table.ListColumns(col).DataBodyRange(rowIndex)
        
        If conditionsParts = "" Then
            conditionsParts = col + "=" + "'" + val + "'"
        Else
            conditionsParts = conditionsParts + " AND " + col + "=" + "'" + val + "'"
        End If
        
    Next

    update = Replace(update, "__Table__", dbTable)
    update = Replace(update, "__UpdateColumns__", updateColumnsParts)
    update = Replace(update, "__Conditions__", conditionsParts)

    ' Debug.Print update
        
    Call IDao_Save(insert, update)
    Debug.Print "レコードSave完了"

End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルのレコードをDelete
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                             conditions As Collection)
                           
    Dim delete As String
    Dim col As Variant, val As String
    
    delete = "DELETE FROM __Table__ WHERE __Conditions__"
                           
    ' DELETE文を生成
    Dim conditionsParts As String
    conditionsParts = ""
                           
    For Each col In conditions
        
        ' 値
        val = table.ListColumns(col).DataBodyRange(rowIndex)
        
        If conditionsParts = "" Then
            conditionsParts = col + "=" + "'" + val + "'"
        Else
            conditionsParts = conditionsParts + " AND " + col + "=" + "'" + val + "'"
        End If
        
    Next

    delete = Replace(delete, "__Table__", dbTable)
    delete = Replace(delete, "__Conditions__", conditionsParts)
    
    Debug.Print delete
    
    Call IDao_Execute(delete)
    Debug.Print "レコードDelete完了"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 実行したSQLで取得したレコードセットをExcelに貼り付け
'           filedsExists => Trueでフィールド名も書き込み
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub PasteRecordset(sheet As Worksheet, startRow As Long, startCol As Long, filedsExists As Boolean)

    Dim i As Long

    If filedsExists = True Then

        ' フィールド名の書き出し
        For i = 0 To m_recordset.Fields.Count - 1
            sheet.Cells(startRow, startCol + i).Value = m_recordset.Fields(i).name
        Next i
        
        startRow = startRow + 1
    
    End If
    
    ' CopyFromRecordsetメソッドで基準セルを指定してデータの書き出し
    sheet.Cells(startRow, startCol).CopyFromRecordset m_recordset
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 実行したSQLで取得したレコードセットをExcelテーブルに貼り付け
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_PasteRecordsetToTable(table As ListObject)

    Dim i As Long
    Dim startRow As Long, startCol As Long

    ' テーブルの全行を削除
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.delete
    End If
    
    startRow = table.Range.Cells(1, 1).Row
    startCol = table.Range.Cells(1, 1).Column

    ' フィールド名の書き出し
    For i = 0 To m_recordset.Fields.Count - 1
        table.Range.Cells(1, 1 + i).Value = m_recordset.Fields(i).name
    Next i
        
    ' テーブルにデータを貼り付け
    table.Range.Cells(2, 1).CopyFromRecordset m_recordset

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle接続テスト
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOracleConnection()
                            
    ' Oracle接続
    Call IDao_OpenConnection
    Call IDao_CloseConnection
    
    MsgBox "Oracle接続テスト 成功"

End Sub
