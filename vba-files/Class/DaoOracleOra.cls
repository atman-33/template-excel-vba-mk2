VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoOracleOra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDao

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' クラス：DaoOracleOra
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Private connection_ As Object    ' Connection
Private recordset_ As Object     ' Recordset

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コンストラクタ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
    
    Set connection_ = CreateObject("Adodb.Connection")
    Set recordset_ = CreateObject("Adodb.Recordset")
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : オラクルへの接続処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_OpenConnection(dataSource As String, username As String, password As String)
    
    Dim constr As String
    
    constr = "Provider=" & "OraOLEDB.Oracle" _
                & ";Data Source=" & dataSource _
                & ";User ID=" & username _
                & ";Password=" & password & ";"
    
    Debug.Print (constr)
    connection_.ConnectionString = constr
    connection_.Open
    Debug.Print "オラクルへの接続完了"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : トランザクション開始処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub BeginTrans()

    connection_.BeginTrans
    Debug.Print "トランザクション開始"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コミット処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub CommitTrans()
    
    connection_.CommitTrans
    Debug.Print "コミット処理実施"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ロールバック処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub RollbackTrans()
    
    connection_.RollbackTrans
    Debug.Print "ロールバック処理実施"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : DB切断処理
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_CloseConnection()

    On Error Resume Next

    connection_.Close
    recordset_.Close
    
    Set connection_ = Nothing
    Set recordset_ = Nothing
    
    On Error GoTo 0     ' エラー処理の命令取り消し
    Debug.Print "DBへの切断完了"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : レコードセットのクローズ
'           SQL実行でレコードセットにデータが格納された後はクローズ必要
'           （連続でSQLを実行してレコードセットをOpenすることはできない）
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_CloseRecordset()
    recordset_.Close
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQLの実行（SELECT）
'           SELECT文の実行後は、レコードセットにデータが格納される。
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_Query(sql As String)
        
    Debug.Print sql & " を実行"
    recordset_.Open sql, connection_
    
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
    connection_.Execute sql, recordsAffected
    Call CommitTrans
    
    'Finally:へ飛ぶ
    GoTo Finally
    
'例外処理
ErrorHandler:
    
    Call RollbackTrans
    
    'エラーメッセージを表示する
    Debug.Print "[No:" & Err.Number & "]" & Err.Description
    MsgBox "[No:" & Err.Number & "]" & Err.Description, vbCritical & vbOKOnly, "エラー"
        
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
        For i = 0 To recordset_.Fields.Count - 1
            sheet.Cells(startRow, startCol + i).Value = recordset_.Fields(i).name
        Next i
        
        startRow = startRow + 1
    
    End If
    
    ' CopyFromRecordsetメソッドで基準セルを指定してデータの書き出し
    sheet.Cells(startRow, startCol).CopyFromRecordset recordset_
    
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
    For i = 0 To recordset_.Fields.Count - 1
        table.Range.Cells(1, 1 + i).Value = recordset_.Fields(i).name
    Next i
        
    ' テーブルにデータを貼り付け
    table.Range.Cells(2, 1).CopyFromRecordset recordset_
        
    ' ---- 下記は、利用しないソースコード ----
        
    ' 下記の利用しない①、②用
'    table.ListRows.Add      ' 空テーブルはエラーになるため1行追加
    
    ' 利用しない①：CopyFromRecordsetはセル書式が変更されるため使わない
'    table.DataBodyRange.CopyFromRecordset recordset_
    
    ' 利用しない②：セル書式は変更されないが、レコード数が多いと処理が遅い
'    Dim x As Long, y As Long
'    y = 2
'    Do Until recordset_.EOF
'        For x = 1 To recordset_.Fields.Count
'            table.Range.Cells(y, x).Value = recordset_.Fields(x - 1).Value
'        Next
'        y = y + 1
'        recordset_.MoveNext
'    Loop

End Sub
