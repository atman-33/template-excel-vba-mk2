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
' �N���X�FDao_OracleOra
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Private connection_ As Object    ' Connection
Private recordset_ As Object     ' Recordset

' Oracle�ڑ����
Private ora_data_source_ As String
Private ora_user_id_ As String
Private ora_password_ As String

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �R���X�g���N�^
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
    
    Set connection_ = CreateObject("Adodb.Connection")
    Set recordset_ = CreateObject("Adodb.Recordset")
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ������
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Init(dataSource As String, username As String, password As String)
    
    ora_data_source_ = dataSource
    ora_user_id_ = username
    ora_password_ = password
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �I���N���ւ̐ڑ�����
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_OpenConnection()
    
    Dim constr As String
    
    constr = "Provider=" & "OraOLEDB.Oracle" _
                & ";Data Source=" & ora_data_source_ _
                & ";User ID=" & ora_user_id_ _
                & ";Password=" & ora_password_ & ";"
    
    Debug.Print (constr)
    connection_.ConnectionString = constr
    connection_.Open
    Debug.Print "�I���N���ւ̐ڑ�����"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �g�����U�N�V�����J�n����
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub BeginTrans()

    connection_.BeginTrans
    Debug.Print "�g�����U�N�V�����J�n"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �R�~�b�g����
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub CommitTrans()
    
    connection_.CommitTrans
    Debug.Print "�R�~�b�g�������{"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ���[���o�b�N����
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub RollbackTrans()
    
    connection_.RollbackTrans
    Debug.Print "���[���o�b�N�������{"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : DB�ؒf����
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_CloseConnection()

    On Error Resume Next

    connection_.Close
    recordset_.Close
    
    Set connection_ = Nothing
    Set recordset_ = Nothing
    
    On Error GoTo 0     ' �G���[�����̖��ߎ�����
    Debug.Print "DB�ւ̐ؒf����"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ���R�[�h�Z�b�g�̃N���[�Y
'           SQL���s�Ń��R�[�h�Z�b�g�Ƀf�[�^���i�[���ꂽ��̓N���[�Y�K�v
'           �i�A����SQL�����s���ă��R�[�h�Z�b�g��Open���邱�Ƃ͂ł��Ȃ��j
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_CloseRecordset()
    recordset_.Close
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQL�̎��s�iSELECT�j
'           SELECT���̎��s��́A���R�[�h�Z�b�g�Ƀf�[�^���i�[�����B
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_Query(sql As String)
        
    Debug.Print sql & " �����s"
    recordset_.Open sql, connection_
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQL�̎��s�iINSERT,UPDATE,DELETE�j
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Function IDao_Execute(sql As String) As Long

    Dim recordsAffected  As Long
    recordsAffected = 0
    
    Call BeginTrans
    
On Error GoTo ErrorHandler

    Debug.Print sql & " �����s"
    connection_.Execute sql, recordsAffected
    Call CommitTrans
    
    'Finally:�֔��
    GoTo Finally
    
'��O����
ErrorHandler:
    
    Call RollbackTrans
    
    '�G���[���b�Z�[�W��\������
    Debug.Print "[No:" & Err.Number & "]" & Err.Description
    MsgBox "[No:" & Err.Number & "]" & Err.Description, vbCritical & vbOKOnly, "�G���["
    MsgBox sql
    MsgBox "�G���[�����̂��߃}�N�����I�����܂��B", vbCritical & vbOKOnly, "�G���["
    End
        
'�ŏI����
Finally:
        
    Debug.Print "�������R�[�h��:" & CStr(recordsAffected)
    IDao_Execute = recordsAffected

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : UPDATE�����s���A�Ώۃ��R�[�h���������INSERT
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_Save(insert As String, update As String)
    
    Dim recordsAffected  As Long
    recordsAffected = IDao_Execute(update)
    Debug.Print "UPDATE���R�[�h��:" & CStr(recordsAffected)
    
    If recordsAffected = 0 Then
    
        recordsAffected = IDao_Execute(insert)
        Debug.Print "INSERT���R�[�h��:" & CStr(recordsAffected)
            
    End If

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �e�[�u���̃��R�[�h��Save
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                           keyColumns As Collection, updateColumns As Collection)

    Dim insert As String, update As String
    Dim col As Variant, val As String
    
    insert = "INSERT INTO __Table__ (__InsertColumns__) VALUES (__InsertValues__)"
    update = "UPDATE __Table__ SET __UpdateColumns__ WHERE __Conditions__"

    ' INSERT���𐶐��iINSERT���́A keyColumns �� updateColumns �����킹�����j
    Dim insertColumnsParts As String, insertValuesParts As String
    
    insertColumnsParts = ""
    insertValuesParts = ""
    
    For Each col In keyColumns
        
        ' �J����
        If insertColumnsParts = "" Then
            insertColumnsParts = col
        Else
            insertColumnsParts = insertColumnsParts + "," + col
        End If
        
        ' �l
        val = table.ListColumns(col).DataBodyRange(rowIndex)
        If insertValuesParts = "" Then
            insertValuesParts = "'" + val + "'"
        Else
            insertValuesParts = insertValuesParts + ",'" + val + "'"
        End If
    
    Next
    
    For Each col In updateColumns
        
        ' �J����
        If insertColumnsParts = "" Then
            insertColumnsParts = col
        Else
            insertColumnsParts = insertColumnsParts + "," + col
        End If
        
        ' �l
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
    
    ' UPDATE���𐶐�
    Dim updateColumnsParts As String, conditionsParts As String
    updateColumnsParts = ""
    conditionsParts = ""

    For Each col In updateColumns
        
        ' �l
        val = table.ListColumns(col).DataBodyRange(rowIndex)
        
        If updateColumnsParts = "" Then
            updateColumnsParts = col + "=" + "'" + val + "'"
        Else
            updateColumnsParts = updateColumnsParts + "," + col + "=" + "'" + val + "'"
        End If
        
    Next
    
    For Each col In keyColumns
        
        ' �l
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
    Debug.Print "���R�[�hSave����"

End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �e�[�u���̃��R�[�h��Delete
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                             conditions As Collection)
                           
    Dim delete As String
    Dim col As Variant, val As String
    
    delete = "DELETE FROM __Table__ WHERE __Conditions__"
                           
    ' DELETE���𐶐�
    Dim conditionsParts As String
    conditionsParts = ""
                           
    For Each col In conditions
        
        ' �l
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
    Debug.Print "���R�[�hDelete����"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ���s����SQL�Ŏ擾�������R�[�h�Z�b�g��Excel�ɓ\��t��
'           filedsExists => True�Ńt�B�[���h������������
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub PasteRecordset(sheet As Worksheet, startRow As Long, startCol As Long, filedsExists As Boolean)

    Dim i As Long

    If filedsExists = True Then

        ' �t�B�[���h���̏����o��
        For i = 0 To recordset_.Fields.Count - 1
            sheet.Cells(startRow, startCol + i).Value = recordset_.Fields(i).name
        Next i
        
        startRow = startRow + 1
    
    End If
    
    ' CopyFromRecordset���\�b�h�Ŋ�Z�����w�肵�ăf�[�^�̏����o��
    sheet.Cells(startRow, startCol).CopyFromRecordset recordset_
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ���s����SQL�Ŏ擾�������R�[�h�Z�b�g��Excel�e�[�u���ɓ\��t��
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_PasteRecordsetToTable(table As ListObject)

    Dim i As Long
    Dim startRow As Long, startCol As Long

    ' �e�[�u���̑S�s���폜
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.delete
    End If
    
    startRow = table.Range.Cells(1, 1).Row
    startCol = table.Range.Cells(1, 1).Column

    ' �t�B�[���h���̏����o��
    For i = 0 To recordset_.Fields.Count - 1
        table.Range.Cells(1, 1 + i).Value = recordset_.Fields(i).name
    Next i
        
    ' �e�[�u���Ƀf�[�^��\��t��
    table.Range.Cells(2, 1).CopyFromRecordset recordset_
        
    ' ---- ���L�́A���p���Ȃ��\�[�X�R�[�h ----
        
    ' ���L�̗��p���Ȃ��@�A�A�p
'    table.ListRows.Add      ' ��e�[�u���̓G���[�ɂȂ邽��1�s�ǉ�
    
    ' ���p���Ȃ��@�FCopyFromRecordset�̓Z���������ύX����邽�ߎg��Ȃ�
'    table.DataBodyRange.CopyFromRecordset recordset_
    
    ' ���p���Ȃ��A�F�Z�������͕ύX����Ȃ����A���R�[�h���������Ə������x��
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

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle�ڑ��e�X�g
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOracleConnection()
                            
    ' Oracle�ڑ�
    Call IDao_OpenConnection
    
    MsgBox "Oracle�ڑ��e�X�g ����"

End Sub