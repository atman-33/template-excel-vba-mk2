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
' �N���X�FDaoOracleOra
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Private connection_ As Object    ' Connection
Private recordset_ As Object     ' Recordset

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �R���X�g���N�^
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
    
    Set connection_ = CreateObject("Adodb.Connection")
    Set recordset_ = CreateObject("Adodb.Recordset")
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �I���N���ւ̐ڑ�����
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_OpenConnection(servicename As String, username As String, password As String)
    
    Dim constr As String
    
    constr = "Provider=" & "OraOLEDB.Oracle" _
                & ";Data Source=" & servicename _
                & ";User ID=" & username _
                & ";Password=" & password & ";"
    
    Debug.Print (constr)
    connection_.ConnectionString = constr
    connection_.Open
    Debug.Print "�I���N���ւ̐ڑ�����"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �g�����U�N�V�����J�n����
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub BeginTrans()

    connection_.BeginTrans
    Debug.Print "�g�����U�N�V�����J�n"
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �R�~�b�g����
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub CommitTrans()
    
    connection_.CommitTrans
    Debug.Print "�R�~�b�g�������{"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : ���[���o�b�N����
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub RollbackTrans()
    
    connection_.RollbackTrans
    Debug.Print "���[���o�b�N�������{"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : DB�ؒf����
' Note  :
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
' Brief : ���R�[�h�Z�b�g�̃N���[�Y
' Note  : SQL���s�Ń��R�[�h�Z�b�g�Ƀf�[�^���i�[���ꂽ��̓N���[�Y�K�v
'         �i�A����SQL�����s���ă��R�[�h�Z�b�g��Open���邱�Ƃ͂ł��Ȃ�
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_CloseRecordset()
    recordset_.Close
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : SQL�̎��s�iSELECT�j
' Note  : SELECT���̎��s��́A���R�[�h�Z�b�g�Ƀf�[�^���i�[�����B
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_Query(sql As String)
        
    Debug.Print sql & " �����s"
    recordset_.Open sql, connection_
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : SQL�̎��s�iINSERT,UPDATE,DELETE�j
' Note  :
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
        
    MsgBox "�G���[�����̂��߃}�N�����I�����܂��B", vbCritical & vbOKOnly, "�G���["
    End
        
'�ŏI����
Finally:
        
    Debug.Print "�������R�[�h��:" & CStr(recordsAffected)
    IDao_Execute = recordsAffected

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : UPDATE�����s���A�Ώۃ��R�[�h���������INSERT
' Note  :
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
' Brief : �e�[�u���̃��R�[�h��Save
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub IDao_SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                           updateColumns As Collection, conditions As Collection)

    Dim insert As String, update As String
    Dim col As Variant, val As String
    
    insert = "INSERT INTO __Table__ (__InsertColumns__) VALUES (__InsertValues__)"
    update = "UPDATE __Table__ SET __UpdateColumns__ WHERE __Conditions__"

    ' INSERT���𐶐�
    Dim insertColumnsParts As String, insertValuesParts As String
    
    insertColumnsParts = ""
    insertValuesParts = ""
    
    For Each col In conditions
        
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
    
    For Each col In conditions
        
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
' Brief : �e�[�u���̃��R�[�h��Delete
' Note  :
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
' Brief : ���s����SQL�Ŏ擾�������R�[�h�Z�b�g��Excel�ɓ\��t��
' Note  : filedsExists => True�Ńt�B�[���h������������
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
' Brief : ���s����SQL�Ŏ擾�������R�[�h�Z�b�g��Excel�e�[�u���ɓ\��t��
' Note  :
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