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
' �N���X�FRepository
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- �萔�ݒ� ---- '
Const SQL_SHEET = "SQL"
Const SQL_TABLE = "sqlTable"
Const SQL_COL_NAME = "Name"
Const SQL_COL_SHEET = "Sheet"
Const SQL_COL_TABLE = "Table"
Const SQL_COL_SQL = "SQL"
' ------------------ '

' DAO
Private dao_ As IDao

' DB�ڑ����[�h
Private connectionMode_ As Integer

' Oracle�ڑ����
Private ora_data_source_ As String
Private ora_user_id_ As String
Private ora_password_ As String

' DB�ڑ����[�h
Enum ConnectionModeEnum
    OracleOra = 1
    Access = 2
End Enum

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �������iOracle ver�j
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub InitOracleOra(dataSource As String, username As String, password As String)
    
    Set dao_ = New DaoOracleOra
    
    ora_data_source_ = dataSource
    ora_user_id_ = username
    ora_password_ = password
    
    connectionMode_ = ConnectionModeEnum.OracleOra
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �V�[�g�ɓo�^����Ă���eSELECT�������s
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()
                    
    ' DB�ڑ�
    Call OpenConnection(connectionMode_)
    
    ' SQL�e�[�u�����i�[
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL�����J��Ԃ����s
    Dim sql As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
    
        sql = table.ListColumns(SQL_COL_SQL).DataBodyRange(i).Value
        sheet = table.ListColumns(SQL_COL_SHEET).DataBodyRange(i).Value
        sheetTable = table.ListColumns(SQL_COL_TABLE).DataBodyRange(i).Value
        
        If Not TableExists(sheetTable) Then
            MsgBox "�e�[�u�� " & sheetTable & " �����݂��܂���B�e�[�u�����̂����������m�F���ĉ������B"
            Exit Sub
        End If
        
        Set sqlTable = ThisWorkbook.Worksheets(sheet).ListObjects(sheetTable)
        
        ' �N�G�����s
        Call dao_.Query(sql)
        Call dao_.PasteRecordsetToTable(sqlTable)
        
        ' ���R�[�h�Z�b�g��ؒf
        Call dao_.CloseRecordset
    
    Next i
        
    ' Oracle�ؒf
    Call dao_.CloseConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �V�[�g�ɓo�^����Ă���SELECT�������s
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSql(tableName As String)
                    
    ' DB�ڑ�
    Call OpenConnection(connectionMode_)
    
    ' SQL�e�[�u�����i�[
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL�����J��Ԃ����s
    Dim sql As String, sqlName As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
    
        sqlName = table.ListColumns(SQL_COL_NAME).DataBodyRange(i).Value
        sql = table.ListColumns(SQL_COL_SQL).DataBodyRange(i).Value
        sheet = table.ListColumns(SQL_COL_SHEET).DataBodyRange(i).Value
        sheetTable = table.ListColumns(SQL_COL_TABLE).DataBodyRange(i).Value
        
        ' �w�肵��name��SQL�̂ݎ��s
        If sheetTable = tableName Then
            
            If Not TableExists(sheetTable) Then
                MsgBox "�e�[�u�� " & sheetTable & " �����݂��܂���B�e�[�u�����̂����������m�F���ĉ������B"
                Exit Sub
            End If
            
            Set sqlTable = ThisWorkbook.Worksheets(sheet).ListObjects(sheetTable)
            
            ' �N�G�����s
            Call dao_.Query(sql)
            Call dao_.PasteRecordsetToTable(sqlTable)
            
            ' ���R�[�h�Z�b�g��ؒf
            Call dao_.CloseRecordset
            
        End If
    
    Next i
        
    ' Oracle�ؒf
    Call dao_.CloseConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �e�[�u���̑S���R�[�h��ۑ�
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecords(table As ListObject, dbTable As String, _
                       keyColumns As Collection, updateColumns As Collection)
                        
    ' DB�ڑ�
    Call OpenConnection(connectionMode_)
    
    Dim i As Long
    For i = 1 To table.ListRows.Count
        Call dao_.SaveRecord(table, i, dbTable, keyColumns, updateColumns)
    Next

End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �e�[�u����1���R�[�h��ۑ�
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                       keyColumns As Collection, updateColumns As Collection)
    
    ' DB�ڑ�
    Call OpenConnection(connectionMode_)
    Call dao_.SaveRecord(table, rowIndex, dbTable, keyColumns, updateColumns)

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �e�[�u����1���R�[�h���폜
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                        keyColumns As Collection)
    
    ' DB�ڑ�
    Call OpenConnection(connectionMode_)
    Call dao_.DeleteRecord(table, rowIndex, dbTable, keyColumns)

End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle�ڑ��e�X�g
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOpenConnectionOracleOra()
                            
    ' Oracle�ڑ�
    Call dao_.OpenConnection(ora_data_source_, ora_user_id_, ora_password_)
    
    MsgBox "Oracle�ڑ��e�X�g ����"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : DB�ڑ�
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub OpenConnection(connectionMode As Integer)
                    
    If connectionMode = ConnectionModeEnum.OracleOra Then
        ' Oracle�ڑ�
        Set dao_ = New DaoOracleOra
        Call dao_.OpenConnection(ora_data_source_, ora_user_id_, ora_password_)
        Exit Sub
    
    End If
                    
End Sub
