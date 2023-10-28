Attribute VB_Name = "Repository"
Option Explicit

Dim dao As IDao

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �V�[�g�ɓo�^����Ă���eSELECT�������s
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()

    Dim i As Long

    ' Config����������
    Call InitConfig
                    
    ' DB�ڑ�
    Call OpenConnection(DB_CONNECTION_MODE)
    
    ' SQL�e�[�u�����i�[
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL�����J��Ԃ����s
    Dim sql As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
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
        dao.Query sql
        dao.PasteRecordsetToTable sqlTable
        
        ' ���R�[�h�Z�b�g��ؒf
        dao.CloseRecordset
    
    Next i
        
    ' Oracle�ؒf
    dao.CloseConnection
    Set dao = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �V�[�g�ɓo�^����Ă���SELECT�������s
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSql(name As String)

    Dim i As Long

    ' Config����������
    Call InitConfig
                    
    ' DB�ڑ�
    Call OpenConnection(DB_CONNECTION_MODE)
    
    ' SQL�e�[�u�����i�[
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SQL_SHEET).ListObjects(SQL_TABLE)
    
    ' SQL�����J��Ԃ����s
    Dim sql As String, sqlName As String, sheet As String, sheetTable As String
    Dim sqlTable As ListObject
    For i = 1 To table.ListRows.Count
    
        sqlName = table.ListColumns(SQL_COL_NAME).DataBodyRange(i).Value
        
        ' �w�肵��name��SQL�̂ݎ��s
        If sqlName = name Then
            sql = table.ListColumns(SQL_COL_SQL).DataBodyRange(i).Value
            sheet = table.ListColumns(SQL_COL_SHEET).DataBodyRange(i).Value
            sheetTable = table.ListColumns(SQL_COL_TABLE).DataBodyRange(i).Value
            
            If Not TableExists(sheetTable) Then
                MsgBox "�e�[�u�� " & sheetTable & " �����݂��܂���B�e�[�u�����̂����������m�F���ĉ������B"
                Exit Sub
            End If
            
            Set sqlTable = ThisWorkbook.Worksheets(sheet).ListObjects(sheetTable)
            
            ' �N�G�����s
            dao.Query sql
            dao.PasteRecordsetToTable sqlTable
            
            ' ���R�[�h�Z�b�g��ؒf
            dao.CloseRecordset
            
        End If
    
    Next i
        
    ' Oracle�ؒf
    dao.CloseConnection
    Set dao = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �e�[�u���̑S���R�[�h��ۑ�
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecords(table As ListObject, dbTable As String, _
                       updateColumns As Collection, conditions As Collection)
    
    ' Config����������
    Call InitConfig
                    
    ' Config����������
    Call InitConfig
                    
    ' DB�ڑ�
    Call OpenConnection(DB_CONNECTION_MODE)
    
    Dim i As Long
    
    For i = 1 To table.ListRows.Count
        dao.SaveRecord table, i, dbTable, updateColumns, conditions
    Next

End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �e�[�u����1���R�[�h��ۑ�
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                      updateColumns As Collection, conditions As Collection)
    
    ' Config����������
    Call InitConfig
                    
    ' Config����������
    Call InitConfig
                    
    ' DB�ڑ�
    Call OpenConnection(DB_CONNECTION_MODE)
    
    dao.SaveRecord table, rowIndex, dbTable, updateColumns, conditions

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �e�[�u����1���R�[�h���폜
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub DeleteRecord(table As ListObject, rowIndex As Long, dbTable As String, _
                        conditions As Collection)
    
    ' Config����������
    Call InitConfig
                    
    ' Config����������
    Call InitConfig
                    
    ' DB�ڑ�
    Call OpenConnection(DB_CONNECTION_MODE)
    
    dao.DeleteRecord table, rowIndex, dbTable, conditions

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : DB�ڑ�
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub OpenConnection(connectionMode As String)
                    
    If connectionMode = "OracleOra" Then
    
        ' Oracle�ڑ�
        Set dao = New DaoOracleOra
        dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD
        Exit Sub
    
    End If
                    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : Oracle�ڑ��e�X�g
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOpenConnectionOracleOra()
        
    ' Config����������
    Call InitConfig
                    
    ' Oracle�ڑ�
    Call OpenConnection("OracleOra")
    
    MsgBox "Oracle�ڑ��e�X�g ����"

End Sub
