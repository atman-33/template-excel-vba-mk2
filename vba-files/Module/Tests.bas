Attribute VB_Name = "Tests"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub �R���t�B�O�N���X���e�X�g()
    Dim conf As Config
    Set conf = New Config
    
    Debug.Print conf.Item("ORA_DATA_SOURCE")
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Oracle�e�[�u���̃f�[�^���擾����()

    Call InitConfig
        
    Dim dao As DaoOracleOra
    Set dao = New DaoOracleOra
    
    dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD
    
    dao.Query "select * from sample_master"
    dao.PasteRecordset ThisWorkbook.Worksheets("Sample1"), 1, 1, True
        
    dao.CloseRecordset
    dao.CloseConnection
    
    Set dao = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Oracle�e�[�u���̃f�[�^��Excel�e�[�u���Ɋi�[����()

    Call InitConfig
        
    Dim dao As DaoOracleOra
    Set dao = New DaoOracleOra
    
    dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD
    
    dao.Query "select * from sample_master"
    dao.PasteRecordsetToTable ThisWorkbook.Worksheets("Sample1").ListObjects("sample1_tbl1")
    
    dao.CloseRecordset
    dao.CloseConnection
    
    Set dao = Nothing

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub ListObject�̊J�n�Z�����擾����()
    
    Dim rng As Range
    Set rng = ActiveSheet.ListObjects(1).Range.Cells(1, 1)
    
    ' �s��
    Debug.Print rng.Address(False, False)
    
    ' �s
    Debug.Print rng.Row
    
    ' ��
    Debug.Print rng.Column
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub ListObject�̗�͈͂��擾����()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets("Sample2").ListObjects("sample2_tbl1")
    Debug.Print table.name
    
    Debug.Print table.ListColumns("�X�V�{�^��").DataBodyRange.Address
    
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub �e�[�u�����X�g()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim i As Long
        For i = 1 To ws.ListObjects.Count
            Debug.Print ws.name, i, ws.ListObjects(i).name
        Next i
    Next ws
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub INSERT���̃e�X�g()

    ' Config����������
    Call InitConfig
                    
    ' Oracle�ڑ�
    Dim dao As IDao
    Set dao = New DaoOracleOra
    
    dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD

    Dim i As Long
    Dim sql As String
    
    sql = "INSERT INTO sample_master(sample_code, sample_code_name) VALUES('007', 'aaa')"

    i = dao.Execute(sql)

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub �e�[�u���̃��R�[�h��Save()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets("Sample2").ListObjects("sample2_tbl1")
    
    ' Config����������
    Call InitConfig
                    
    ' Oracle�ڑ�
    Dim dao As IDao
    Set dao = New DaoOracleOra
    
    dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD

    Dim i As Long
    Dim sql As String
    
    Dim updateColumns As Collection, conditions As Collection
    Set updateColumns = New Collection
    Set conditions = New Collection
    
    updateColumns.Add ("SAMPLE_TEXT")
    updateColumns.Add ("SAMPLE_VALUE")
    
    conditions.Add ("SAMPLE_ID")
    conditions.Add ("SAMPLE_CODE")
    
    dao.SaveRecord table, 1, "sample_table", updateColumns, conditions

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub �e�[�u���̃��R�[�h��Delete()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets("Sample2").ListObjects("sample2_tbl1")
    
    ' Config����������
    Call InitConfig
                    
    ' Oracle�ڑ�
    Dim dao As IDao
    Set dao = New DaoOracleOra
    
    dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD

    Dim i As Long
    Dim sql As String
    
    Dim conditions As Collection
    Set conditions = New Collection
    
    conditions.Add ("SAMPLE_ID")
    conditions.Add ("SAMPLE_VALUE")
    
    dao.DeleteRecord table, 1, "sample_table", conditions

End Sub

