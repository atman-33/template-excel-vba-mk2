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
    Set table = ThisWorkbook.Worksheets("Sample1").ListObjects("SampleTable_tbl")
    Debug.Print table.name
    
    Debug.Print table.ListColumns("�ۑ�").DataBodyRange.Address
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub �e�[�u�����X�g()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim i As Long
        For i = 1 To ws.ListObjects.Count
            Debug.Print ws.name, i, ws.ListObjects(i).name
        Next i
    Next ws
End Sub
