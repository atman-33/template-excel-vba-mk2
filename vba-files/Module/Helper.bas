Attribute VB_Name = "Helper"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : ��ʕ`��Ȃǂ��~���Ď��s�𑁂�����B
' Note  : Focus = True  -> �`���~�A�C�x���g�}���A�蓮�v�Z
'         Focus = False -> �`��ĊJ�A�C�x���g�Ď��ĊJ�A�����v�Z
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Focus(ByVal flag As Boolean)
    With Application
        .EnableEvents = Not flag
        .ScreenUpdating = Not flag
        .Calculation = IIf(flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �e�[�u�����̂����݂���ꍇ��True
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Function TableExists(tableName As String) As Boolean
    
    TableExists = False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim i As Long
        For i = 1 To ws.ListObjects.Count
            
            If ws.ListObjects(i).name = tableName Then
                TableExists = True
                Exit Function
            End If
        
        Next i
    Next ws
    
End Function
