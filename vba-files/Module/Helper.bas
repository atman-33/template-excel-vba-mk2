Attribute VB_Name = "Helper"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : 画面描画などを停止して実行を早くする。
' Note  : Focus = True  -> 描画停止、イベント抑制、手動計算
'         Focus = False -> 描画再開、イベント監視再開、自動計算
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Focus(ByVal flag As Boolean)
    With Application
        .EnableEvents = Not flag
        .ScreenUpdating = Not flag
        .Calculation = IIf(flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : テーブル名称が存在する場合はTrue
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
