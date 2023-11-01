Attribute VB_Name = "Tests"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub コンフィグクラスをテスト()
    Dim conf As Config
    Set conf = New Config
    
    Debug.Print conf.Item("ORA_DATA_SOURCE")
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub ListObjectの開始セルを取得する()
    
    Dim rng As Range
    Set rng = ActiveSheet.ListObjects(1).Range.Cells(1, 1)
    
    ' 行列
    Debug.Print rng.Address(False, False)
    
    ' 行
    Debug.Print rng.Row
    
    ' 列
    Debug.Print rng.Column
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub ListObjectの列範囲を取得する()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets("Sample1").ListObjects("SampleTable_tbl")
    Debug.Print table.name
    
    Debug.Print table.ListColumns("保存").DataBodyRange.Address
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub テーブルリスト()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim i As Long
        For i = 1 To ws.ListObjects.Count
            Debug.Print ws.name, i, ws.ListObjects(i).name
        Next i
    Next ws
End Sub
