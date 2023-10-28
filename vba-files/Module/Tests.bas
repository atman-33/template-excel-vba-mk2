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
Private Sub Oracleテーブルのデータを取得する()

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
Private Sub OracleテーブルのデータをExcelテーブルに格納する()

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
    Set table = ThisWorkbook.Worksheets("Sample2").ListObjects("sample2_tbl1")
    Debug.Print table.name
    
    Debug.Print table.ListColumns("更新ボタン").DataBodyRange.Address
    
End Sub


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Sub テーブルリスト()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim i As Long
        For i = 1 To ws.ListObjects.Count
            Debug.Print ws.name, i, ws.ListObjects(i).name
        Next i
    Next ws
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub INSERT文のテスト()

    ' Config情報を初期化
    Call InitConfig
                    
    ' Oracle接続
    Dim dao As IDao
    Set dao = New DaoOracleOra
    
    dao.OpenConnection ORA_DATA_SOURCE, ORA_USER_ID, ORA_PASSWORD

    Dim i As Long
    Dim sql As String
    
    sql = "INSERT INTO sample_master(sample_code, sample_code_name) VALUES('007', 'aaa')"

    i = dao.Execute(sql)

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub テーブルのレコードをSave()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets("Sample2").ListObjects("sample2_tbl1")
    
    ' Config情報を初期化
    Call InitConfig
                    
    ' Oracle接続
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
Private Sub テーブルのレコードをDelete()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets("Sample2").ListObjects("sample2_tbl1")
    
    ' Config情報を初期化
    Call InitConfig
                    
    ' Oracle接続
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

