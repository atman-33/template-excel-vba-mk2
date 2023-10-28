Attribute VB_Name = "ModuleCommon"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' 共通モジュール
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle接続テスト
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOpenConnectionOracleOra()
        
    Dim conf As New Config
    Dim repo As New Repository
    
    Call repo.InitOracleOra(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    Call repo.TestOpenConnectionOracleOra

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQLテーブルのSQLを全て実行
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()
        
    Dim conf As New Config
    Dim repo As New Repository
    
    Call repo.InitOracleOra(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    Call repo.ExecuteSelectSqls

End Sub

