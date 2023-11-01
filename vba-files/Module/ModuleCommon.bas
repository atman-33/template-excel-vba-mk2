Attribute VB_Name = "ModuleCommon"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' 共通モジュール
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle接続テスト
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOracleConnection()
        
    Dim conf As New Config
    Dim dao As New DaoOracleOra
    
    Call dao.Init(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    Call dao.TestOracleConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQLテーブルのSQLを全て実行
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()
        
    Dim conf As New Config
    
    ' ---- Oracle以外のDBに接続する時は下記のDaoを変更 ---- '
    Dim dao As New DaoOracleOra
    Call dao.Init(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    ' ----------------------------------------------------- '
    
    Dim repo As New Repository
    Call repo.Init(dao)
    Call repo.ExecuteSelectSqls

End Sub

