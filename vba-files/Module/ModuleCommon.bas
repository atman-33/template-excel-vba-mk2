Attribute VB_Name = "ModuleCommon"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' ���ʃ��W���[��
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle�ڑ��e�X�g
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOpenConnectionOracleOra()
        
    Dim conf As New Config
    Dim repo As New Repository
    
    Call repo.InitOracleOra(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    Call repo.TestOpenConnectionOracleOra

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQL�e�[�u����SQL��S�Ď��s
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()
        
    Dim conf As New Config
    Dim repo As New Repository
    
    Call repo.InitOracleOra(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    Call repo.ExecuteSelectSqls

End Sub

