Attribute VB_Name = "Module_DataAccess"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' �f�[�^�A�N�Z�X���W���[��
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle�ڑ��e�X�g
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOracleConnection()
        
    Dim conf As New Config
    Dim dao As New Dao_OracleOra
    
    Call dao.Init(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    Call dao.TestOracleConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQL�e�[�u����SQL��S�Ď��s
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub ExecuteSelectSqls()
        
    Dim conf As New Config
    
    ' ---- Oracle�ȊO��DB�ɐڑ����鎞�͉��L��Dao��ύX ---- '
    Dim dao As New Dao_OracleOra
    Call dao.Init(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    ' ----------------------------------------------------- '
    
    Dim repo As New Repository
    Call repo.Init(dao)
    Call repo.ExecuteSelectSqls

End Sub

