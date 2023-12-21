Attribute VB_Name = "Module_DataAccessDB"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' �ėpDB�A�N�Z�X���W���[��
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Oracle�ڑ��e�X�g
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestOracleConnection()
        
    Dim dao As New Dao_OracleOra
    
    Call dao.Init(glb_Config.Item("ORA_DATA_SOURCE"), glb_Config.Item("ORA_USER_ID"), glb_Config.Item("ORA_PASSWORD"))
    Call dao.TestOracleConnection

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : SQL�e�[�u����SQL��S�Ď��s
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub TestExecuteSqls()
            
    Dim repo As Repository
    Set repo = glb_Factory.CreateRepository
    
    Call repo.OpenConnection
    Call repo.ExecuteSelectSqls
    Call repo.CloseConnection

End Sub

