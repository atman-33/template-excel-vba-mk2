Attribute VB_Name = "Module_Shared"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' ���ʂ̃p�u���b�N�I�u�W�F�N�g�p���W���[��
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' Config�f�[�^�i�[
Public glb_Config As New Config

' DAO�iOracle�ȊO��DB�ɐڑ����鎞�͉��L��Dao�N���X��ύX�j
Public glb_Dao As New Dao_OracleOra

' Class�����pFactory
Public glb_Factory As New Factory
