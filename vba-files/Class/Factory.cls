VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Class: Factory
'
' - VBAでは引数有りのインスタンス生成ができないため、Factory経由でインスタンス生成
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コンストラクタ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
    
    Set glb_Dao = New Dao_OracleOra
    Call glb_Dao.Init(glb_Config.Item("ORA_DATA_SOURCE"), glb_Config.Item("ORA_USER_ID"), glb_Config.Item("ORA_PASSWORD"))
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Repositoryを生成
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Function CreateRepository() As Repository
    
    Set CreateRepository = New Repository
    Call CreateRepository.Init(glb_Dao)

End Function

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : Table_Sampleを生成
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Function CreateSampleTable() As Table_Sample
    
    Set CreateSampleTable = New Table_Sample

End Function


