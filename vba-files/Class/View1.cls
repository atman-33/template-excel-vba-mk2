VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "View1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private m_sampleTable As Table_Sample
Private m_repository As Repository

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 初期化
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Init()

    Set m_sampleTable = glb_Factory.CreateSampleTable
    Set m_repository = glb_Factory.CreateRepository

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 各テーブルのデータを更新
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub GetDataButton_Click()

    Call Init
    Call m_repository.OpenConnection
    Call m_sampleTable.LoadData(m_repository)
    Call m_repository.CloseConnection
    
    MsgBox "データを更新しました。"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 昇順に並び替え
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SampleIdOrderByButton_Click()

    Call Init
    Call m_sampleTable.Base.OrderBy("ID")

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 降順に並び替え
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SampleIdOrderByDescendingButton_Click()

    Call Init
    Call m_sampleTable.Base.OrderByDescending("ID")

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルをフィルター
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub FilterButton_Click()

    Call Init

    Dim keywords As Collection
    Set keywords = New Collection

    keywords.Add ("aaa")
    keywords.Add ("ccc")

    Call m_sampleTable.Base.Filter("TEXT", keywords)
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : テーブルのフィルターを解除
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub FilterClearButton_Click()

    Call Init
    Call m_sampleTable.Base.FilterClear

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 全レコードのデータ保存
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SaveAllButton_Click()

    Call Init
    Call m_repository.OpenConnection
    Call m_sampleTable.SaveAll(m_repository)
    Call m_repository.CloseConnection
    
    MsgBox "全レコードのデータを保存しました。"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 1レコードのデータ保存もしくは削除
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Call Init

    Debug.Print "Double Clicked!"
    
    Call m_repository.OpenConnection
    Call m_sampleTable.SaveOrDeleteOne(Target, Cancel, m_repository)
    Call m_repository.CloseConnection

End Sub

