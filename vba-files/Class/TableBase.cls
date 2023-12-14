VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' �N���X�FTableBase
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Private listObject_ As ListObject

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ������
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Init(lo As ListObject)
    Set listObject_ = lo
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ListObject�v���p�e�B
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Let ListObject(ByRef lo As ListObject)
    Set listObject_ = lo
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = listObject_
End Property


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �w�肵���J�����ɑ΂��āA�����ŕ��ёւ�
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub OrderBy(colName As String)

    listObject_.Range.Sort key1:=listObject_.ListColumns(colName).Range, order1:=xlAscending, Header:=xlYes
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �w�肵���J�����ɑ΂��āA�~���ŕ��ёւ�
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub OrderByDescending(colName As String)

    listObject_.Range.Sort key1:=listObject_.ListColumns(colName).Range, order1:=xlDescending, Header:=xlYes
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �w�肵���J�����ɑ΂��āA�t�B���^�[
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Filter(colName As String, criterias As Collection)

    Dim criteriasArray As Variant
    criteriasArray = CollectionToArray(criterias)


    listObject_.Range.AutoFilter _
        Field:=listObject_.ListColumns(colName).Index, _
        Criteria1:=Array(criteriasArray), _
        Operator:=xlFilterValues
                
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �t�B���^�[���N���A
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub FilterClear()

    ' �I�[�g�t�B���^������
    listObject_.ShowAutoFilter = False
    
    ' �I�[�g�t�B���^��ݒ�
    listObject_.ShowAutoFilter = True
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �R���N�V������ 1 �����z��ɕϊ�
'           �z��̃C���f�b�N�X�� 0 ����J�n
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Function CollectionToArray(ByVal colTarget As Collection) As Variant
    
    Dim vntResult As Variant
    ReDim vntResult(colTarget.Count - 1)
    
    Dim i As Long
    i = LBound(vntResult)
    
    Dim v As Variant
    For Each v In colTarget
        vntResult(i) = v
        i = i + 1
    Next v
    
    CollectionToArray = vntResult
End Function