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
' クラス：TableBase
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

Private listObject_ As ListObject

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 初期化
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub Init(lo As ListObject)
    Set listObject_ = lo
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ListObjectプロパティ
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Let ListObject(ByRef lo As ListObject)
    Set listObject_ = lo
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = listObject_
End Property


' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 指定したカラムに対して、昇順で並び替え
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub OrderBy(colName As String)

    listObject_.Range.Sort key1:=listObject_.ListColumns(colName).Range, order1:=xlAscending, Header:=xlYes
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 指定したカラムに対して、降順で並び替え
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub OrderByDescending(colName As String)

    listObject_.Range.Sort key1:=listObject_.ListColumns(colName).Range, order1:=xlDescending, Header:=xlYes
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 指定したカラムに対して、フィルター
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
' Summary : フィルターをクリア
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub FilterClear()

    ' オートフィルタを解除
    listObject_.ShowAutoFilter = False
    
    ' オートフィルタを設定
    listObject_.ShowAutoFilter = True
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : コレクションを 1 次元配列に変換
'           配列のインデックスは 0 から開始
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
