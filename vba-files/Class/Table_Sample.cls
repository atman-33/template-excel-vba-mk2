VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table_Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' �N���X�FTable_Sample
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' ---- �萔�ݒ� ---- '
Const SHEET_NAME = "Sample1"
Const TABLE_NAME = "SampleTable_tbl"

' DB�e�[�u����
Const DB_TABLE = "SAMPLE_TBL"

' DB CRUD�����p�̃J������`�i�����J�����͔��p�J���}��؂�j
Const DB_KEY_COLS = "ID"
Const DB_UPDATE_COLS = "TEXT,VALUE,CREATED_AT"

' �f�[�^�ۑ�/�폜�{�^����
Const SAVE_COL_NAME = "�ۑ�"
Const DELETE_COL_NAME = "�폜"
' ------------------ '

Private listObject_ As ListObject
Private repository_ As Repository
Private base_ As TableBase

Private dbKeyCols As Collection
Private dbUpdateCols As Collection

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �R���X�g���N�^
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Class_Initialize()
        
    ' Config�ݒ�
    Dim conf As New Config
        
    ' �e�[�u���ݒ�
    Set listObject_ = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    ' �x�[�X�N���X�ݒ�
    Set base_ = New TableBase
    Call base_.Init(listObject_)
    
    ' DAO�ݒ�
    Dim dao As New Dao_OracleOra
    Call dao.Init(conf.Item("ORA_DATA_SOURCE"), conf.Item("ORA_USER_ID"), conf.Item("ORA_PASSWORD"))
    
    ' ���|�W�g���ݒ�
    Set repository_ = New Repository
    Call repository_.Init(dao)
    
    ' DB CRUD�����p�R���N�V�����ݒ�
    Dim varItem As Variant
    
    ' �L�[�J�������i�[�i�z����R���N�V�����ɕϊ��j
    Set dbKeyCols = New Collection
    Dim keyColArray As Variant: keyColArray = Split(DB_KEY_COLS, ",")
        For Each varItem In keyColArray
        Call dbKeyCols.Add(varItem)
    Next

    ' �L�[�łȂ��J�������i�[�i�z����R���N�V�����ɕϊ��j
    Set dbUpdateCols = New Collection
    Dim updateColArray As Variant: updateColArray = Split(DB_UPDATE_COLS, ",")
    For Each varItem In updateColArray
        Call dbUpdateCols.Add(varItem)
    Next
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : ListObject�v���p�e�B
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Get ListObject() As ListObject
    Set ListObject = listObject_
End Property

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : TableBase�v���p�e�B
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Property Get Base() As TableBase
    Set Base = base_
End Property

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �f�[�^�擾
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub LoadData()
    Call repository_.ExecuteSelectSql(TABLE_NAME)
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : �e�[�u���̑S�f�[�^��ۑ�
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveAll()

    Dim varItem As Variant

    Dim keyCols As New Collection
    Dim updateCols As New Collection
            
    Call repository_.SaveRecords(listObject_, DB_TABLE, dbKeyCols, dbUpdateCols)
    
End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Summary : 1���R�[�h�̃f�[�^�ۑ��������͍폜
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Public Sub SaveOrDeleteOne(ByVal Target As Range, Cancel As Boolean)
    
    ' �s�C���f�b�N�X���擾
    Dim indexRow As Long
    indexRow = ActiveCell.Row - listObject_.DataBodyRange(1).Row + 1
    
    ' ���o���s�A�W�v�s�̑I���`�F�b�N
    If indexRow = 0 Or listObject_.ListRows.Count < indexRow Then Exit Sub
    
    Dim keyCols As New Collection
    Dim updateCols As New Collection
    
    ' ---- 1���R�[�h�̃f�[�^��ۑ� ----
    Dim saveRange As Range
    Set saveRange = listObject_.ListColumns(SAVE_COL_NAME).DataBodyRange

    If Not Intersect(Target, saveRange) Is Nothing Then '�Ώ۔͈͂Əd�Ȃ�Ƃ��낪�u�Ȃ��v�ł͂Ȃ��ꍇ
        
        Cancel = True ' �A�N�e�B�u������
            
        Call repository_.SaveRecord(listObject_, indexRow, DB_TABLE, dbKeyCols, dbUpdateCols)
                
        MsgBox "�N���b�N�������R�[�h�̃f�[�^��ۑ����܂����B"
        Exit Sub
        
    End If
    
    ' ---- 1���R�[�h�̃f�[�^���폜 ----
    Dim deleteRange As Range
    Set deleteRange = listObject_.ListColumns(DELETE_COL_NAME).DataBodyRange
    
    If Not Intersect(Target, deleteRange) Is Nothing Then '�Ώ۔͈͂Əd�Ȃ�Ƃ��낪�u�Ȃ��v�ł͂Ȃ��ꍇ
        
        Cancel = True ' �A�N�e�B�u������
        
        ' �_�C�A���O��\��
        Dim response As VbMsgBoxResult
        response = MsgBox("�N���b�N�������R�[�h���폜���܂����H", vbQuestion + vbYesNo, "�폜�̊m�F")
    
        ' ���[�U�[�̑I���ɉ����ď��������s�܂��͏I��
        If response <> vbYes Then
            ' �L�����Z�����ꂽ�ꍇ�̏���
            Exit Sub
        End If
                        
        Call repository_.DeleteRecord(listObject_, indexRow, DB_TABLE, dbKeyCols)
        Call repository_.ExecuteSelectSql(TABLE_NAME)
        
        MsgBox "�N���b�N�������R�[�h�̃f�[�^���폜���܂����B"
        Exit Sub
        
    End If
    
End Sub