VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "View2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const SHEET_NAME = "Sample2"
Private Const TABLE_NAME = "sample2_tbl1"

Private Const SAVE_COL_NAME = "�f�[�^�ۑ�"
Private Const DELETE_COL_NAME = "�f�[�^�폜"

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : �S���R�[�h�̃f�[�^�ۑ�
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub SaveAllButton_Click()

    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)

    Dim updateColumns As Collection, conditions As Collection

    Set updateColumns = New Collection
    Set conditions = New Collection
            
    updateColumns.Add ("SAMPLE_TEXT")
    updateColumns.Add ("SAMPLE_COMBOBOX_TEXT")
    updateColumns.Add ("SAMPLE_VALUE")
    updateColumns.Add ("SAMPLE_DATE")
    updateColumns.Add ("SAMPLE_FLAG")
    updateColumns.Add ("SAMPLE_CODE")
    
    conditions.Add ("SAMPLE_ID")
    
    Call SaveRecords(table, "sample_table", updateColumns, conditions)
    
    MsgBox "�S���R�[�h�̃f�[�^��ۑ����܂����B"

End Sub

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' Brief : 1���R�[�h�̃f�[�^�ۑ��������͍폜
' Note  :
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    ' Debug.Print table.ListColumns(SAVE_COL_NAME).DataBodyRange.Address
    
    ' �s�C���f�b�N�X���擾
    Dim indexRow As Long
    indexRow = ActiveCell.Row - table.DataBodyRange(1).Row + 1
    
    ' ���o���s�A�W�v�s�̑I���`�F�b�N
    If indexRow = 0 Or table.ListRows.Count < indexRow Then Exit Sub
    
    
    Dim updateColumns As Collection, conditions As Collection
    
    ' ---- 1���R�[�h�̃f�[�^��ۑ� ----
    Dim saveRange As Range
    Set saveRange = table.ListColumns(SAVE_COL_NAME).DataBodyRange

    If Not Intersect(Target, saveRange) Is Nothing Then '�Ώ۔͈͂Əd�Ȃ�Ƃ��낪�u�Ȃ��v�ł͂Ȃ��ꍇ
        
        Cancel = True ' �A�N�e�B�u������
                
        Set updateColumns = New Collection
        Set conditions = New Collection
                
        updateColumns.Add ("SAMPLE_TEXT")
        updateColumns.Add ("SAMPLE_COMBOBOX_TEXT")
        updateColumns.Add ("SAMPLE_VALUE")
        updateColumns.Add ("SAMPLE_DATE")
        updateColumns.Add ("SAMPLE_FLAG")
        updateColumns.Add ("SAMPLE_CODE")
        
        conditions.Add ("SAMPLE_ID")
        
        Call SaveRecord(table, indexRow, "sample_table", updateColumns, conditions)
        
        MsgBox "�N���b�N�������R�[�h�̃f�[�^��ۑ����܂����B"
        Exit Sub
        
    End If
    
    ' ---- 1���R�[�h�̃f�[�^���폜 ----
    Dim deleteRange As Range
    Set deleteRange = table.ListColumns(DELETE_COL_NAME).DataBodyRange
    
    If Not Intersect(Target, deleteRange) Is Nothing Then '�Ώ۔͈͂Əd�Ȃ�Ƃ��낪�u�Ȃ��v�ł͂Ȃ��ꍇ
        
        Cancel = True ' �A�N�e�B�u������
                
        Set conditions = New Collection
                        
        conditions.Add ("SAMPLE_ID")
        
        Call DeleteRecord(table, indexRow, "sample_table", conditions)
        Call ExecuteSelectSql("�T���v��2-1")
        
        MsgBox "�N���b�N�������R�[�h�̃f�[�^���폜���܂����B"
        Exit Sub
        
    End If
End Sub