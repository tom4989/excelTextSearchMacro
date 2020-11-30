Attribute VB_Name = "��������"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' �萔�i���ʁj
' ---------------------------------------------------------------------------------------------------------------------

' ���`�V�[�g�R�s�[�p�i���ʁj
Public Const TEMPLATE_SHEET_NAME = "���`"
Public Const RESULT_SHEET_NAME = "��������"

' �Ώۂ̊g���q
Public Const FILE_EXTENSION = "xls,xlsx,xlsm"

' �������ʃV�[�g�f�[�^�\�t�����̗�
Private Const RESULT_COL_LENGTH = 6

' ---------------------------------------------------------------------------------------------------------------------
' �萔�i�ʁj
' ---------------------------------------------------------------------------------------------------------------------

Private Const KEY_�p�X = "�p�X"
Private Const KEY_�������[�h = "�������[�h"
Private Const KEY_�Ώۃu�b�N�V�[�g = "�Ώۃu�b�N�V�[�g"
Private Const KEY_�ΏۊO�u�b�N�V�[�g = "�ΏۊO�u�b�N�V�[�g"

' ---------------------------------------------------------------------------------------------------------------------
' �ϐ�
' ---------------------------------------------------------------------------------------------------------------------

' ������������
Dim lngResultCount As Long

' #####################################################################################################################
' #
' # �e���v���[�g���\�b�h(�e���v���[�g��������Ăяo����郁�\�b�h�j
' #
' # 1. �S�̑O����()            �������s�O��1�x�������s��������������������
' # 2. �u�b�NOPEN�㏈��()      ���o���ꂽ�t�@�C���̃u�b�N���Ƃɍs��������������������
' #                            �i�V�[�g���̏����Ăяo�����s�v���̔���l(boolean)��ԋp����j
' # 3. �V�[�g������()          ���o���ꂽ�t�@�C����1�V�[�g���Ƃɍs��������������������
' # 4. �u�b�NCLOSE�O����()     ���o���ꂽ�t�@�C���̃u�b�N���Ƃɍs�������㏈������������
' # 5. ���s���ʓ��e�ҏW����()  ���s���ʂɂ��āA�t�@�C���ɏo�͂���O�ɕҏW�������ꍇ�Ɏ�������i�d���̍폜�A�\�[�g���j
' # 6. ���s���ʏ����ҏW����()  �t�@�C���ɏo�͂�����̎��s���ʂ�ҏW�������ꍇ�Ɏ�������i�n�C�p�[�����N�̐ݒ蓙�j
' # 7. �S�̌㏈��()            �������s���1�x�������s��������������������
' #
' #####################################################################################################################
'

' *********************************************************************************************************************
' �@�\�@�F�ŗL�������̑O����
' *********************************************************************************************************************
'
Function �S�̑O����(targetSheet As Worksheet)

    ' -----------------------------------------------------------------------------------------------------------------
    ' ����������
    ' -----------------------------------------------------------------------------------------------------------------

    ' �������������̏�����
    ' resultCount = 0

    ' -----------------------------------------------------------------------------------------------------------------
    ' �O����
    ' -----------------------------------------------------------------------------------------------------------------

End Function

' *********************************************************************************************************************
' �@�\�@�F���o���ꂽ�t�@�C���̃u�b�N���Ƃɍs��������������������i�V�[�g���̏����Ăяo�����s�v���̔���l(boolean)��ԋp����j
' *********************************************************************************************************************
'
Function �u�b�NOPEN�㏈��(fileName As Variant, targetWB As Workbook, ByRef results() As Variant) As Boolean


End Function

' *********************************************************************************************************************
' �@�\�@�F���o���ꂽ�t�@�C����1�V�[�g���Ƃɍs��������������������
' *********************************************************************************************************************
'
Function �V�[�g������(fileName As Variant, targetSheet As Worksheet, ByRef results() As Variant)

    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim targetWB As Workbook
    
    Dim ShapesInfoList As Variant
    Dim ShapesInf As Variant
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' ����
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' �w�肳�ꂽ�����������X�g�̕�����̌������ʂ����W����B
    
    ' �ΏۃV�[�g�̌������ʂ��uFoundAddr�v�Ɋi�[����B
    Dim firstAddress As String
    Dim FoundCell As Range
    
    Dim lngResultCount As Long ' ���ʌ���
    
    Dim txt�������[�h As Variant
    
    ' �����������X�g�����[�v
    For Each txt�������[�h In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�������[�h)
    
        ' �����������Ȃ��ꍇ�A���̌�����������������
        If "" = txt�������[�h Then
        
            GoTo ContinueBySearchArg
        End If
        
        ' ���Z���̌�����
        Set FoundCell = targetSheet.UsedRange.Find(what:=txt�������[�h, LookIn:=xlValues, _
            LookAt:=xlPart, MatchCase:=False, MatchByte:=False)
            
        ' �Z���ւ̌������ʂ��Ȃ��ꍇ
        If FoundCell Is Nothing Then
            ' �������ʂ��Ȃ������ꍇ���̌�����������������
            GoTo GotoCellSearchEnd
        End If
        
        firstAddress = FoundCell.Address ' �������ʂ̃A�h���X��z��Ɋi�[
        
        Do
            ' ���ʂ��i�[����
            Call reDimResult(RESULT_COL_LENGTH, results)                   ' ���ʕێ��̔z��쐬
            lngResultCount = UBound(results, 2)
            
            results(0, lngResultCount) = txt�������[�h                     ' ��������
            results(1, lngResultCount) = FSO.GetParentFolderName(fileName) ' �t�H���_��
            results(2, lngResultCount) = FSO.GetFileName(fileName)         ' �t�@�C����
            results(3, lngResultCount) = targetSheet.Name                  ' �V�[�g��
            results(4, lngResultCount) = FoundCell.Address(False, False)   ' ���W
            results(5, lngResultCount) = "�Z��"                            ' �Z���^�I�[�g�V�F�C�v
            results(6, lngResultCount) = FoundCell.Value                   ' ������
            
            Set FoundCell = targetSheet.UsedRange.FindNext(After:=FoundCell)
            
        Loop Until FoundCell.Address = firstAddress
        
GotoCellSearchEnd:

        ' ���I�[�g�V�F�C�v�̌�����
        ShapesInfoList = getShapesProperty(targetSheet)
        Dim i As Integer
        Dim textValue As Variant
        i = 0
        
        ' �����������X�g�����[�v
        If Not IsEmpty(ShapesInfoList) Then
            For i = LBound(ShapesInfoList) To UBound(ShapesInfoList)
                textValue = ShapesInfoList(i, 2)
                If Not IsEmpty(textValue) And InStr(textValue, txt�������[�h) Then
                
                    ' ���ʂ��i�[����
                    Call reDimResult(RESULT_COL_LENGTH, results)                   ' ���ʕێ��̔z��쐬
                    lngResultCount = UBound(results, 2)
                    
                    results(0, lngResultCount) = txt�������[�h                     ' ��������
                    results(1, lngResultCount) = FSO.GetParentFolderName(fileName) ' �t�H���_��
                    results(2, lngResultCount) = FSO.GetFileName(fileName)         ' �t�@�C����
                    results(3, lngResultCount) = targetSheet.Name                  ' �V�[�g��
                    results(4, lngResultCount) = ShapesInfoList(i, 7)              ' ���W
                    results(5, lngResultCount) = "�I�[�g�V�F�C�v"                  ' �Z���^�I�[�g�V�F�C�v
                    results(6, lngResultCount) = textValue                         ' ������
                End If
            Next i
        End If
        
ContinueBySearchArg:

    Next
    
End Function

' *********************************************************************************************************************
' �@�\�@�F���o���ꂽ�t�@�C���̃u�b�N���Ƃɍs�������㏈������������
' *********************************************************************************************************************
'
Function �u�b�NCLOSE�O����(fileName As Variant, targetWB As Workbook, ByRef results() As Variant) As Long


End Function

' *********************************************************************************************************************
' �@�\�@�F���s���ʂɂ��āA�t�@�C���ɏo�͂���O�ɕҏW�������ꍇ�Ɏ�������i�d���̍폜�A�\�[�g���j
' *********************************************************************************************************************
'
Function ���s���ʓ��e�ҏW����(ByRef var�ϊ���() As Variant) As Variant

End Function

' *********************************************************************************************************************
' �@�\�@�F�t�@�C���ɏo�͂�����̎��s���ʂ�ҏW�������ꍇ�Ɏ�������i�n�C�p�[�����N�̐ݒ蓙�j
' *********************************************************************************************************************
'
Sub ���s���ʏ����ҏW����(ByRef targetSheet As Worksheet)

    Dim i, MaxRow, MaxCol As Long
    
    With targetSheet
        MaxRow = .UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        MaxCol = .UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
        
        ' �����R�s�[
        .Range(Cells(2, 1), Cells(2, MaxCol)).Copy
        .Range(Cells(2 + 1, 1), Cells(MaxRow, MaxCol)).PasteSpecial (xlPasteFormats)
        
        For i = 2 To MaxRow
            ' �n�C�p�[�����N�ݒ�
            Dim strHyperLink As String
            strHyperLink = editHYPERLINK����(.Cells(i, 2), .Cells(i, 3), .Cells(i, 4), .Cells(i, 5))
            
            .Range(.Cells(i, 5), .Cells(i, 5)).Value = strHyperLink
            
            ' �ԕ���
            Call �����Y�������̐ԑ�������(.Range(Cells(i, 7), Cells(i, 7)), Cells(i, 1))
            
        Next
    End With
          
End Sub

' *********************************************************************************************************************
' �@�\�@�F�������s���1�x�������s��������������������
' *********************************************************************************************************************
'
Function �S�̌㏈��(targetSheet As Worksheet)

End Function

' #####################################################################################################################
' #
' # �e���v���[�g���\�b�h�ȊO�̃��\�b�h
' #
' #####################################################################################################################
'

' �Ȃ�
