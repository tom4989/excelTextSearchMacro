Attribute VB_Name = "�e���v���[�g����"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' �萔
' ---------------------------------------------------------------------------------------------------------------------

Private Const KEY_�p�X = "�p�X"

' ---------------------------------------------------------------------------------------------------------------------
' �ϐ�
' ---------------------------------------------------------------------------------------------------------------------

' �ݒ�l���X�g
Public obj�ݒ�l�V�[�g As cls�ݒ�l�V�[�g

' *********************************************************************************************************************
' * �@�\�@�F�}�N���Ăяo�����i�V�[�g����̎w��p�j
' *********************************************************************************************************************

Sub �}�N���J�n()

    Call init�J�n����
    
    Dim wsMainSheet As Worksheet
    Dim fileCheck As Long
    
    Set obj�ݒ�l�V�[�g = New cls�ݒ�l�V�[�g
    obj�ݒ�l�V�[�g.���[�h (ActiveSheet.Name)
    
    ' �^�C�g�����ɑ΂��郊�X�g�̏��iRange���j
    ' Dim currentDirPathRangeList As Range, currentDirPathRange As Range
    ' Dim subDirCheckBoxRangeList As Range, subDirCheckBoxRange As Range
    
    ' �����Ώۂ̃t�@�C�����ꗗ�i�t���p�X���t�@�C�����j
    Dim fileNames() As String
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' ����������
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' �����Ώۂ̊g���q��ݒ肷��B
    Dim fileExtention As Variant
    fileExtention = Split(FILE_EXTENSION, ",")
    
    ' �ŗL�����i�}�N���Ăяo�����j���̃V�[�g�����擾����B
    ' Set wsMainSheet = MainSheet
    Set wsMainSheet = ActiveSheet
    
    ' �ŗL�����i�}�N���Ăяo�����j���̃p�X�����擾����B
    ' Set currentDirPathRangeList = �^�C�g�����w��Ń��X�g�l��Range�����擾(TITLE_NAME_BY_TARGET_DIR, wsMainSheet)
    ' Set subDirCheckBoxRangeList = �^�C�g�����w��Ń��X�g�l��Range�����擾(TITLE_NAME_BY_DO_SUB_DIR, wsMainSheet)
    
    ' ��ConcreateProcess���̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
    Call �S�̑O����(wsMainSheet)


    ' -----------------------------------------------------------------------------------------------------------------
    ' �p�X�̑��݃`�F�b�N
    ' -----------------------------------------------------------------------------------------------------------------

    Dim txt�p�X As Variant

    With wsMainSheet

        Dim i As Long
        i = 0
        ' �Ώۃf�B���N�g�������[�v
        ' If Not (obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�p�X) Is Nothing) Then
            For Each txt�p�X In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�p�X)
            
                ' �f�B���N�g���܂��́A�t�@�C���̑��݃`�F�b�N
                fileCheck = isDirectoryExist(CStr(txt�p�X))
                
                If 0 > fileCheck Then
                    MsgBox "�ȉ��̃p�X�͑��݂��܂���B" + Chr(10) + "�u" + txt�p�X + "�v"
                    End
                End If
                i = i + 1
            Next
        ' End If
    End With

    ' -----------------------------------------------------------------------------------------------------------------
    ' �t�@�C�����̎��W
    ' -----------------------------------------------------------------------------------------------------------------

    Call set�X�e�[�^�X�o�[("�Ώۃt�@�C���W�v��...")
    
    With ActiveSheet
    
        i = 1
        '�Ώۃf�B���N�g�������[�v
        ' If Not (obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�p�X) Is Nothing) Then
            For Each txt�p�X In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�p�X)
            
                '�w��̒l���t�@�C���̏ꍇ�A���̒l�����X�g�ɒǉ����A�f�B���N�g���̏ꍇ�́A�t�@�C�����̈ꗗ�𓮓I�Ɏ擾���Ēǉ�����B
                fileCheck = isDirectoryExist(CStr(txt�p�X))
                If 2 = fileCheck Then
                    ' �w��̒l���t�@�C���������ꍇ�A���̒l�����X�g�ɒǉ�
                    ' �t���p�X���t�@�C������ǉ��i�[�B
                    Call �ꎟ�z��ɒl��ǉ�(fileNames, CStr(txt�p�X))
                Else
                    
                    ' ���I�[�g�V�F�C�v���̎擾��
                    Dim shapesCount As Long
                    Dim checkBoxChecked As Variant
                    Dim topLeftCellRow As Variant, topLeftCellColumn As Variant
            
                    ' �I�[�g�V�F�C�v�i�`�F�b�N�{�b�N�X�j�����擾�B
                    Dim ShapesInfoList As Variant
                    ShapesInfoList = getShapesProperty(wsMainSheet, msoFormControl, xlCheckBox)
                    
                    ' �ΏۃZ���s�̃`�F�b�N�{�b�N�X�̃`�F�b�N��Ԃ��擾�iboolean�`���Łj
                    checkBoxChecked = True
                    
                    ' ���݂̃f�B���N�g���z���̃t�@�C�������擾
                    Call doRepeat(txt�p�X, fileExtention, fileNames, checkBoxChecked)
                
                End If
                
                i = i + 1
            Next
            
        ' End If
        
    End With
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' �t�@�C���������\�b�h�̌Ăяo��
    ' -----------------------------------------------------------------------------------------------------------------
    
    Call �t�@�C������(fileNames)
    
    ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
    Call �S�̌㏈��(wsMainSheet)
    
    MsgBox "�������I�����܂����B�i�������ԁF" & get��������() & ")"

End Sub

' *********************************************************************************************************************
' * �@�\�@�F�Ώۃt�@�C���̏������s���B
' * �����@�FvarArray �z��
' * �߂�l�F���茋�ʁi1:�z��/0:��̔z��/-1:�z��ł͂Ȃ��j
' *********************************************************************************************************************
'
Function �t�@�C������(fileNames() As String)

    ' �t�@�C�����̈ꗗ���󂾂����ꍇ�A��Function�𒆒f����B
    If 1 > IsArrayEx(fileNames) Then
        MsgBox "�����Ώۃt�@�C�������݂��܂���B"
        Exit Function
    End If
    
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileName As Variant
    Dim targetWB As Workbook
    Dim targetSheet As Worksheet
    
    Dim index As Long, total As Long
    
    Dim defaultSaveFormat As Long
    defaultSaveFormat = Application.defaultSaveFormat
    
    ' �V�[�g���̏����Ăяo���s�v�t���O
    Dim unDealTargetSheetFlag As Boolean
    
    ' �������ʕێ��p
    Dim results() As Variant
    
    index = 1
    total = UBound(fileNames) + 1
    
    Application.DisplayAlerts = False ' �t�@�C�����J���ۂ̌x���𖳌�
    Application.ScreenUpdating = False ' ��ʕ\���X�V�𖳌�
    
    For Each fileName In fileNames
    
        ' -------------------------------------------------------------------------------------------------------------
        ' �Ώۃu�b�N���J���āA�S�V�[�g���̏������s���B
        ' -------------------------------------------------------------------------------------------------------------

        Call set�X�e�[�^�X�o�[("(" & index & "/" & total & ")" & FSO.GetFileName(fileName))
        index = index + 1
        
        Set targetWB = Workbooks.Open(fileName, UpdateLinks:=0, IgnoreReadOnlyRecommended:=False)
        
        ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
        unDealTargetSheetFlag = �u�b�NOPEN�㏈��(fileName, targetWB, results)
        
        If False = unDealTargetSheetFlag Then
            Dim i As Integer
            For i = 1 To targetWB.Worksheets.Count ' �V�[�g�̐������[�v����
            
                Set targetSheet = targetWB.Worksheets(i)
                
                ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
                Call �V�[�g������(fileName, targetSheet, results)
                
            Next i
            
        End If
        
        Dim �t�@�C��CLOSE���@�敪�l As Long
        
        ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
        �t�@�C��CLOSE���@�敪�l = �u�b�NCLOSE�O����(fileName, targetWB, results)
        
        If �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����Ȃ��ŕ��� Then
            targetWB.Close
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����ĕ��� Then
            targetWB.Save
            targetWB.Close
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����Ȃ��ŕ��Ȃ� Then
            
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����ĕ��Ȃ� Then
            targetWB.Save
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�������f Then
            End
        End If
    Next
        
    ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
    ' ���s���ʂ̕ҏW�i���ʂ̃}�[�W�A���ёւ��A�t�B���^�����O���j
    Call ���s���ʓ��e�ҏW����(results)
    
    If Not Not results Then
    
        If UBound(results, 2) <> 0 Then
        
            ' �t�@�C���̕ۑ��`����excel2007�`���i.xlsx)�ɕύX
            Application.defaultSaveFormat = xlOpenXMLWorkbook
            
            Set targetWB = Workbooks.Add
            
            ' ���u�b�N�ɃV�[�g�u���`�v���p�ӂ���Ă���ꍇ�A�w��u�b�N�̐擪�ɃR�s�[������A
            ' �V�[�g�����u�������ʁv�ɕύX����B�i�Ȃ��ꍇ�͐V�K�쐬�u�b�N��sheet1�𗘗p�j
            Call ���`�V�[�g�R�s�[(targetWB)
            
            ' ���ʓ\��t���s�̎擾�B
            ' A��ɒl���ݒ肳��Ă���s���A�\�藓�Ƃ��Ă��̍s�����擾����
            Dim MaxRow As Integer
            With targetWB.ActiveSheet.UsedRange
                MaxRow = .Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
            End With
            ' ���ʓ\��t���s�̐ݒ�B
            MaxRow = MaxRow + 1
            
            ' ���ʓ\��t��
            targetWB.ActiveSheet.Range(Cells(MaxRow, 1), Cells(UBound(results, 2) + 2, UBound(results) + 1)) = �񎟌��z��s��t�](results)
            
            Dim MaxCol As Integer
            ' �����R�s�[
            With targetWB.ActiveSheet
                MaxRow = .UsedRange.Find("*", , xlFormulas, xlByRows, xlPrevious).Row
                MaxCol = .UsedRange.Find("*", , xlFormulas, xlByColumns, xlPrevious).Column
                
                .Range(.Cells(2, 1), .Cells(2, MaxCol)).Copy
                .Range(.Cells(2 + 1, 1), .Cells(MaxRow, MaxCol)).PasteSpecial (xlPasteFormats)
            End With
            
            ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
            Call ���s���ʏ����ҏW����(targetWB.ActiveSheet)
            
            ' "A1"��I����Ԃɂ���
            targetWB.ActiveSheet.Cells(1, 1).Select
            
            ' �V�[�g���u�������ʁv�ȊO�̃V�[�g���폜����
            Call �s�v�V�[�g�폜(targetWB, RESULT_SHEET_NAME)
            
        Else
            
            MsgBox "�������ʂ�0���ł��B"
        End If
        
    Else
    
        MsgBox "�������ʂ�0���ł��B"
        
    End If
            
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Application.StatusBar = False
    
    ' �t�@�C���̕ۑ��`�������̏�Ԃɖ߂�
    Application.defaultSaveFormat = defaultSaveFormat
    
    If Not Not results Then
        If UBound(results, 2) <> 0 Then
            targetWB.Activate
        End If
    End If

End Function


' *********************************************************************************************************************
' * �@�\�@�F���u�b�N�̃V�[�g�u���`�v���w��u�b�N�̐擪�ɃR�s�[������A
' * �@�@�@�@�V�[�g�����u�������ʁv�ɕύX����
' *********************************************************************************************************************
'
Sub ���`�V�[�g�R�s�[(targetWB As Workbook)

    Dim myWorkBook  As String
    Dim newWorkBook As String
    Dim targetSheet As Worksheet
    Dim sheetName   As String
    
    ' �}�N�������s���̃u�b�N�����擾
    myWorkBook = ThisWorkbook.Name
    
    ' �V�K�u�b�N�����擾
    newWorkBook = targetWB.Name
    
    ' �}�N�����s���̃u�b�N���A�N�e�B�u�ɂ���
    Workbooks(myWorkBook).Activate
    
    ' �V�[�g�u���`�v���������ꍇ�A�w��u�b�N�ɃR�s�[�i��ԑO�ɑ}���j����
    Dim i As Integer
    For i = 1 To Workbooks(myWorkBook).Worksheets.Count ' �V�[�g�̐������[�v����
    
        Set targetSheet = Workbooks(myWorkBook).Worksheets(i)
        
        If TEMPLATE_SHEET_NAME = targetSheet.Name Then
            Workbooks(myWorkBook).Sheets(TEMPLATE_SHEET_NAME).Copy _
            Before:=Workbooks(newWorkBook).Sheets(1)
        End If
        
    Next i
    
    ' �}�N�������s���̃u�b�N���A�N�e�B�u�ɂ���
    Workbooks(targetWB.Name).Sheets(TEMPLATE_SHEET_NAME).Activate
    ' �V�[�g�����u�������ʁv�ɕύX����
    Workbooks(targetWB.Name).Sheets(TEMPLATE_SHEET_NAME).Name = RESULT_SHEET_NAME
    
End Sub
