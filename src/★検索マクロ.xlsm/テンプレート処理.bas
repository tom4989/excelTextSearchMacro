Attribute VB_Name = "�e���v���[�g����"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' �萔
' ---------------------------------------------------------------------------------------------------------------------

Public Const TEMPLATE_SHEET_NAME = "���`"

' �Ώۂ̊g���q
Private Const FILE_EXTENSION = "xls,xlsx,xlsm"

Private Const KEY_�p�X = "�p�X"
Private Const KEY_�Ώۃu�b�N�V�[�g = "�Ώۃu�b�N�V�[�g��"
Private Const KEY_�ΏۊO�u�b�N�V�[�g = "�ΏۊO�u�b�N�V�[�g��"

' ---------------------------------------------------------------------------------------------------------------------
' �ϐ�
' ---------------------------------------------------------------------------------------------------------------------

' �ݒ�l���X�g
Public obj�ݒ�l�V�[�g As cls�ݒ�l�V�[�g

' ���`�ŏI��
Public lng���`�ŏI�� As Long

' ���`�J�n�s
Public lng���`�J�n�s As Long


' *********************************************************************************************************************
' * �@�\�@�F�}�N���Ăяo�����i�V�[�g����̎w��p�j
' *********************************************************************************************************************

Sub �}�N���J�n()

    Call init�J�n����

    log ("----------------------------------------------------------------------------------------------------")
    log ("�}�N���J�n")
    log ("----------------------------------------------------------------------------------------------------")
    
    Set obj�ݒ�l�V�[�g = New cls�ݒ�l�V�[�g
    obj�ݒ�l�V�[�g.���[�h (ActiveSheet.Name)
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' ����������
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' �ŗL�����i�}�N���Ăяo�����j���̃V�[�g�����擾����B
    Dim ws�}�N���Ăяo�����V�[�g As Worksheet
    Set ws�}�N���Ăяo�����V�[�g = ActiveSheet

    ' ��ConcreateProcess���̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
    Call �S�̑O����(ws�}�N���Ăяo�����V�[�g)

    lng���`�J�n�s = �ŏI�s�擾(ThisWorkbook.Sheets(TEMPLATE_SHEET_NAME), False) + 1
    lng���`�ŏI�� = �ŏI��擾(ThisWorkbook.Sheets(TEMPLATE_SHEET_NAME))

    ' -----------------------------------------------------------------------------------------------------------------
    ' �p�X�̑��݃`�F�b�N
    ' -----------------------------------------------------------------------------------------------------------------

    Dim txt�p�X As Variant

    With ws�}�N���Ăяo�����V�[�g

        ' �Ώۃf�B���N�g�������[�v
        For Each txt�p�X In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�p�X)
            
            ' �f�B���N�g���܂��́A�t�@�C���̑��݃`�F�b�N
            If isDirectoryExist(CStr(txt�p�X)) < 0 Then
                
                MsgBox "�ȉ��̃p�X�͑��݂��܂���B" + Chr(10) + "�u" + txt�p�X + "�v"
                End
            End If
        Next
    End With

    ' -----------------------------------------------------------------------------------------------------------------
    ' �t�@�C�����̎��W
    ' -----------------------------------------------------------------------------------------------------------------

    Call set�X�e�[�^�X�o�[("�Ώۃt�@�C���W�v��...")
    
    ' �����Ώۂ̊g���q��ݒ肷��B
    Dim var�t�@�C���g���q As Variant
    var�t�@�C���g���q = Split(FILE_EXTENSION, ",")
    
    ' �����Ώۂ̃t�@�C�����ꗗ�i�t���p�X���t�@�C�����j
    Dim txt�p�X�ꗗ() As String
    
    '�Ώۃf�B���N�g�������[�v
    For Each txt�p�X In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�p�X)
            
        '�w��̒l���t�@�C���̏ꍇ�A���̒l�����X�g�ɒǉ����A
        ' �f�B���N�g���̏ꍇ�́A�t�@�C�����̈ꗗ�𓮓I�Ɏ擾���Ēǉ�����B
        If isDirectoryExist(CStr(txt�p�X)) = 2 Then
            
            ' �w��̒l���t�@�C���������ꍇ�A���̒l�����X�g�ɒǉ�
            ' �t���p�X���t�@�C������ǉ��i�[�B
            Call �ꎟ�z��ɒl��ǉ�(txt�p�X�ꗗ, CStr(txt�p�X))
        Else
                    
            ' ���݂̃f�B���N�g���z���̃t�@�C�������擾
            Call doRepeat(txt�p�X, var�t�@�C���g���q, txt�p�X�ꗗ, True)
                
        End If
    Next

    ' -----------------------------------------------------------------------------------------------------------------
    ' �t�@�C���������\�b�h�̌Ăяo��
    ' -----------------------------------------------------------------------------------------------------------------
    
    Call �t�@�C������(txt�p�X�ꗗ)
    
    ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
    Call �S�̌㏈��(ws�}�N���Ăяo�����V�[�g)
    
    MsgBox "�������I�����܂����B�i�������ԁF" & get��������() & ")"

    log ("----------------------------------------------------------------------------------------------------")
    log ("�}�N���I��")
    log ("----------------------------------------------------------------------------------------------------")

End Sub

' *********************************************************************************************************************
' * �@�\�@�F�Ώۃt�@�C���̏������s���B
' * �����@�FvarArray �z��
' * �߂�l�F���茋�ʁi1:�z��/0:��̔z��/-1:�z��ł͂Ȃ��j
' *********************************************************************************************************************
'
Function �t�@�C������(txt�p�X�ꗗ() As String)

    ' �t�@�C�����̈ꗗ���󂾂����ꍇ�A��Function�𒆒f����B
    If IsArrayEx(txt�p�X�ꗗ) < 1 Then
        MsgBox "�����Ώۃt�@�C�������݂��܂���B"
        Exit Function
    End If
    
    Dim defaultSaveFormat As Long
    defaultSaveFormat = Application.defaultSaveFormat
    
    Application.DisplayAlerts = False ' �t�@�C�����J���ۂ̌x���𖳌�
    Application.ScreenUpdating = False ' ��ʕ\���X�V�𖳌�
    
    ' �������ʕێ��p
    Dim results() As Variant
    
    Dim index As Long, total As Long
        
    index = 1
    total = UBound(txt�p�X�ꗗ) + 1
    
    Dim txt�p�X As Variant
    
    For Each txt�p�X In txt�p�X�ꗗ
    
        ' -------------------------------------------------------------------------------------------------------------
        ' �Ώۃu�b�N���J���āA�S�V�[�g���̏������s���B
        ' -------------------------------------------------------------------------------------------------------------

        Call set�X�e�[�^�X�o�[("(" & index & "/" & total & ")" & �t�@�C�����擾(CStr(txt�p�X)))
        index = index + 1
        
        Dim wb�Ώۃu�b�N As Workbook
        Set wb�Ώۃu�b�N = Workbooks.Open(txt�p�X, UpdateLinks:=0, IgnoreReadOnlyRecommended:=False)
        
        ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
        If �u�b�NOPEN�㏈��(txt�p�X, wb�Ώۃu�b�N, results) Then
        
            ' �u�b�NOPEN�㏈���̕Ԃ�l��True�̏ꍇ�A�V�[�g���̏����𑱍s����
        
            Dim var�Ώۃu�b�N�V�[�g As Variant
            var�Ώۃu�b�N�V�[�g = obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�Ώۃu�b�N�V�[�g)
            
            Dim var�ΏۊO�u�b�N�V�[�g As Variant
            var�ΏۊO�u�b�N�V�[�g = obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�ΏۊO�u�b�N�V�[�g)
        
            Dim i As Long
            For i = 1 To wb�Ώۃu�b�N.Worksheets.Count
            
                Dim ws�ΏۃV�[�g As Worksheet
                Set ws�ΏۃV�[�g = wb�Ώۃu�b�N.Worksheets(i)
                
                Dim txt�u�b�N�V�[�g�� As String
                txt�u�b�N�V�[�g�� = "[" & wb�Ώۃu�b�N.Name & "]" & ws�ΏۃV�[�g.Name
                
                ' �Ώۂ̃u�b�N�^�V�[�g���`�F�b�N
                If Not f_�z��܂܂�Ă��邩�`�F�b�N(var�Ώۃu�b�N�V�[�g, txt�u�b�N�V�[�g��, False, True) Then
                
                    log ("�Ώۃu�b�N�V�[�g�ɕs��v�F" & txt�u�b�N�V�[�g��)
                
                ' �ΏۊO�̃u�b�N�^�V�[�g�łȂ����`�F�b�N
                ElseIf f_�z��܂܂�Ă��邩�`�F�b�N(var�ΏۊO�u�b�N�V�[�g, txt�u�b�N�V�[�g��, False, True) Then
                
                    log ("�ΏۊO�u�b�N�V�[�g�Ɉ�v�F" & txt�u�b�N�V�[�g��)
                Else
                
                    ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
                    Call �V�[�g������(txt�p�X, ws�ΏۃV�[�g, results)
                End If
                
            Next i
            
        End If
        
        ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
        Dim �t�@�C��CLOSE���@�敪�l As Long
        �t�@�C��CLOSE���@�敪�l = �u�b�NCLOSE�O����(txt�p�X, wb�Ώۃu�b�N, results)
        
        If �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����Ȃ��ŕ��� Then
            wb�Ώۃu�b�N.Close
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����ĕ��� Then
            wb�Ώۃu�b�N.Save
            wb�Ώۃu�b�N.Close
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����Ȃ��ŕ��Ȃ� Then
            
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�ۑ����ĕ��Ȃ� Then
            wb�Ώۃu�b�N.Save
        ElseIf �t�@�C��CLOSE���@�敪�l = �t�@�C��CLOSE���@�敪.�������f Then
            End
        End If
    Next
        
    ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
    ' ���s���ʂ̕ҏW�i���ʂ̃}�[�W�A���ёւ��A�t�B���^�����O���j
    Call ���s���ʓ��e�ҏW����(results)
    
    Dim wb���ʃu�b�N As Workbook
    
    If Not Not results Then
    
        If UBound(results, 2) <> 0 Then
        
            ' �t�@�C���̕ۑ��`����excel2007�`���i.xlsx)�ɕύX
            Application.defaultSaveFormat = xlOpenXMLWorkbook
            
            Set wb���ʃu�b�N = Workbooks.Add
            
            ' ���u�b�N�ɃV�[�g�u���`�v���p�ӂ���Ă���ꍇ�A�w��u�b�N�̐擪�ɃR�s�[������A
            ' �V�[�g�����u�������ʁv�ɕύX����B�i�Ȃ��ꍇ�͐V�K�쐬�u�b�N��sheet1�𗘗p�j
            Call ���`�V�[�g�R�s�[(wb���ʃu�b�N)
            
            ' ���ʓ\��t���s�̎擾�B
            ' A��ɒl���ݒ肳��Ă���s���A�\�藓�Ƃ��Ă��̍s�����擾����
            Dim lng�ő�s As Long
            With wb���ʃu�b�N.ActiveSheet.UsedRange
                lng�ő�s = .Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
            End With
            
            ' ���ʓ\��t���s�̐ݒ�B
            lng�ő�s = lng�ő�s + 1
            
            ' ���ʓ\��t��
            wb���ʃu�b�N.ActiveSheet.Range( _
                Cells(lng�ő�s, 1), _
                Cells(UBound(results, 2) + lng���`�J�n�s, UBound(results) + 1)) = �񎟌��z��s��t�](results)
            
            Dim lng�ő�� As Long
            ' �����R�s�[
            With wb���ʃu�b�N.ActiveSheet
                lng�ő�s = �ŏI�s�擾(wb���ʃu�b�N.ActiveSheet, False)
                lng�ő�� = �ŏI��擾(wb���ʃu�b�N.ActiveSheet, False)
                
                .Range(.Cells(lng���`�J�n�s, 1), .Cells(lng���`�J�n�s, lng�ő��)).Copy
                .Range(.Cells(lng���`�J�n�s + 1, 1), .Cells(lng�ő�s, lng�ő��)).PasteSpecial (xlPasteFormats)
                
                For i = lng���`�J�n�s To lng�ő�s
                
                    ' �n�C�p�[�����N�ݒ�
                    Dim strHyperLink As String
                    strHyperLink = "=HYPERLINK(""[""&A" & i & "&""\""&B" & i & "&""]""&" & _
                        "C" & i & "&""!" & .Cells(i, 4) & """,""" & .Cells(i, 4) & """)"
            
                    .Range(.Cells(i, 4), .Cells(i, 4)).Value = strHyperLink
                Next
            End With
            
            ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
            Call ���s���ʏ����ҏW����(wb���ʃu�b�N.ActiveSheet)
            
            ' "A1"��I����Ԃɂ���
            wb���ʃu�b�N.ActiveSheet.Cells(1, 1).Select
            
            ' �V�[�g���u�������ʁv�ȊO�̃V�[�g���폜����
            Call �s�v�V�[�g�폜(wb���ʃu�b�N, RESULT_SHEET_NAME)
            
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
            wb���ʃu�b�N.Activate
        End If
    End If

End Function


' *********************************************************************************************************************
' * �@�\�@�F���u�b�N�̃V�[�g�u���`�v���w��u�b�N�̐擪�ɃR�s�[������A
' * �@�@�@�@�V�[�g�����u�������ʁv�ɕύX����
' *********************************************************************************************************************
'
Sub ���`�V�[�g�R�s�[(wb�R�s�[��u�b�N As Workbook)

    ' �}�N�����s���̃u�b�N���A�N�e�B�u�ɂ���
    ThisWorkbook.Activate
    
    ' �V�[�g�u���`�v���������ꍇ�A�w��u�b�N�ɃR�s�[�i��ԑO�ɑ}���j����
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count ' �V�[�g�̐������[�v����
    
        Dim targetSheet As Worksheet
        Set targetSheet = ThisWorkbook.Worksheets(i)
        
        If TEMPLATE_SHEET_NAME = ThisWorkbook.Worksheets(i).Name Then
        
            ThisWorkbook.Sheets(TEMPLATE_SHEET_NAME).Copy Before:=wb�R�s�[��u�b�N.Sheets(1)
        End If
        
    Next i
    
    ' �}�N�������s���̃u�b�N���A�N�e�B�u�ɂ���
    Workbooks(wb�R�s�[��u�b�N.Name).Sheets(TEMPLATE_SHEET_NAME).Activate
    
    ' �V�[�g�����u�������ʁv�ɕύX����
    Workbooks(wb�R�s�[��u�b�N.Name).Sheets(TEMPLATE_SHEET_NAME).Name = RESULT_SHEET_NAME
    
End Sub
