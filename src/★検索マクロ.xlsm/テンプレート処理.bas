Attribute VB_Name = "�e���v���[�g����"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' �萔
' ---------------------------------------------------------------------------------------------------------------------

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

' ���ʏo�̓V�[�g
Public obj���ʏo�̓V�[�g As cls���ʏo�̓V�[�g

' �p�X���ʕ�
Private txt�p�X���ʕ� As String

' �T�C�����g���[�h
Private flg�T�C�����g���[�h As Boolean

' *********************************************************************************************************************
' * �@�\�@�F�}�N���Ăяo�����i�V�[�g����̎w��p�j
' *********************************************************************************************************************

Sub �}�N���J�n(Optional arg�T�C�����g���[�h As Boolean = False)

    Call init�J�n����

    flg�T�C�����g���[�h = arg�T�C�����g���[�h

    log ("----------------------------------------------------------------------------------------------------")
    log ("�}�N���J�n")
    log ("----------------------------------------------------------------------------------------------------")
    
    Set obj�ݒ�l�V�[�g = New cls�ݒ�l�V�[�g
    obj�ݒ�l�V�[�g.���[�h (ActiveSheet.Name)
    
    Set obj���ʏo�̓V�[�g = New cls���ʏo�̓V�[�g
    Call obj���ʏo�̓V�[�g.������(SHEET_NAME_TEMPLATE, SHEET_NAME_RESULT, arg�T�C�����g���[�h)
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' ����������
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' �ŗL�����i�}�N���Ăяo�����j���̃V�[�g�����擾����B
    Dim ws�}�N���Ăяo�����V�[�g As Worksheet
    Set ws�}�N���Ăяo�����V�[�g = ActiveSheet

    ' ��ConcreateProcess���̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
    Call �S�̑O����(ws�}�N���Ăяo�����V�[�g)

    ' -----------------------------------------------------------------------------------------------------------------
    ' �p�X�̑��݃`�F�b�N
    ' -----------------------------------------------------------------------------------------------------------------

    Dim txt�p�X As Variant
    txt�p�X���ʕ� = ""

    With ws�}�N���Ăяo�����V�[�g

        ' �Ώۃf�B���N�g�������[�v
        For Each txt�p�X In obj�ݒ�l�V�[�g.�ݒ�l���X�g.Item(KEY_�p�X)
            
            ' �f�B���N�g���܂��́A�t�@�C���̑��݃`�F�b�N
            If isDirectoryExist(CStr(txt�p�X)) < 0 Then
                
                MsgBox "�ȉ��̃p�X�͑��݂��܂���B" + Chr(10) + "�u" + txt�p�X + "�v"
                End
            End If
            
            If txt�p�X���ʕ� = "" Then
            
                txt�p�X���ʕ� = txt�p�X
            
            Else
            
                txt�p�X���ʕ� = f_���ʕ��擾(CStr(txt�p�X), txt�p�X���ʕ�)
            
            End If
            
        Next
    End With
    
    txt�p�X���ʕ� = f_RTRIM(txt�p�X���ʕ�, "\")

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
            Call �ꎟ���z��ɒl��ǉ�(txt�p�X�ꗗ, CStr(txt�p�X))
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

    Call s_���b�Z�[�W�ʒm("�������I�����܂����B�i�������ԁF" & get��������() & ")", flg�T�C�����g���[�h)

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
        If �u�b�NOPEN�㏈��(txt�p�X, wb�Ώۃu�b�N) Then
        
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
                    Call �V�[�g������(txt�p�X, ws�ΏۃV�[�g)
                End If
                
            Next i
            
        End If
        
        ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
        Dim �t�@�C��CLOSE���@�敪�l As Long
        �t�@�C��CLOSE���@�敪�l = �u�b�NCLOSE�O����(txt�p�X, wb�Ώۃu�b�N)
        
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
    Call ���s���ʓ��e�ҏW����(obj���ʏo�̓V�[�g.�o�͔z��)

    Dim wb���ʃu�b�N As Workbook
    Set wb���ʃu�b�N = obj���ʏo�̓V�[�g.���ʃu�b�N�쐬(txt�p�X���ʕ�)

    If obj���ʏo�̓V�[�g.�o�͂��� Then
    
        ' �������������̏����̌Ăяo���i�Ăяo�����Procedure���ł̓c�[�����Ƃ̌ŗL�̎������s���j
        Call ���s���ʏ����ҏW����(wb���ʃu�b�N.ActiveSheet)
        wb���ʃu�b�N.Activate
        
        Call obj���ʏo�̓V�[�g.�K�v�ɉ����ĕۑ�(obj�ݒ�l�V�[�g)

    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Application.StatusBar = False

    ' �t�@�C���̕ۑ��`�������̏�Ԃɖ߂�
    Application.defaultSaveFormat = defaultSaveFormat

End Function
