Attribute VB_Name = "FileOperationUtil"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' �萔
' ---------------------------------------------------------------------------------------------------------------------

'�p�X�̃f���~�^
Public Const PATH_DELIMITER = "\"

' �t�@�C�����A�g���q�̃f���~�^
Public Const FILE_DELIMITER = "."

' �t�@�C��������
Public Type �t�@�C��������
    �t���p�X_�t�@�C���� As String
    �t���p�X As String
    �e�f�B���N�g���܂ł̃t���p�X As String
    �Ώۃf�B���N�g���� As String
    �Ώۃt�@�C���� As String
    �Ώۃt�@�C�����() As String
End Type

' �Ώۂ̊g���q�i���W���[���j
Public Const FILE_EXTENSION_OF_MODULE = "bas,cls,frm"

' �t�@�C��CLOSE��ԋ敪
Public Enum �t�@�C��CLOSE���@�敪
    �ۑ����Ȃ��ŕ��� = 0
    �ۑ����ĕ��� = 1
    �ۑ����Ȃ��ŕ��Ȃ� = 2
    �ۑ����ĕ��Ȃ� = 3
    �������f = 99
End Enum

' ---------------------------------------------------------------------------------------------------------------------
' �ϐ�
' ---------------------------------------------------------------------------------------------------------------------

Private objFSO As Object

' ���[�g�p�X�쐬�σt���O
Private rootPathMaked As Boolean

' #####################################################################################################################
' #
' # �A�N�Z�T
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * �@�\�@�F���[�g�p�X�쐬�ς݃t���O��ݒ肷��
' *********************************************************************************************************************
'
Public Function getRootPathMaked() As Boolean
    getRootPathMaked = rootPathMaked
End Function

' *********************************************************************************************************************
' * �@�\�@�F���[�g�p�X�쐬�ς݃t���O���擾����
' *********************************************************************************************************************
'
Public Function setRootPathMaked(isMaked As Boolean)
    rootPathMaked = isMaked
End Function

' #####################################################################################################################
' #
' # �t�@�C�����샆�[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * �@�\�@�F�t�@�C�����擾
' *********************************************************************************************************************
'
Function �t�@�C�����擾(txt�p�X As String) As String

    If objFSO Is Nothing Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    End If

    objFSO.GetFileName (txt�p�X)

End Function

' *********************************************************************************************************************
' * �@�\�@�F�t�@�C�����ɃT�t�B�b�N�X��t�^����B
' *********************************************************************************************************************
'
Function f_�����T�t�B�b�N�X�t�^(txt�p�X As String, txt�T�t�B�b�N�X As String) As String

    If txt�T�t�B�b�N�X <> "" Then
        
        Dim txt�g���q As String
        txt�g���q = Mid(txt�p�X, InStrRev(txt�p�X, "."))
        
        f_�����T�t�B�b�N�X�t�^ = Replace(txt�p�X, txt�g���q, Format(Now, txt�T�t�B�b�N�X) & txt�g���q)
        
    Else
    
        f_�����T�t�B�b�N�X�t�^ = txt�p�X
        
    End If

End Function

' *********************************************************************************************************************
' * �@�\�@�F�p�X�i�p�X���t�@�C���j�̑��݃`�F�b�N
' * �����@�FdirectoryPath �p�X�i�܂��́A�p�X���t�@�C���j
' * �߂�l�F�`�F�b�N���ʁi�p�X���ݎ���1�A�t�@�C�����ݎ���2�A�p�X���t�@�C�������݂��Ȃ��ꍇ��-1�j
' *********************************************************************************************************************
'
Function isDirectoryExist(directoryPath As String) As Long

    If objFSO Is Nothing Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    End If
    
    If True = objFSO.FileExists(directoryPath) Then
        isDirectoryExist = 2
    ElseIf True = objFSO.FolderExists(directoryPath) Then
        isDirectoryExist = 1
    Else
        isDirectoryExist = -1
    End If
        
End Function

' *********************************************************************************************************************
' * �@�\�@�F�t�H���_�����݂��Ȃ�������쐬����
' *********************************************************************************************************************
'
Function mkdirIFNotExist(txt�t�H���_�� As String)

    If Dir(txt�t�H���_��, vbDirectory) = "" Then
        mkdir txt�t�H���_��
    End If

End Function

' *********************************************************************************************************************
' * �@�\�@�F�p�X�z���̊K�w�S�Ẵf�B���N�g������������
' * �����@�FdirectoryPath �p�X
' * �߂�l�F���s���ʁi�J�����g�f�B���N�g�����܂ށA�z���̃f�B���N�g�����̔z��
' *********************************************************************************************************************
'
Function doRepeat(ByVal directoryPath As String, ByVal fileExtensions As Variant, _
    ByRef fileNames() As String, Optional ByVal recursive As Boolean = False)
    
    ' ��������
    Dim buf As String, msg As String, dirName As Variant
    
    ' �z���̃p�X���
    Dim directoryPathBySub As String
    directoryPathBySub = directoryPath
    
    ' �����̃f�B���N�g�����݉ۃt���O
    Dim isExistDir As Boolean
    isExistDir = False
    
    Dim dirNames() As String
    
    Dim resultArray As Variant
    
    If "" <> directoryPath Then
        ' �f�B���N�g���ړ�
        ChDir directoryPath
        
        ' -------------------------------------------------------------------------------------------------------------
        ' �����̃t�@�C������S�Ď擾
        ' -------------------------------------------------------------------------------------------------------------
        Call getFileNames(directoryPath, fileExtensions, fileNames)
        
        If recursive Then
        
            ' ---------------------------------------------------------------------------------------------------------
            ' �����̃f�B���N�g������S�Ď擾
            ' ---------------------------------------------------------------------------------------------------------
            dirNames = getDirNames(directoryPath)
            
            ' ---------------------------------------------------------------------------------------------------------
            ' �擾�����f�B���N�g����1���ċA�I�ɏ�������B
            ' ---------------------------------------------------------------------------------------------------------
            If Not Not dirNames Then
                For Each dirName In dirNames
                    Call doRepeat(dirName, fileExtensions, fileNames, True)
                Next
            End If
            
        End If
        
    End If
    
End Function

' *********************************************************************************************************************
' * �@�\�@�F�p�X�����̃t�@�C������S�Ď擾
' * �����@�FdirectoryPath �p�X
' * �߂�l�F���s���ʁi�J�����g�f�B���N�g�������̃f�B���N�g�����̔z��B�j
' *********************************************************************************************************************
'
Function getFileNames(directoryPath As String, fileExtensions As Variant, ByRef fileNames() As String)

    Dim fileName As String, msg As String
    
    Dim fileNameSize As Integer
    
    Dim fileExtension As Variant
    
    ' �f�B���N�g���ړ�
    ChDir directoryPath
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' �����̃t�@�C������S�Ď擾
    ' -----------------------------------------------------------------------------------------------------------------
    fileName = Dir(directoryPath & "\" & "*.*")
    
    Do While fileName <> ""
    
        ' �t�@�C�����擾
        For Each fileExtension In fileExtensions
            If InStr(1, UCase(fileName), UCase(fileExtension)) > 0 Then
            
                ' �t���p�X���t�@�C������ǉ��i�[�B
                Call �ꎟ���z��ɒl��ǉ�(fileNames, directoryPath & "\" & fileName)
                Exit For
        
            End If
        Next
    
        fileName = Dir()
    Loop

End Function

' *********************************************************************************************************************
' * �@�\�@�F�p�X�����̃f�B���N�g������S�Ď擾
' * �����@�FdirectoryPath �p�X
' * �߂�l�F���s���ʁi�J�����g�f�B���N�g�������̃f�B���N�g�����̔z��B�j
' *********************************************************************************************************************
'
Function getDirNames(directoryPath As String) As String()
    Dim buf As String
    
    Dim dirNames() As String
    
    ' �f�B���N�g���ړ�
    ChDir directoryPath
    
    buf = Dir(directoryPath & "\" & "*.*", vbDirectory)
    Do While buf <> ""
        ' �f�B���N�g�����擾
        If GetAttr(directoryPath & "\" & buf) And vbDirectory Then
            If buf <> "." And buf <> ".." Then
            
                ' �f�B���N�g������ǉ��i�[�B
                Call �ꎟ���z��ɒl��ǉ�(dirNames, directoryPath & "\" & buf)
                
            End If
        End If
        buf = Dir()
        
    Loop
    
    getDirNames = dirNames

End Function


' *********************************************************************************************************************
' * �@�\�@�F�Ώۃf�B���N�g�����쐬����i�Ώۃp�X�������݁A�쐬�f�B���N�g���������݂����ꍇ�͏������f�j
' *********************************************************************************************************************
'
Function �f�B���N�g���쐬(ByVal ���[�g�p�X As String, ByVal �������� As String, ByVal ���΃p�X As String)

    Dim dirCheck As Long
    
    ' ���[�g�p�X�̑��݃`�F�b�N
    dirCheck = isDirectoryExist(CStr(���[�g�p�X & ��������))
    ' �Ώۃp�X�����ݒ�̏ꍇ�i���[�g�p�X�쐬���j
    If "" = ���΃p�X Then
        ' �����������ݒ�ς̏ꍇ�A���[�g�p�X���쐬�ςł���Ώ������f�Ƃ���
        If "" <> �������� Then
            If 0 < dirCheck And False = getRootPathMaked() Then
                MsgBox "�ȉ��̃f�B���N�g���͊��ɑ��݂��邽�ߏ����𒆒f���܂��B" + Chr(10) + "�u" + ���[�g�p�X & �������� + "�v"
                End
            End If
        End If
    End If

    ' �f�B���N�g���쐬
    dirCheck = isDirectoryExist(CStr(���[�g�p�X & �������� & PATH_DELIMITER & ���΃p�X))
    If 0 > dirCheck Then
        mkdir ���[�g�p�X & �������� & PATH_DELIMITER & ���΃p�X
        Call setRootPathMaked(True)
    End If
    
End Function








