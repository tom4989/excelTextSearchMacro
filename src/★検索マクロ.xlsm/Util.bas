Attribute VB_Name = "Util"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' �萔
' ---------------------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------------------
' �ϐ�
' ---------------------------------------------------------------------------------------------------------------------

Dim var�J�n���� As Variant

' #####################################################################################################################
' #
' # ���O�n���[�e�B���e�B
' #
' #####################################################################################################################

Sub log(ByVal str���b�Z�[�W As String)

    Debug.Print Format(Now(), "HH:mm:ss ") & str���b�Z�[�W

End Sub

Function getTimestamp()

    getTimestamp = Format(Now(), "yyyymmdd_HHnnss")

End Function

' #####################################################################################################################
' #
' # ���b�Z�[�W�n���[�e�B���e�B
' #
' #####################################################################################################################

Function get�J�n���b�Z�[�W(ByVal txt������ As String) As String

    Call init�J�n����

    Dim txt���b�Z�[�W As String
    txt���b�Z�[�W = Format(Now(), "HH:mm:ss ") & txt������ & "�������J�n���܂��B"

    Debug.Print txt���b�Z�[�W
    get�J�n���b�Z�[�W = txt���b�Z�[�W

End Function

Function get�I�����b�Z�[�W(ByVal txt������ As String) As String

    Dim txt���b�Z�[�W As String
    txt���b�Z�[�W = Format(Now(), "HH:mm:ss ") & txt������ & "�������I�����܂����B�i�������ԁF" & get�������� & "�j"

    Debug.Print txt���b�Z�[�W
    get�I�����b�Z�[�W = txt���b�Z�[�W

End Function

Function get�ُ펞���b�Z�[�W(ByVal txt������ As String) As String

    Dim txt���b�Z�[�W As String
    txt���b�Z�[�W = Format(Now(), "HH:mm:ss ") & txt������ & "�������I�����܂����B�i�������ԁF" & get�������� & "�j"

    Debug.Print txt���b�Z�[�W
    get�ُ펞���b�Z�[�W = txt���b�Z�[�W
End Function

' #####################################################################################################################
' #
' # �X�e�[�^�X�o�[����n���[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * �@�\�@�F�X�e�[�^�X�o�[�ɕ\�����鏈�����Ԃ�����������
' *********************************************************************************************************************
'
Sub init�J�n����()

    var�J�n���� = Now()
    
End Sub

' *********************************************************************************************************************
' * �@�\�@�F�������Ԃ̊J�n�������擾����
' *********************************************************************************************************************
'
Function get�J�n����()

    get�J�n���� = var�J�n����

End Function

' *********************************************************************************************************************
' * �@�\�@�F�������Ԃ� HH:mm:ss �`���Ŏ擾����
' *********************************************************************************************************************
'
Function get��������()

    get�������� = Format(Now() - var�J�n����, "HH:mm:ss")
    
End Function

' *********************************************************************************************************************
' * �@�\�@�F�X�e�[�^�X�o�[�Ɍo�ߎ��ԕt�Ń��b�Z�[�W��\������
' *********************************************************************************************************************
'
Sub set�X�e�[�^�X�o�[(ByVal str���b�Z�[�W As String)

    If IsEmpty(var�J�n����) Then
        
        var�J�n���� = Now()
        
    End If
    
    Application.StatusBar = get��������() & " " & str���b�Z�[�W

End Sub

' #####################################################################################################################
' #
' # �u�b�N�A�V�[�g����n���[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * �@�\�@�F�����œn���ꂽ�V�[�g���ȊO�̃V�[�g���폜����
' *********************************************************************************************************************
'
Function �s�v�V�[�g�폜(�Ώۃu�b�N��� As Workbook, ByVal �c���V�[�g�� As String)

    Dim �O��� As Boolean
    �O��� = Application.DisplayAlerts
    
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    
    For Each ws In �Ώۃu�b�N���.Worksheets
    
        If ws.Name <> �c���V�[�g�� Then
            Worksheets(ws.Name).Delete
        End If
        
    Next ws
    
    Application.DisplayAlerts = �O���
        
End Function

' *********************************************************************************************************************
' * �@�\�@�F�����œn���ꂽ�V�[�g�̍ŏI�s���擾����
' *********************************************************************************************************************
'
Function �ŏI�s�擾(ws�ΏۃV�[�g As Worksheet, Optional useUsedRange As Boolean = True) As Long

    If useUsedRange Then

        With ws�ΏۃV�[�g.UsedRange
            �ŏI�s�擾 = .Rows(.Rows.Count).Row
        End With
        
    Else
        With ws�ΏۃV�[�g
            �ŏI�s�擾 = .Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        End With
        
    End If
        
End Function


' *********************************************************************************************************************
' * �@�\�@�F�����œn���ꂽ�V�[�g�̍ŏI����擾����
' *********************************************************************************************************************
'
Function �ŏI��擾(ws�ΏۃV�[�g As Worksheet, Optional useUsedRange As Boolean = True) As Long

    If useUsedRange Then

        With ws�ΏۃV�[�g.UsedRange
            �ŏI��擾 = .Columns(.Columns.Count).Column
        End With
        
    Else
        With ws�ΏۃV�[�g
            �ŏI��擾 = .Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
        End With
    
    End If
        
End Function

' *********************************************************************************************************************
' * �@�\�@�F�����œn���ꂽ�V�[�g�̓��e��Variant�ϐ��ɕϊ����ĕԂ�
' *********************************************************************************************************************
'
Function �V�[�g���e�擾(ws���[�N�V�[�g As Worksheet) As Variant

    With ws���[�N�V�[�g

        �V�[�g���e�擾 = .Range( _
            .Cells(1, 1), _
            .Cells(�ŏI�s�擾(ws���[�N�V�[�g), �ŏI��擾(ws���[�N�V�[�g)))
    End With

End Function


' #####################################################################################################################
' #
' # �_�C�A���O����n���[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * �@�\�@�F�������s or ���~�m�F�_�C�A���O��\������
' *********************************************************************************************************************
'
Function �������s���f(message As String)

    Dim rc As VbMsgBoxResult
    rc = MsgBox(message + Chr(10) + "�����𑱍s���܂����H", vbYesNo, vbQuestion)
    
    If rc = vbYes Then
        MsgBox "�����𑱂��܂�", vbInformation
    Else
        MsgBox "�����𒆎~���܂����B", vbCritical
        
        ' �}�N���̎��s���~
        End
    End If

End Function


' #####################################################################################################################
' #
' # �I�[�g�V�F�C�v����n���[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' �@�\���F�ΏۃV�[�g��ɂ���I�u�W�F�N�g�̃v���p�e�B���擾����
' �߂�@�FgetShapesProperty as String(2, n)
'         (0, n)   type
'         (1, n)   name
'         (2, n)   TextFrame.Characters.text
'         (3, n)   Left
'         (4, n)   Top
'         (5, n)   Width
'         (6, n)   Height
'         (7, n)   TopLeftCell.Address(False, False)
'         (8, n)   TopLeftCell.row
'         (9, n)   TopLeftCell.Column
'         (10, n)  BottomRightCell.Address(False, False)
'         (11, n)  BottomRightCell.row
'         (12, n)  BottomRightCell.Column
'
' *********************************************************************************************************************
'
Function getShapesProperty(ByRef targetSheet As Worksheet, Optional ByVal objType As Long = -999, Optional ByVal formCtlType As Long = -999) As Variant

    Dim ret As Variant
    
    Dim i As Long
    Dim obj As Variant
    
    ' �z��̍쐬�B
    i = 0
    For Each obj In targetSheet.Shapes
        ' FORM�R���g���[���̏ꍇ
        If obj.Type = objType Then
            ' �n���ꂽ�t�H�[���R���g���[���^�C�v����v�����ꍇ�A�J�E���g�A�b�v
            If obj.FormControlType = formCtlType Then
                i = i + 1
            End If
            
            ' �w��Ȃ����́A����ȊO�̃I�[�g�V�F�C�v
            ElseIf objType = -999 Or obj.Type = objType Then
                i = i + 1
            End If
    Next
        
    ' �Ώۂ̃I�[�g�V�F�C�v���݂������ꍇ�̂݁A���̃I�u�W�F�N�g�̊i�[���s���B
    If 0 <> i Then
        ReDim ret(i - 1, 12)
        
        ' �z��̍쐬
        i = 0
        ' �I�u�W�F�N�g���̐ݒ�
        For Each obj In targetSheet.Shapes
            
            ' form�R���g���[���̏ꍇ
            If obj.Type = objType Then
                ' �n���ꂽ�t�H�[���Rr���g���[���^�C�v����v�����ꍇ�A�l���擾����B
                If obj.FormControlType = formCtlType Then
                        
                    ret(i, 0) = obj.Type
                    ret(i, 1) = obj.AlternativeText
                        
                    ' TextFrame�v���p�e�B���g�p�ł��Ȃ��i���C�A�E�g�g���Ȃ��j�I�u�W�F�N�g�͏��O
                    On Error Resume Next
                    ret(i, 2) = obj.ControlFormat.Value
                    ret(i, 3) = obj.Left
                    ret(i, 4) = obj.Top
                    ret(i, 5) = obj.Width
                    ret(i, 6) = obj.Height
                    ret(i, 7) = obj.TopLeftCell.Address(False, False)
                    ret(i, 8) = obj.TopLeftCell.Row
                    ret(i, 9) = obj.TopLeftCell.Column
                    ret(i, 10) = obj.Left.BottomRightCell.Address(False, False)
                    ret(i, 11) = obj.Left.BottomRightCell.Row
                    ret(i, 12) = obj.Left.BottomRightCell.Column
                        
                    i = i + 1
                End If
                    
            ' �w��Ȃ����́A����ȊO�̃I�[�g�V�F�C�v�Ȃǂ̏ꍇ
            ElseIf objType = -999 Or obj.Type = objType Then
                
                ret(i, 0) = obj.Type
                ret(i, 1) = obj.AlternativeText
                        
                ' TextFrame�v���p�e�B���g�p�ł��Ȃ��i���C�A�E�g�g���Ȃ��j�I�u�W�F�N�g�͏��O
                On Error Resume Next
                ret(i, 2) = obj.TextFrame.Characters.Text
                    
                ret(i, 3) = obj.Left
                ret(i, 4) = obj.Top
                ret(i, 5) = obj.Width
                ret(i, 6) = obj.Height
                ret(i, 7) = obj.TopLeftCell.Address(False, False)
                ret(i, 8) = obj.TopLeftCell.Row
                ret(i, 9) = obj.TopLeftCell.Column
                ret(i, 10) = obj.Left.BottomRightCell.Address(False, False)
                ret(i, 11) = obj.Left.BottomRightCell.Row
                ret(i, 12) = obj.Left.BottomRightCell.Column
                       
                i = i + 1
            End If
        Next
    End If
        
    getShapesProperty = ret
    
End Function

' #####################################################################################################################
' #
' # �z�񑀍�n���[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' �@�\�@�F�������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' �����@�FvarArray �z��
' �߂�l�F���茋�ʁi1:�z��/0:��̔z��/-1�F�z�񂶂�Ȃ�)
' *********************************************************************************************************************
'
Public Function IsArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If
    
    Exit Function
    
ERROR_:
    If Err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

' *********************************************************************************************************************
' �@�\�@�F�V�[�g���e���i�[�����z�񂩂���W(A1��)���w�肵�Ēl���擾����
' *********************************************************************************************************************
'
Public Function f_�Z�����W�̒l�擾(ByRef var�z�� As Variant, txt�Z�����W As String) As String

    Dim var���W As Variant
    var���W = CAlpNum2Num(txt�Z�����W)

    func�Z�����W�̒l�擾 = var�z��(var���W(0), var���W(1))

End Function

' *********************************************************************************************************************
' �@�\�@�F�Ώۂ̔z��ɁA�w�肳�ꂽ�����񂪊i�[����Ă��邩���肷��
' *********************************************************************************************************************
'
Public Function containArray(var�z�� As Variant, txt�Ώە����� As String) As Boolean

    Dim i As Long
    
    For i = LBound(var�z��) To UBound(var�z��)
    
        If var�z��(i) = txt�Ώە����� Then
        
            containArray = True
            Exit Function
        
        End If
    Next i
    
    containArray = False

End Function

' *********************************************************************************************************************
' �@�\�F�Ώۂ̔z��ɁA�w�肳�ꂽ�����񂪊i�[����Ă��邩���肷��
' *********************************************************************************************************************
'
Public Function f_�z��܂܂�Ă��邩�`�F�b�N( _
    var�z�� As Variant, txt�Ώە����� As String, _
    Optional flg������Ƀ��C���h�J�[�h���� As Boolean = False, _
    Optional flg�z��Ƀ��C���h�J�[�h���� As Boolean = False) As Boolean

    Dim i As Long
    
    If IsArrayEx(var�z��) <> 1 Then
    
        f_�z��܂܂�Ă��邩�`�F�b�N = False
        Exit Function
    End If
    
    For i = LBound(var�z��) To UBound(var�z��)
    
        If var�z��(i) = txt�Ώە����� Then
        
            f_�z��܂܂�Ă��邩�`�F�b�N = True
            Exit Function
        End If
        
        If flg������Ƀ��C���h�J�[�h���� Then
        
            If var�z��(i) Like txt�Ώە����� Then
            
                f_�z��܂܂�Ă��邩�`�F�b�N = True
                Exit Function
            End If
        End If
        
        If flg�z��Ƀ��C���h�J�[�h���� Then
        
            If txt�Ώە����� Like var�z��(i) Then
            
                f_�z��܂܂�Ă��邩�`�F�b�N = True
                Exit Function
            End If
        End If
        
    Next i
    
    f_�z��܂܂�Ă��邩�`�F�b�N = False

End Function

' *********************************************************************************************************************
' �@�\�@�F���s���ʂ�ێ�����񎟌��z��ϐ����`����Function
' *********************************************************************************************************************
'
Function reDimResult(ByVal topLevelElementSize As Integer, ByRef results() As Variant)

    Select Case IsArrayEx(results)
        Case 1
            ' results���������ς̏ꍇ
            ' ���݂̃��R�[�h�� + 1�s�̈���m��
            ReDim Preserve results(topLevelElementSize, UBound(results, 2) + 1)
        Case 0
            ' results��1�x������������Ă��Ȃ��ꍇ
            ' 1�s�̈���m��
            ReDim Preserve results(topLevelElementSize, 0)
    End Select
        
End Function

' *********************************************************************************************************************
' �@�\�@�F�ꎟ���z��ɐV���ȗv�f��ǉ�����
' *********************************************************************************************************************
'
Function �ꎟ���z��ɒl��ǉ�(ByRef valueList As Variant, ByVal �ǉ��ݒ�l As String)

    ' �t�@�C�������擾����
    Select Case IsArrayEx(valueList)
        Case 1
            ReDim Preserve valueList(UBound(valueList) + 1)
        Case 0
            ReDim Preserve valueList(0)
    End Select
    
    ' �ǉ��������X�g�ɁA�ݒ�l���i�[�B
    valueList(UBound(valueList)) = �ǉ��ݒ�l
    
End Function

' *********************************************************************************************************************
' �@�\�@�F�񎟌��z��̍s�Ɨ�����ւ���
' *********************************************************************************************************************
'
Function �񎟌��z��s��t�](ByRef var�񎟌��z�� As Variant)

    Dim var�t�]��z�� As Variant
    
    ReDim var�t�]��z��( _
        LBound(var�񎟌��z��, 2) To UBound(var�񎟌��z��, 2), _
        LBound(var�񎟌��z��) To UBound(var�񎟌��z��))
        
    Dim i, j As Long
    
    For i = LBound(var�񎟌��z��) To UBound(var�񎟌��z��, 2)
        
        For j = LBound(var�񎟌��z��) To UBound(var�񎟌��z��)
            
            var�t�]��z��(i, j) = var�񎟌��z��(j, i)
            
        Next
    Next
    
    �񎟌��z��s��t�] = var�t�]��z��
        
    
End Function


' #####################################################################################################################
' #
' # �����n���[�e�B���e�B
' #
' #####################################################################################################################
    
' *********************************************************************************************************************
' �@�\�@�F�ΏۃZ���͈͓��Ō���������ɊY�������������ԑ������ɂ���
' *********************************************************************************************************************
'
Function �����Y�������̐ԑ�������(prmRange As Range, prmTargetString As String)

    Dim txt As String
    Dim i, m As Integer
    Dim targetRange As Range
    
    If prmTargetString = "" Then
        Exit Function
    End If

    For Each targetRange In prmRange
        txt = targetRange.Value
        m = Len(prmTargetString)
        i = InStr(1, txt, prmTargetString)
        Do Until i = 0
            With prmRange.Characters(i, m)
                .Font.Bold = True
                .Font.ColorIndex = 3
            End With
            i = InStr(i + 1, txt, prmTargetString)
        Loop
    Next
    
    Set targetRange = Nothing
    
End Function

' #####################################################################################################################
' #
' # �V�[�g���擾�n���[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' �@�\�@�F�^�C�g�����w��Ń��X�g�l���擾
'         �����X�g�l���Ȃ������ꍇ�A�z��̗v�f��1�i�l�͋�j���ԋp����܂��B
' *********************************************************************************************************************
'
Function �^�C�g�����w��Ń��X�g�l���擾(titleName As String, targetSheet As Worksheet) As Variant

    Dim targetRangeList As Range
    Dim targetVariantList As Variant
    
    Set targetRangeList = �^�C�g�����w��Ń��X�g�l��Range�����擾(titleName, targetSheet)
    ' �z�񂩔���
    If targetRangeList.Count = 1 Then
        targetVariantList = Array(targetRangeList.Item(1).Value)
    Else
        targetVariantList = targetRangeList.Value
    End If
    
    �^�C�g�����w��Ń��X�g�l���擾 = targetVariantList
    
End Function

' *********************************************************************************************************************
' �@�\�@�F�^�C�g�����w��Ń��X�g�l��Range�����擾
'         �����X�g�l���Ȃ������ꍇ�A���X�g�l�G���A��1�s�ځi�l�͋�j��Range��񂪕ԋp����܂��B
' *********************************************************************************************************************
'
Function �^�C�g�����w��Ń��X�g�l��Range�����擾(titleName As String, targetSheet As Worksheet) As Range

    ' �����q�b�g��
    Dim matchCount As Long
    Dim checkValue As String
    
    ' �V�[�g���Ƀ^�C�g�����������ݒ肳��Ă��Ȃ������m�F����B
    matchCount = WorksheetFunction.CountIf(targetSheet.UsedRange, titleName)
    If 1 <> matchCount Then
        MsgBox "�^�C�g���u" & titleName & "�v�����������������߁A�����𒆒f���܂����B"
        End
    End If
    
    ' �^�C�g������Range�����擾
    Dim FoundCell As Range
    Set FoundCell = targetSheet.UsedRange.Find(what:=titleName, LookIn:=xlValues, _
        LookAt:=xlPart, MatchCase:=False, MatchByte:=False)
    Dim i, MaxRow, MaxCol As Long
    
    ' �^�C�g���ɑ΂��郊�X�g�l���擾�i�󔒍s���݁j
    With targetSheet
        With .Range(.Cells(FoundCell.Row, FoundCell.Column), .Cells(Rows.Count, FoundCell.Column))
            MaxRow = .Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
            MaxCol = .Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
        End With
    
        ' MaxRow���A�󔒍s����s���̃��X�g�l�̍s���ɐݒ肷��B
        For i = 1 To (MaxRow - FoundCell.Row)
            checkValue = .Cells(FoundCell.Row + i, MaxCol).Value
            If "" = checkValue Or InStr(1, checkValue, TITLE_NAME_PREFIX) > 0 Then
                If 1 = i Then
                    Call �������s���f("�^�C�g�����u" + titleName + "�v�ɑ΂��郊�X�g�l���ݒ肳��Ă��܂���B")
                    MaxRow = FoundCell.Row + 1
                Else
                    MaxRow = FoundCell.Row + i - 1
                End If
                    Exit For
            End If
        Next
        
        ' ���X�g�l��ԋp
        Set �^�C�g�����w��Ń��X�g�l��Range�����擾 = _
            targetSheet.Range(.Cells((FoundCell.Row + 1), FoundCell.Column), .Cells(MaxRow, MaxCol))
        
    End With

End Function

' *********************************************************************************************************************
' �@�\�@�F�����Ŏw�肳�ꂽ�s���I����Ԃł��邩���肷��
' *********************************************************************************************************************
'
Function is�I�����(ByVal lng�Ώۍs As Long)

    Dim rng As Range
    
    For Each rng In Selection.Rows
    
        If rng.Row = lng�Ώۍs Then
        
            is�I����� = True
            Exit Function
            
        End If
        
    Next rng
        
    is�I����� = False


End Function

' *********************************************************************************************************************
' �@�\�F��̒l�𐔎��ɕϊ�����
' *********************************************************************************************************************
'
Function CAlp2Num(txtAlphabet As String) As Long
  
    CAlp2Num = ActiveSheet.Range(txtAlphabet & "1").Column
    
End Function


' *********************************************************************************************************************
' �@�\�F��(A:B�`��)�̒l�𐔎��ɕϊ�����
' *********************************************************************************************************************
'
Function CAlpxAlp2Num(txtAlpxAlp As String) As Variant
  
    Dim var���� As Variant

    var���� = Split(txtAlpxAlp, ":")

    var����(0) = CAlp2Num(CStr(var����(0)))
    
    If UBound(var����) >= 1 Then
        var����(1) = CAlp2Num(CStr(var����(1)))
    Else
        Call �ꎟ�z��ɒl��ǉ�(var����, var����(0))
    End If

    CAlpxAlp2Num = var����
    
End Function

' *********************************************************************************************************************
' �@�\�F�Z�����W�𐔎��ɕϊ�����
' *********************************************************************************************************************
'
Function CAlpNum2Num(txt���W As String) As Variant

    Dim var���� As Variant
    ReDim var����(1)

    Dim objReg As Object, objMatch As Object
    
    Set objReg = CreateObject("VBScript.RegExp")
    objReg.Pattern = "^([A-Z]+)([0-9]+)$"

    Set objMatch = objReg.Execute(txt���W)

    Dim txt�s As String, txt�� As String
    
    var����(0) = CLng(CAlp2Num(objMatch(0).SubMatches(0)))
    var����(1) = CLng(objMatch(0).SubMatches(1))

    CAlpNum2Num = var����

End Function

' #####################################################################################################################
' #
' # �����񃆁[�e�B���e�B
' #
' #####################################################################################################################

' *********************************************************************************************************************
' �@�\�F�����Ŏw�肳�ꂽ������̋��ʕ���Ԃ�
' *********************************************************************************************************************
'
Public Function f_���ʕ��擾(txt������1 As String, txt������2 As String) As String

    Dim lng������ As Long
    
    If Len(txt������1) <= Len(txt������2) Then
    
        lng������ = Len(txt������1)
    Else
        lng������ = Len(txt������2)
    
    End If
    
    Dim i As Long
    
    For i = 1 To lng������
    
        If Left(txt������1, i) <> Left(txt������2, i) Then
        
            Exit For
            
        End If
        
    Next i
    
    If i > 1 Then
    
        f_���ʕ��擾 = Left(txt������1, i - 1)
    
    Else
        f_���ʕ��擾 = ""
    
    End If
    
End Function

' *********************************************************************************************************************
' �@�\�F�w�蕶����RTRIM
' *********************************************************************************************************************
'
Function f_RTRIM(txt�Ώە����� As String, txt�w�蕶�� As String) As String

    If txt�Ώە����� <> "" Then

        If Right(txt�Ώە�����, 1) = txt�w�蕶�� Then
    
            f_RTRIM = Left(txt�Ώە�����, Len(txt�Ώە�����) - 1)
            Exit Function
    
        End If
        
    End If
    
    f_RTRIM = txt�Ώە�����
    
End Function
