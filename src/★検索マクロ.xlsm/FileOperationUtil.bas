Attribute VB_Name = "FileOperationUtil"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

' ファイルCLOSE状態区分
Public Enum ファイルCLOSE方法区分
    保存しないで閉じる = 0
    保存して閉じる = 1
    保存しないで閉じない = 2
    保存して閉じない = 3
    処理中断 = 99
End Enum

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------

Private objFSO As Object

' #####################################################################################################################
' #
' # ファイル操作ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * 機能　：FileSystemObjectの初期化
' *********************************************************************************************************************
'
Private Sub subFSO初期化()

    If objFSO Is Nothing Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    End If

End Sub

' *********************************************************************************************************************
' * 機能　：ファイル名取得
' *********************************************************************************************************************
'
Function ファイル名取得(txtパス As String) As String

    Call subFSO初期化
    objFSO.GetFileName (txtパス)

End Function


' *********************************************************************************************************************
' * 機能　：フォルダ名取得
' *********************************************************************************************************************
'
Function funフォルダ名取得(txtパス As String) As String

    Call subFSO初期化
    
    If objFSO.FolderExists(txtパス) Then
    
        funフォルダ名取得 = txtパス
        Exit Function
    
    End If
    
    funフォルダ名取得 = objFSO.getParentFolderName(txtパス)

End Function

' *********************************************************************************************************************
' * 機能　：ファイル名にサフィックスを付与する。
' *********************************************************************************************************************
'
Function f_日時サフィックス付与(txtパス As String, txtサフィックス As String) As String

    If txtサフィックス <> "" Then
        
        Dim txt拡張子 As String
        txt拡張子 = Mid(txtパス, InStrRev(txtパス, "."))
        
        f_日時サフィックス付与 = Replace(txtパス, txt拡張子, Format(Now, txtサフィックス) & txt拡張子)
        
    Else
    
        f_日時サフィックス付与 = txtパス
        
    End If

End Function

' *********************************************************************************************************************
' * 機能　：パス（パス＆ファイル）の存在チェック
' * 引数　：directoryPath パス（または、パス＆ファイル）
' * 戻り値：チェック結果（パス存在時は1、ファイル存在時は2、パスもファイルも存在しない場合は-1）
' *********************************************************************************************************************
'
Function isDirectoryExist(directoryPath As String) As Long

    Call subFSO初期化
    
    If True = objFSO.FileExists(directoryPath) Then
        isDirectoryExist = 2
    ElseIf True = objFSO.FolderExists(directoryPath) Then
        isDirectoryExist = 1
    Else
        isDirectoryExist = -1
    End If
        
End Function

' *********************************************************************************************************************
' * 機能　：フォルダが存在しなかったら作成する
' *********************************************************************************************************************
'
Function mkdirIFNotExist(txtフォルダ名 As String)

    If Dir(txtフォルダ名, vbDirectory) = "" Then
        mkdir txtフォルダ名
    End If

End Function

' *********************************************************************************************************************
' * 機能　：エラーログの出力を行う
' *********************************************************************************************************************
'
Function subエラーログファイル出力(txtエラー内容 As String)

    Call subFSO初期化

    Dim txtエラーログファイルパス As String
    txtエラーログファイルパス = ThisWorkbook.Path & "\log"

    mkdirIFNotExist (txtエラーログファイルパス)
    txtエラーログファイルパス = txtエラーログファイルパス & "\" & ThisWorkbook.Name & ".log"
    
    With objFSO.OpenTextFile(txtエラーログファイルパス, 8, True, -2)

        .WriteLine Now & vbCrLf & txtエラー内容
        .Close

    End With

End Function

' *********************************************************************************************************************
' * 機能　：パス配下の階層全てのディレクトリを処理する
' * 引数　：directoryPath パス
' * 戻り値：実行結果（カレントディレクトリを含む、配下のディレクトリ名の配列
' *********************************************************************************************************************
'
Function doRepeat(ByVal directoryPath As String, ByVal fileExtensions As Variant, _
    ByRef fileNames() As String, Optional ByVal recursive As Boolean = False)
    
    ' 検索結果
    Dim buf As String, msg As String, dirName As Variant
    
    ' 配下のパス情報
    Dim directoryPathBySub As String
    directoryPathBySub = directoryPath
    
    ' 直下のディレクトリ存在可否フラグ
    Dim isExistDir As Boolean
    isExistDir = False
    
    Dim dirNames() As String
    
    Dim resultArray As Variant
    
    If "" <> directoryPath Then
        ' ディレクトリ移動
        ChDir directoryPath
        
        ' -------------------------------------------------------------------------------------------------------------
        ' 直下のファイル名を全て取得
        ' -------------------------------------------------------------------------------------------------------------
        Call getFileNames(directoryPath, fileExtensions, fileNames)
        
        If recursive Then
        
            ' ---------------------------------------------------------------------------------------------------------
            ' 直下のディレクトリ名を全て取得
            ' ---------------------------------------------------------------------------------------------------------
            dirNames = getDirNames(directoryPath)
            
            ' ---------------------------------------------------------------------------------------------------------
            ' 取得したディレクトリ名1つずつ再帰的に処理する。
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
' * 機能　：パス直下のファイル名を全て取得
' * 引数　：directoryPath パス
' * 戻り値：実行結果（カレントディレクトリ直下のディレクトリ名の配列。）
' *********************************************************************************************************************
'
Function getFileNames(directoryPath As String, fileExtensions As Variant, ByRef fileNames() As String)

    Dim fileName As String, msg As String
    
    Dim fileNameSize As Integer
    
    Dim fileExtension As Variant
    
    ' ディレクトリ移動
    ChDir directoryPath
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' 直下のファイル名を全て取得
    ' -----------------------------------------------------------------------------------------------------------------
    fileName = Dir(directoryPath & "\" & "*.*")
    
    Do While fileName <> ""
    
        ' ファイル名取得
        For Each fileExtension In fileExtensions
            If InStr(1, UCase(fileName), UCase(fileExtension)) > 0 Then
            
                ' フルパス＆ファイル名を追加格納。
                Call 一次元配列に値を追加(fileNames, directoryPath & "\" & fileName)
                Exit For
        
            End If
        Next
    
        fileName = Dir()
    Loop

End Function

' *********************************************************************************************************************
' * 機能　：パス直下のディレクトリ名を全て取得
' * 引数　：directoryPath パス
' * 戻り値：実行結果（カレントディレクトリ直下のディレクトリ名の配列。）
' *********************************************************************************************************************
'
Function getDirNames(directoryPath As String) As String()

    Call subFSO初期化

    Dim buf As String
    Dim dirNames() As String
    
    ' ディレクトリ移動
    ChDir directoryPath
    
    buf = Dir(directoryPath & "\" & "*.*", vbDirectory)
    Do While buf <> ""
        ' ディレクトリ名取得
        If objFSO.FolderExists(directoryPath & "\" & buf) And vbDirectory Then
            If buf <> "." And buf <> ".." Then
            
                ' ディレクトリ名を追加格納。
                Call 一次元配列に値を追加(dirNames, directoryPath & "\" & buf)
                
            End If
        End If
        buf = Dir()
        
    Loop
    
    getDirNames = dirNames

End Function
