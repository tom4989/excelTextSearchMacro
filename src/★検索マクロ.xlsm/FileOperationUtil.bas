Attribute VB_Name = "FileOperationUtil"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

'パスのデリミタ
Public Const PATH_DELIMITER = "\"

' ファイル名、拡張子のデリミタ
Public Const FILE_DELIMITER = "."

' ファイル操作情報
Public Type ファイル操作情報
    フルパス_ファイル名 As String
    フルパス As String
    親ディレクトリまでのフルパス As String
    対象ディレクトリ名 As String
    対象ファイル名 As String
    対象ファイル情報() As String
End Type

' 対象の拡張子（モジュール）
Public Const FILE_EXTENSION_OF_MODULE = "bas,cls,frm"

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

' ルートパス作成済フラグ
Private rootPathMaked As Boolean

' #####################################################################################################################
' #
' # アクセサ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * 機能　：ルートパス作成済みフラグを設定する
' *********************************************************************************************************************
'
Public Function getRootPathMaked() As Boolean
    getRootPathMaked = rootPathMaked
End Function

' *********************************************************************************************************************
' * 機能　：ルートパス作成済みフラグを取得する
' *********************************************************************************************************************
'
Public Function setRootPathMaked(isMaked As Boolean)
    rootPathMaked = isMaked
End Function

' #####################################################################################################################
' #
' # ファイル操作ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * 機能　：ファイル名取得
' *********************************************************************************************************************
'
Function ファイル名取得(txtパス As String) As String

    If objFSO Is Nothing Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    End If

    objFSO.GetFileName (txtパス)

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
' * 機能　：フォルダが存在しなかったら作成する
' *********************************************************************************************************************
'
Function mkdirIFNotExist(txtフォルダ名 As String)

    If Dir(txtフォルダ名, vbDirectory) = "" Then
        mkdir txtフォルダ名
    End If

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
    Dim buf As String
    
    Dim dirNames() As String
    
    ' ディレクトリ移動
    ChDir directoryPath
    
    buf = Dir(directoryPath & "\" & "*.*", vbDirectory)
    Do While buf <> ""
        ' ディレクトリ名取得
        If GetAttr(directoryPath & "\" & buf) And vbDirectory Then
            If buf <> "." And buf <> ".." Then
            
                ' ディレクトリ名を追加格納。
                Call 一次元配列に値を追加(dirNames, directoryPath & "\" & buf)
                
            End If
        End If
        buf = Dir()
        
    Loop
    
    getDirNames = dirNames

End Function


' *********************************************************************************************************************
' * 機能　：対象ディレクトリを作成する（対象パスが未存在、作成ディレクトリ名が存在した場合は処理中断）
' *********************************************************************************************************************
'
Function ディレクトリ作成(ByVal ルートパス As String, ByVal 処理日時 As String, ByVal 相対パス As String)

    Dim dirCheck As Long
    
    ' ルートパスの存在チェック
    dirCheck = isDirectoryExist(CStr(ルートパス & 処理日時))
    ' 対象パスが未設定の場合（ルートパス作成時）
    If "" = 相対パス Then
        ' 処理日時が設定済の場合、ルートパスが作成済であれば処理中断とする
        If "" <> 処理日時 Then
            If 0 < dirCheck And False = getRootPathMaked() Then
                MsgBox "以下のディレクトリは既に存在するため処理を中断します。" + Chr(10) + "「" + ルートパス & 処理日時 + "」"
                End
            End If
        End If
    End If

    ' ディレクトリ作成
    dirCheck = isDirectoryExist(CStr(ルートパス & 処理日時 & PATH_DELIMITER & 相対パス))
    If 0 > dirCheck Then
        mkdir ルートパス & 処理日時 & PATH_DELIMITER & 相対パス
        Call setRootPathMaked(True)
    End If
    
End Function








