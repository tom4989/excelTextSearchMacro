Attribute VB_Name = "テンプレート処理"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

Private Const KEY_パス = "パス"

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------

' 設定値リスト
Public obj設定値シート As cls設定値シート

' *********************************************************************************************************************
' * 機能　：マクロ呼び出し時（シートからの指定用）
' *********************************************************************************************************************

Sub マクロ開始()

    Call init開始時刻
    
    Dim wsMainSheet As Worksheet
    Dim fileCheck As Long
    
    Set obj設定値シート = New cls設定値シート
    obj設定値シート.ロード (ActiveSheet.Name)
    
    ' タイトル名に対するリストの情報（Range情報）
    ' Dim currentDirPathRangeList As Range, currentDirPathRange As Range
    ' Dim subDirCheckBoxRangeList As Range, subDirCheckBoxRange As Range
    
    ' 処理対象のファイル名一覧（フルパス＆ファイル名）
    Dim fileNames() As String
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' 初期化処理
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' 処理対象の拡張子を設定する。
    Dim fileExtention As Variant
    fileExtention = Split(FILE_EXTENSION, ",")
    
    ' 固有処理（マクロ呼び出し元）側のシート情報を取得する。
    ' Set wsMainSheet = MainSheet
    Set wsMainSheet = ActiveSheet
    
    ' 固有処理（マクロ呼び出し元）側のパス情報を取得する。
    ' Set currentDirPathRangeList = タイトル名指定でリスト値のRange情報を取得(TITLE_NAME_BY_TARGET_DIR, wsMainSheet)
    ' Set subDirCheckBoxRangeList = タイトル名指定でリスト値のRange情報を取得(TITLE_NAME_BY_DO_SUB_DIR, wsMainSheet)
    
    ' ★ConcreateProcess側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
    Call 全体前処理(wsMainSheet)


    ' -----------------------------------------------------------------------------------------------------------------
    ' パスの存在チェック
    ' -----------------------------------------------------------------------------------------------------------------

    Dim txtパス As Variant

    With wsMainSheet

        Dim i As Long
        i = 0
        ' 対象ディレクトリ分ループ
        ' If Not (obj設定値シート.設定値リスト.Item(KEY_パス) Is Nothing) Then
            For Each txtパス In obj設定値シート.設定値リスト.Item(KEY_パス)
            
                ' ディレクトリまたは、ファイルの存在チェック
                fileCheck = isDirectoryExist(CStr(txtパス))
                
                If 0 > fileCheck Then
                    MsgBox "以下のパスは存在しません。" + Chr(10) + "「" + txtパス + "」"
                    End
                End If
                i = i + 1
            Next
        ' End If
    End With

    ' -----------------------------------------------------------------------------------------------------------------
    ' ファイル名の収集
    ' -----------------------------------------------------------------------------------------------------------------

    Call setステータスバー("対象ファイル集計中...")
    
    With ActiveSheet
    
        i = 1
        '対象ディレクトリ分ループ
        ' If Not (obj設定値シート.設定値リスト.Item(KEY_パス) Is Nothing) Then
            For Each txtパス In obj設定値シート.設定値リスト.Item(KEY_パス)
            
                '指定の値がファイルの場合、その値をリストに追加し、ディレクトリの場合は、ファイル名の一覧を動的に取得して追加する。
                fileCheck = isDirectoryExist(CStr(txtパス))
                If 2 = fileCheck Then
                    ' 指定の値がファイルだった場合、その値をリストに追加
                    ' フルパス＆ファイル名を追加格納。
                    Call 一次配列に値を追加(fileNames, CStr(txtパス))
                Else
                    
                    ' ＜オートシェイプ情報の取得＞
                    Dim shapesCount As Long
                    Dim checkBoxChecked As Variant
                    Dim topLeftCellRow As Variant, topLeftCellColumn As Variant
            
                    ' オートシェイプ（チェックボックス）情報を取得。
                    Dim ShapesInfoList As Variant
                    ShapesInfoList = getShapesProperty(wsMainSheet, msoFormControl, xlCheckBox)
                    
                    ' 対象セル行のチェックボックスのチェック状態を取得（boolean形式で）
                    checkBoxChecked = True
                    
                    ' 現在のディレクトリ配下のファイル名を取得
                    Call doRepeat(txtパス, fileExtention, fileNames, checkBoxChecked)
                
                End If
                
                i = i + 1
            Next
            
        ' End If
        
    End With
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' ファイル処理メソッドの呼び出し
    ' -----------------------------------------------------------------------------------------------------------------
    
    Call ファイル処理(fileNames)
    
    ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
    Call 全体後処理(wsMainSheet)
    
    MsgBox "処理が終了しました。（処理時間：" & get処理時刻() & ")"

End Sub

' *********************************************************************************************************************
' * 機能　：対象ファイルの処理を行う。
' * 引数　：varArray 配列
' * 戻り値：判定結果（1:配列/0:空の配列/-1:配列ではない）
' *********************************************************************************************************************
'
Function ファイル処理(fileNames() As String)

    ' ファイル名の一覧が空だった場合、当Functionを中断する。
    If 1 > IsArrayEx(fileNames) Then
        MsgBox "処理対象ファイルが存在しません。"
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
    
    ' シート毎の処理呼び出し不要フラグ
    Dim unDealTargetSheetFlag As Boolean
    
    ' 処理結果保持用
    Dim results() As Variant
    
    index = 1
    total = UBound(fileNames) + 1
    
    Application.DisplayAlerts = False ' ファイルを開く際の警告を無効
    Application.ScreenUpdating = False ' 画面表示更新を無効
    
    For Each fileName In fileNames
    
        ' -------------------------------------------------------------------------------------------------------------
        ' 対象ブックを開いて、全シート分の処理を行う。
        ' -------------------------------------------------------------------------------------------------------------

        Call setステータスバー("(" & index & "/" & total & ")" & FSO.GetFileName(fileName))
        index = index + 1
        
        Set targetWB = Workbooks.Open(fileName, UpdateLinks:=0, IgnoreReadOnlyRecommended:=False)
        
        ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
        unDealTargetSheetFlag = ブックOPEN後処理(fileName, targetWB, results)
        
        If False = unDealTargetSheetFlag Then
            Dim i As Integer
            For i = 1 To targetWB.Worksheets.Count ' シートの数分ループする
            
                Set targetSheet = targetWB.Worksheets(i)
                
                ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
                Call シート毎処理(fileName, targetSheet, results)
                
            Next i
            
        End If
        
        Dim ファイルCLOSE方法区分値 As Long
        
        ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
        ファイルCLOSE方法区分値 = ブックCLOSE前処理(fileName, targetWB, results)
        
        If ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存しないで閉じる Then
            targetWB.Close
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存して閉じる Then
            targetWB.Save
            targetWB.Close
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存しないで閉じない Then
            
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存して閉じない Then
            targetWB.Save
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.処理中断 Then
            End
        End If
    Next
        
    ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
    ' 実行結果の編集（結果のマージ、並び替え、フィルタリング当）
    Call 実行結果内容編集処理(results)
    
    If Not Not results Then
    
        If UBound(results, 2) <> 0 Then
        
            ' ファイルの保存形式をexcel2007形式（.xlsx)に変更
            Application.defaultSaveFormat = xlOpenXMLWorkbook
            
            Set targetWB = Workbooks.Add
            
            ' 当ブックにシート「雛形」が用意されている場合、指定ブックの先頭にコピーした後、
            ' シート名を「処理結果」に変更する。（ない場合は新規作成ブックのsheet1を利用）
            Call 雛形シートコピー(targetWB)
            
            ' 結果貼り付け行の取得。
            ' A列に値が設定されている行を、表題欄としてその行数を取得する
            Dim MaxRow As Integer
            With targetWB.ActiveSheet.UsedRange
                MaxRow = .Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
            End With
            ' 結果貼り付け行の設定。
            MaxRow = MaxRow + 1
            
            ' 結果貼り付け
            targetWB.ActiveSheet.Range(Cells(MaxRow, 1), Cells(UBound(results, 2) + 2, UBound(results) + 1)) = 二次元配列行列逆転(results)
            
            Dim MaxCol As Integer
            ' 書式コピー
            With targetWB.ActiveSheet
                MaxRow = .UsedRange.Find("*", , xlFormulas, xlByRows, xlPrevious).Row
                MaxCol = .UsedRange.Find("*", , xlFormulas, xlByColumns, xlPrevious).Column
                
                .Range(.Cells(2, 1), .Cells(2, MaxCol)).Copy
                .Range(.Cells(2 + 1, 1), .Cells(MaxRow, MaxCol)).PasteSpecial (xlPasteFormats)
            End With
            
            ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
            Call 実行結果書式編集処理(targetWB.ActiveSheet)
            
            ' "A1"を選択状態にする
            targetWB.ActiveSheet.Cells(1, 1).Select
            
            ' シート名「処理結果」以外のシートを削除する
            Call 不要シート削除(targetWB, RESULT_SHEET_NAME)
            
        Else
            
            MsgBox "処理結果は0件です。"
        End If
        
    Else
    
        MsgBox "処理結果は0件です。"
        
    End If
            
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Application.StatusBar = False
    
    ' ファイルの保存形式を元の状態に戻す
    Application.defaultSaveFormat = defaultSaveFormat
    
    If Not Not results Then
        If UBound(results, 2) <> 0 Then
            targetWB.Activate
        End If
    End If

End Function


' *********************************************************************************************************************
' * 機能　：当ブックのシート「雛形」を指定ブックの先頭にコピーした後、
' * 　　　　シート名を「処理結果」に変更する
' *********************************************************************************************************************
'
Sub 雛形シートコピー(targetWB As Workbook)

    Dim myWorkBook  As String
    Dim newWorkBook As String
    Dim targetSheet As Worksheet
    Dim sheetName   As String
    
    ' マクロを実行中のブック名を取得
    myWorkBook = ThisWorkbook.Name
    
    ' 新規ブック名を取得
    newWorkBook = targetWB.Name
    
    ' マクロ実行時のブックをアクティブにする
    Workbooks(myWorkBook).Activate
    
    ' シート「雛形」があった場合、指定ブックにコピー（一番前に挿入）する
    Dim i As Integer
    For i = 1 To Workbooks(myWorkBook).Worksheets.Count ' シートの数分ループする
    
        Set targetSheet = Workbooks(myWorkBook).Worksheets(i)
        
        If TEMPLATE_SHEET_NAME = targetSheet.Name Then
            Workbooks(myWorkBook).Sheets(TEMPLATE_SHEET_NAME).Copy _
            Before:=Workbooks(newWorkBook).Sheets(1)
        End If
        
    Next i
    
    ' マクロを実行中のブックをアクティブにする
    Workbooks(targetWB.Name).Sheets(TEMPLATE_SHEET_NAME).Activate
    ' シート名を「処理結果」に変更する
    Workbooks(targetWB.Name).Sheets(TEMPLATE_SHEET_NAME).Name = RESULT_SHEET_NAME
    
End Sub
