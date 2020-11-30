Attribute VB_Name = "実装処理"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数（共通）
' ---------------------------------------------------------------------------------------------------------------------

' 雛形シートコピー用（共通）
Public Const TEMPLATE_SHEET_NAME = "雛形"
Public Const RESULT_SHEET_NAME = "処理結果"

' 対象の拡張子
Public Const FILE_EXTENSION = "xls,xlsx,xlsm"

' 処理結果シートデータ貼付け部の列数
Private Const RESULT_COL_LENGTH = 6

' ---------------------------------------------------------------------------------------------------------------------
' 定数（個別）
' ---------------------------------------------------------------------------------------------------------------------

Private Const KEY_パス = "パス"
Private Const KEY_検索ワード = "検索ワード"
Private Const KEY_対象ブックシート = "対象ブックシート"
Private Const KEY_対象外ブックシート = "対象外ブックシート"

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------

' 処理した件数
Dim lngResultCount As Long

' #####################################################################################################################
' #
' # テンプレートメソッド(テンプレート処理から呼び出されるメソッド）
' #
' # 1. 全体前処理()            処理実行前に1度だけ実行したい処理を実装する
' # 2. ブックOPEN後処理()      検出されたファイルのブックごとに行いたい処理を実装する
' #                            （シート毎の処理呼び出しが不要かの判定値(boolean)を返却する）
' # 3. シート毎処理()          検出されたファイルの1シートごとに行いたい処理を実装する
' # 4. ブックCLOSE前処理()     検出されたファイルのブックごとに行いたい後処理を実装する
' # 5. 実行結果内容編集処理()  実行結果について、ファイルに出力する前に編集したい場合に実装する（重複の削除、ソート等）
' # 6. 実行結果書式編集処理()  ファイルに出力した後の実行結果を編集したい場合に実装する（ハイパーリンクの設定等）
' # 7. 全体後処理()            処理実行後に1度だけ実行したい処理を実装する
' #
' #####################################################################################################################
'

' *********************************************************************************************************************
' 機能　：固有処理側の前処理
' *********************************************************************************************************************
'
Function 全体前処理(targetSheet As Worksheet)

    ' -----------------------------------------------------------------------------------------------------------------
    ' 初期化処理
    ' -----------------------------------------------------------------------------------------------------------------

    ' 処理した件数の初期化
    ' resultCount = 0

    ' -----------------------------------------------------------------------------------------------------------------
    ' 前処理
    ' -----------------------------------------------------------------------------------------------------------------

End Function

' *********************************************************************************************************************
' 機能　：検出されたファイルのブックごとに行いたい処理を実装する（シート毎の処理呼び出しが不要かの判定値(boolean)を返却する）
' *********************************************************************************************************************
'
Function ブックOPEN後処理(fileName As Variant, targetWB As Workbook, ByRef results() As Variant) As Boolean


End Function

' *********************************************************************************************************************
' 機能　：検出されたファイルの1シートごとに行いたい処理を実装する
' *********************************************************************************************************************
'
Function シート毎処理(fileName As Variant, targetSheet As Worksheet, ByRef results() As Variant)

    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim targetWB As Workbook
    
    Dim ShapesInfoList As Variant
    Dim ShapesInf As Variant
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' 処理
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' 指定された検索文言リストの文字列の検索結果を収集する。
    
    ' 対象シートの検索結果を「FoundAddr」に格納する。
    Dim firstAddress As String
    Dim FoundCell As Range
    
    Dim lngResultCount As Long ' 結果件数
    
    Dim txt検索ワード As Variant
    
    ' 検索文言リスト分ループ
    For Each txt検索ワード In obj設定値シート.設定値リスト.Item(KEY_検索ワード)
    
        ' 検索文言がない場合、次の検索文言を処理する
        If "" = txt検索ワード Then
        
            GoTo ContinueBySearchArg
        End If
        
        ' ＜セルの検索＞
        Set FoundCell = targetSheet.UsedRange.Find(what:=txt検索ワード, LookIn:=xlValues, _
            LookAt:=xlPart, MatchCase:=False, MatchByte:=False)
            
        ' セルへの検索結果がない場合
        If FoundCell Is Nothing Then
            ' 検索結果がなかった場合次の検索文言を処理する
            GoTo GotoCellSearchEnd
        End If
        
        firstAddress = FoundCell.Address ' 検索結果のアドレスを配列に格納
        
        Do
            ' 結果を格納する
            Call reDimResult(RESULT_COL_LENGTH, results)                   ' 結果保持の配列作成
            lngResultCount = UBound(results, 2)
            
            results(0, lngResultCount) = txt検索ワード                     ' 検索文言
            results(1, lngResultCount) = FSO.GetParentFolderName(fileName) ' フォルダ名
            results(2, lngResultCount) = FSO.GetFileName(fileName)         ' ファイル名
            results(3, lngResultCount) = targetSheet.Name                  ' シート名
            results(4, lngResultCount) = FoundCell.Address(False, False)   ' 座標
            results(5, lngResultCount) = "セル"                            ' セル／オートシェイプ
            results(6, lngResultCount) = FoundCell.Value                   ' 文字列
            
            Set FoundCell = targetSheet.UsedRange.FindNext(After:=FoundCell)
            
        Loop Until FoundCell.Address = firstAddress
        
GotoCellSearchEnd:

        ' ＜オートシェイプの検索＞
        ShapesInfoList = getShapesProperty(targetSheet)
        Dim i As Integer
        Dim textValue As Variant
        i = 0
        
        ' 検索文言リスト分ループ
        If Not IsEmpty(ShapesInfoList) Then
            For i = LBound(ShapesInfoList) To UBound(ShapesInfoList)
                textValue = ShapesInfoList(i, 2)
                If Not IsEmpty(textValue) And InStr(textValue, txt検索ワード) Then
                
                    ' 結果を格納する
                    Call reDimResult(RESULT_COL_LENGTH, results)                   ' 結果保持の配列作成
                    lngResultCount = UBound(results, 2)
                    
                    results(0, lngResultCount) = txt検索ワード                     ' 検索文言
                    results(1, lngResultCount) = FSO.GetParentFolderName(fileName) ' フォルダ名
                    results(2, lngResultCount) = FSO.GetFileName(fileName)         ' ファイル名
                    results(3, lngResultCount) = targetSheet.Name                  ' シート名
                    results(4, lngResultCount) = ShapesInfoList(i, 7)              ' 座標
                    results(5, lngResultCount) = "オートシェイプ"                  ' セル／オートシェイプ
                    results(6, lngResultCount) = textValue                         ' 文字列
                End If
            Next i
        End If
        
ContinueBySearchArg:

    Next
    
End Function

' *********************************************************************************************************************
' 機能　：検出されたファイルのブックごとに行いたい後処理を実装する
' *********************************************************************************************************************
'
Function ブックCLOSE前処理(fileName As Variant, targetWB As Workbook, ByRef results() As Variant) As Long


End Function

' *********************************************************************************************************************
' 機能　：実行結果について、ファイルに出力する前に編集したい場合に実装する（重複の削除、ソート等）
' *********************************************************************************************************************
'
Function 実行結果内容編集処理(ByRef var変換元() As Variant) As Variant

End Function

' *********************************************************************************************************************
' 機能　：ファイルに出力した後の実行結果を編集したい場合に実装する（ハイパーリンクの設定等）
' *********************************************************************************************************************
'
Sub 実行結果書式編集処理(ByRef targetSheet As Worksheet)

    Dim i, MaxRow, MaxCol As Long
    
    With targetSheet
        MaxRow = .UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        MaxCol = .UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
        
        ' 書式コピー
        .Range(Cells(2, 1), Cells(2, MaxCol)).Copy
        .Range(Cells(2 + 1, 1), Cells(MaxRow, MaxCol)).PasteSpecial (xlPasteFormats)
        
        For i = 2 To MaxRow
            ' ハイパーリンク設定
            Dim strHyperLink As String
            strHyperLink = editHYPERLINK数式(.Cells(i, 2), .Cells(i, 3), .Cells(i, 4), .Cells(i, 5))
            
            .Range(.Cells(i, 5), .Cells(i, 5)).Value = strHyperLink
            
            ' 赤文字
            Call 検索該当文字の赤太文字化(.Range(Cells(i, 7), Cells(i, 7)), Cells(i, 1))
            
        Next
    End With
          
End Sub

' *********************************************************************************************************************
' 機能　：処理実行後に1度だけ実行したい処理を実装する
' *********************************************************************************************************************
'
Function 全体後処理(targetSheet As Worksheet)

End Function

' #####################################################################################################################
' #
' # テンプレートメソッド以外のメソッド
' #
' #####################################################################################################################
'

' なし
