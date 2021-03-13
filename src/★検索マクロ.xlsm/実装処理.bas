Attribute VB_Name = "実装処理"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数（共通）
' ---------------------------------------------------------------------------------------------------------------------

' 雛形シートコピー用（共通）
Public Const RESULT_SHEET_NAME = "処理結果"

' 処理結果シートデータ貼付け部の列数
Private Const RESULT_COL_LENGTH = 6

' ---------------------------------------------------------------------------------------------------------------------
' 定数（個別）
' ---------------------------------------------------------------------------------------------------------------------

Private Const KEY_検索ワード = "検索ワード"

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
' #                            （シート毎の処理呼び出しが必要かの判定値(boolean)を返却する）
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
Function 全体前処理(ws対象シート As Worksheet)

    ' -----------------------------------------------------------------------------------------------------------------
    ' 初期化処理
    ' -----------------------------------------------------------------------------------------------------------------

    ' -----------------------------------------------------------------------------------------------------------------
    ' 前処理
    ' -----------------------------------------------------------------------------------------------------------------

End Function

' *********************************************************************************************************************
' 機能　：検出されたファイルのブックごとに行いたい処理を実装する（シート毎の処理呼び出しが必要かの判定値(boolean)を返却する）
' *********************************************************************************************************************
'
Function ブックOPEN後処理(txtファイルパス As Variant, wb対象ブック As Workbook, ByRef results() As Variant) As Boolean

    ブックOPEN後処理 = True

End Function

' *********************************************************************************************************************
' 機能　：検出されたファイルの1シートごとに行いたい処理を実装する
' *********************************************************************************************************************
'
Function シート毎処理(txtファイルパス As Variant, ws対象シート As Worksheet, ByRef results() As Variant)

    ' -----------------------------------------------------------------------------------------------------------------
    ' 処理
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' 指定された検索文言リストの文字列の検索結果を収集する。
    Dim txt検索ワード As Variant
    
    ' 検索文言リスト分ループ
    For Each txt検索ワード In obj設定値シート.設定値リスト.Item(KEY_検索ワード)
    
        ' 検索文言がない場合、次の検索文言を処理する
        If txt検索ワード = "" Then
        
            GoTo ContinueBySearchArg
        End If
        
        ' ＜セルの検索＞
        Dim rng検索結果 As Range
        Set rng検索結果 = ws対象シート.UsedRange.Find( _
            what:=txt検索ワード, LookIn:=xlValues, _
            LookAt:=xlPart, MatchCase:=False, MatchByte:=False)
            
        ' セルへの検索結果がない場合
        If rng検索結果 Is Nothing Then
            ' 検索結果がなかった場合次の検索文言を処理する
            GoTo GotoCellSearchEnd
        End If

        Dim txt検索結果アドレス As String
        txt検索結果アドレス = rng検索結果.Address ' 検索結果のアドレスを配列に格納

        Do
            ' 結果を格納する
            Call 結果記録(results, rng検索結果, Array(txt検索ワード, "セル", rng検索結果.Value))
            
            Set rng検索結果 = ws対象シート.UsedRange.FindNext(After:=rng検索結果)
            
        Loop Until rng検索結果.Address = txt検索結果アドレス
        
GotoCellSearchEnd:

        ' ＜オートシェイプの検索＞
        Dim varオートシェイプリスト As Variant
        varオートシェイプリスト = getShapesProperty(ws対象シート)
        
        Dim i As Long
        Dim varオートシェイプの文字列 As Variant
        i = 0
        
        ' 検索文言リスト分ループ
        If Not IsEmpty(varオートシェイプリスト) Then
            For i = LBound(varオートシェイプリスト) To UBound(varオートシェイプリスト)
                varオートシェイプの文字列 = varオートシェイプリスト(i, 2)
                If Not IsEmpty(varオートシェイプの文字列) And InStr(varオートシェイプの文字列, txt検索ワード) Then
                
                    ' 結果を格納する
                    Call 結果記録(results, ws対象シート.Range(varオートシェイプリスト(i, 7)), _
                        Array(txt検索ワード, "オートシェイプ", varオートシェイプの文字列))
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
Function ブックCLOSE前処理(txtファイルパス As Variant, wb対象ブック As Workbook, ByRef results() As Variant) As Long


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
Sub 実行結果書式編集処理(ByRef ws対象シート As Worksheet)

    Dim i, lng最終行, lng最終列 As Long
    
    With ws対象シート
        lng最終行 = .UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        lng最終列 = .UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
        
        ' 書式コピー
        .Range(Cells(2, 1), Cells(2, lng最終列)).Copy
        .Range(Cells(2 + 1, 1), Cells(lng最終行, lng最終列)).PasteSpecial (xlPasteFormats)
        
        For i = 2 To lng最終行
            ' ハイパーリンク設定
            Dim strHyperLink As String
            strHyperLink = "=HYPERLINK(""[""&A" & i & "&""\""&B" & i & "&""]""&C" & i & "&""!" & .Cells(i, 4) & """,""" & .Cells(i, 4) & """)"
            
            .Range(.Cells(i, 4), .Cells(i, 4)).Value = strHyperLink
            
            ' 赤文字
            Call 検索該当文字の赤太文字化(.Range(Cells(i, 7), Cells(i, 7)), Cells(i, 5))
            
        Next
    End With
          
End Sub

' *********************************************************************************************************************
' 機能　：処理実行後に1度だけ実行したい処理を実装する
' *********************************************************************************************************************
'
Function 全体後処理(ws対象シート As Worksheet)


End Function

' #####################################################################################################################
' #
' # テンプレートメソッド以外のメソッド
' #
' #####################################################################################################################
'

Private Function 結果記録(ByRef results() As Variant, rng対象セル As Range, var出力内容 As Variant) As Variant

    Call reDimResult(RESULT_COL_LENGTH, results)

    Dim lng列 As Long: lng列 = UBound(results, 2)

    ' フォルダ名
    results(0, lng列) = rng対象セル.Parent.Parent.Path
    ' ファイル名
    results(1, lng列) = rng対象セル.Parent.Parent.Name
    ' シート名
    results(2, lng列) = rng対象セル.Parent.Name
    ' セル座標
    results(3, lng列) = rng対象セル.Address(False, False)
    
    Dim i As Long
    
    For i = LBound(var出力内容) To UBound(var出力内容)
    
        results(4 + i, lng列) = var出力内容(i)
    
    Next i
    
End Function
