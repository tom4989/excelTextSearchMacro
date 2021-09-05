Attribute VB_Name = "Util"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------

Dim var開始時刻 As Variant

' #####################################################################################################################
' #
' # ログ系ユーティリティ
' #
' #####################################################################################################################

Sub log(ByVal strメッセージ As String)

    Debug.Print Format(Now(), "HH:mm:ss ") & strメッセージ

End Sub

Function getTimestamp()

    getTimestamp = Format(Now(), "yyyymmdd_HHnnss")

End Function

' #####################################################################################################################
' #
' # メッセージ系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * 機能　：開始メッセージを取得する
' *********************************************************************************************************************
'
Function get開始メッセージ(ByVal txt処理名 As String) As String

    Call init開始時刻

    Dim txtメッセージ As String
    txtメッセージ = Format(Now(), "HH:mm:ss ") & txt処理名 & "処理を開始します。"

    Debug.Print txtメッセージ
    get開始メッセージ = txtメッセージ

End Function

' *********************************************************************************************************************
' * 機能　：終了メッセージを取得する
' *********************************************************************************************************************
'
Function get終了メッセージ(ByVal txt処理名 As String) As String

    Dim txtメッセージ As String
    txtメッセージ = Format(Now(), "HH:mm:ss ") & txt処理名 & "処理が終了しました。（処理時間：" & get処理時刻 & "）"

    Debug.Print txtメッセージ
    get終了メッセージ = txtメッセージ

End Function

' *********************************************************************************************************************
' * 機能　：エラー時のメッセージを取得する
' *********************************************************************************************************************
'
Function get異常時メッセージ(ByVal txt処理名 As String) As String

    Dim txtメッセージ As String
    txtメッセージ = Format(Now(), "HH:mm:ss ") & txt処理名 & "処理が終了しました。（処理時間：" & get処理時刻 & "）"

    Debug.Print txtメッセージ
    get異常時メッセージ = txtメッセージ
End Function

' *********************************************************************************************************************
' * 機能　：エラーオブジェクトの内容をメッセージダイアログ、ログに出力する。
' *********************************************************************************************************************
'
Sub subエラー表示(Optional argサイレントモード As Boolean = False)

    Dim txtエラー内容 As clsStringBuilder
    Set txtエラー内容 = New clsStringBuilder
    
    txtエラー内容.append ("Description: ")
    txtエラー内容.appendLine (err.Description)
    
    txtエラー内容.append ("HelpContext: ")
    txtエラー内容.appendLine (err.HelpContext)
    
    txtエラー内容.append ("HelpFile: ")
    txtエラー内容.appendLine (err.HelpFile)
    
    txtエラー内容.append ("LastDllError: ")
    txtエラー内容.appendLine (err.LastDllError)
    
    txtエラー内容.append ("Number: ")
    txtエラー内容.appendLine (err.Number)

    If Not argサイレントモード Then
        MsgBox txtエラー内容.toString, vbCritical
    End If
    
    subエラーログファイル出力 (txtエラー内容.toString)
        
End Sub


' #####################################################################################################################
' #
' # ステータスバー操作系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * 機能　：ステータスバーに表示する処理時間を初期化する
' *********************************************************************************************************************
'
Sub init開始時刻()

    var開始時刻 = Now()
    
End Sub

' *********************************************************************************************************************
' * 機能　：処理時間の開始時刻を取得する
' *********************************************************************************************************************
'
Function get開始時刻()

    get開始時刻 = var開始時刻

End Function

' *********************************************************************************************************************
' * 機能　：処理時間を HH:mm:ss 形式で取得する
' *********************************************************************************************************************
'
Function get処理時刻()

    get処理時刻 = Format(Now() - var開始時刻, "HH:mm:ss")
    
End Function

' *********************************************************************************************************************
' * 機能　：ステータスバーに経過時間付でメッセージを表示する
' *********************************************************************************************************************
'
Sub setステータスバー(ByVal strメッセージ As String)

    If IsEmpty(var開始時刻) Then
        
        var開始時刻 = Now()
        
    End If
    
    Application.StatusBar = get処理時刻() & " " & strメッセージ

End Sub

' *********************************************************************************************************************
' * 機能　：サイレントモードか否かでメッセージの出し分けを行う
' * 　サイレントモードの場合：
' * 　　ステータスバーに表示
' * 　サイレントモードでない場合：
' * 　　ダイアログで表示
' *********************************************************************************************************************
'
Sub s_メッセージ通知(txtメッセージ As String, flgサイレントモード As Boolean)

    If flgサイレントモード Then
        
        setステータスバー (txtメッセージ)
        
    Else
    
        MsgBox txtメッセージ

    End If

End Sub


' #####################################################################################################################
' #
' # ブック、シート操作系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * 機能　：引数で渡されたシート名以外のシートを削除する
' *********************************************************************************************************************
'
Function 不要シート削除(対象ブック情報 As Workbook, ByVal 残すシート名 As String)

    Dim 前状態 As Boolean
    前状態 = Application.DisplayAlerts
    
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    
    For Each ws In 対象ブック情報.Worksheets
    
        If ws.Name <> 残すシート名 Then
            Worksheets(ws.Name).Delete
        End If
        
    Next ws
    
    Application.DisplayAlerts = 前状態
        
End Function

' *********************************************************************************************************************
' * 機能　：引数で渡されたシートの最終行を取得する
' *********************************************************************************************************************
'
Function 最終行取得(ws対象シート As Worksheet, Optional useUsedRange As Boolean = True) As Long

    If useUsedRange Then

        With ws対象シート.UsedRange
            最終行取得 = .Rows(.Rows.Count).Row
        End With
        
    Else
        With ws対象シート
            最終行取得 = .Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        End With
        
    End If
        
End Function


' *********************************************************************************************************************
' * 機能　：引数で渡されたシートの最終列を取得する
' *********************************************************************************************************************
'
Function 最終列取得(ws対象シート As Worksheet, Optional useUsedRange As Boolean = True) As Long

    If useUsedRange Then

        With ws対象シート.UsedRange
            最終列取得 = .Columns(.Columns.Count).Column
        End With
        
    Else
        With ws対象シート
            最終列取得 = .Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
        End With
    
    End If
        
End Function

' *********************************************************************************************************************
' * 機能　：引数で渡されたシートの内容をVariant変数に変換して返す
' *********************************************************************************************************************
'
Function シート内容取得(wsワークシート As Worksheet) As Variant

    With wsワークシート

        シート内容取得 = .Range( _
            .Cells(1, 1), _
            .Cells(最終行取得(wsワークシート), 最終列取得(wsワークシート)))
    End With

End Function


' *********************************************************************************************************************
' * 機能　：シートのコピー
' *********************************************************************************************************************
'
Sub f_シートコピー(wb対象ブック As Workbook, txtコピー元シート名 As String, txtコピー先シート名 As String)

    If f_シート存在チェック(wb対象ブック, txtコピー先シート名) Then
    
        Dim flgDisplayAlerts As Boolean
        flgDisplayAlerts = Application.DisplayAlerts
    
        Application.DisplayAlerts = False
        wb対象ブック.Sheets(txtコピー先シート名).Delete
        Application.DisplayAlerts = flgDisplayAlerts

    End If

    wb対象ブック.Sheets(txtコピー元シート名).Copy Before:=wb対象ブック.Sheets(txtコピー元シート名)
    ActiveSheet.Name = txtコピー先シート名
    
End Sub

' *********************************************************************************************************************
' * 機能　：シート存在チェック
' *********************************************************************************************************************
'
Function f_シート存在チェック(wb対象ブック As Workbook, txtシート名 As String) As Boolean

    Dim wsワークシート As Worksheet
    
    For Each wsワークシート In wb対象ブック.Sheets
    
        If wsワークシート.Name = txtシート名 Then
        
            f_シート存在チェック = True
            Exit Function
            
        End If
    
    Next wsワークシート

    f_シート存在チェック = False

End Function


' #####################################################################################################################
' #
' # ダイアログ操作系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' * 機能　：処理続行 or 中止確認ダイアログを表示する
' *********************************************************************************************************************
'
Function 処理続行判断(message As String)

    Dim rc As VbMsgBoxResult
    rc = MsgBox(message + Chr(10) + "処理を続行しますか？", vbYesNo, vbQuestion)
    
    If rc = vbYes Then
        MsgBox "処理を続けます", vbInformation
    Else
        MsgBox "処理を中止しました。", vbCritical
        
        ' マクロの実行中止
        End
    End If

End Function


' #####################################################################################################################
' #
' # オートシェイプ操作系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' 機能名：対象シート上にあるオブジェクトのプロパティを取得する
' 戻り　：getShapesProperty as String(2, n)
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
    
    ' 配列の作成。
    i = 0
    For Each obj In targetSheet.Shapes
        ' FORMコントロールの場合
        If obj.Type = objType Then
            ' 渡されたフォームコントロールタイプが一致した場合、カウントアップ
            If obj.FormControlType = formCtlType Then
                i = i + 1
            End If
            
            ' 指定なし又は、それ以外のオートシェイプ
            ElseIf objType = -999 Or obj.Type = objType Then
                i = i + 1
            End If
    Next
        
    ' 対象のオートシェイプがみつかった場合のみ、そのオブジェクトの格納を行う。
    If 0 <> i Then
        ReDim ret(i - 1, 12)
        
        ' 配列の作成
        i = 0
        ' オブジェクト情報の設定
        For Each obj In targetSheet.Shapes
            
            ' formコントロールの場合
            If obj.Type = objType Then
                ' 渡されたフォームコrントロールタイプが一致した場合、値を取得する。
                If obj.FormControlType = formCtlType Then
                        
                    ret(i, 0) = obj.Type
                    ret(i, 1) = obj.AlternativeText
                        
                    ' TextFrameプロパティが使用できない（レイアウト枠がない）オブジェクトは除外
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
                    
            ' 指定なし又は、それ以外のオートシェイプなどの場合
            ElseIf objType = -999 Or obj.Type = objType Then
                
                ret(i, 0) = obj.Type
                ret(i, 1) = obj.AlternativeText
                        
                ' TextFrameプロパティが使用できない（レイアウト枠がない）オブジェクトは除外
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
' # 配列操作系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' 機能　：引数が配列か判定し、配列の場合は空かどうかも判定する
' 引数　：varArray 配列
' 戻り値：判定結果（1:配列/0:空の配列/-1：配列じゃない)
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
    If err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

' *********************************************************************************************************************
' 機能　：シート内容を格納した配列から座標(A1等)を指定して値を取得する
' *********************************************************************************************************************
'
Public Function f_セル座標の値取得(ByRef var配列 As Variant, txtセル座標 As String) As String

    Dim var座標 As Variant
    var座標 = CAlpNum2Num(txtセル座標)

    funcセル座標の値取得 = var配列(var座標(0), var座標(1))

End Function

' *********************************************************************************************************************
' 機能　：対象の配列に、指定された文字列が格納されているか判定する
' *********************************************************************************************************************
'
Public Function containArray(var配列 As Variant, txt対象文字列 As String) As Boolean

    Dim i As Long
    
    For i = LBound(var配列) To UBound(var配列)
    
        If var配列(i) = txt対象文字列 Then
        
            containArray = True
            Exit Function
        
        End If
    Next i
    
    containArray = False

End Function

' *********************************************************************************************************************
' 機能：対象の配列に、指定された文字列が格納されているか判定する
' *********************************************************************************************************************
'
Public Function f_配列含まれているかチェック( _
    var配列 As Variant, txt対象文字列 As String, _
    Optional flg文字列にワイルドカードあり As Boolean = False, _
    Optional flg配列にワイルドカードあり As Boolean = False) As Boolean

    Dim i As Long
    
    If IsArrayEx(var配列) <> 1 Then
    
        f_配列含まれているかチェック = False
        Exit Function
    End If
    
    For i = LBound(var配列) To UBound(var配列)
    
        If var配列(i) = txt対象文字列 Then
        
            f_配列含まれているかチェック = True
            Exit Function
        End If
        
        If flg文字列にワイルドカードあり Then
        
            If var配列(i) Like txt対象文字列 Then
            
                f_配列含まれているかチェック = True
                Exit Function
            End If
        End If
        
        If flg配列にワイルドカードあり Then
        
            If txt対象文字列 Like var配列(i) Then
            
                f_配列含まれているかチェック = True
                Exit Function
            End If
        End If
        
    Next i
    
    f_配列含まれているかチェック = False

End Function

' *********************************************************************************************************************
' 機能　：実行結果を保持する二次元配列変数を定義するFunction
' *********************************************************************************************************************
'
Function reDimResult(ByVal topLevelElementSize As Integer, ByRef results() As Variant)

    Select Case IsArrayEx(results)
        Case 1
            ' resultsが初期化済の場合
            ' 現在のレコード数 + 1行領域を確保
            ReDim Preserve results(topLevelElementSize, UBound(results, 2) + 1)
        Case 0
            ' resultsが1度も初期化されていない場合
            ' 1行領域を確保
            ReDim Preserve results(topLevelElementSize, 0)
    End Select
        
End Function

' *********************************************************************************************************************
' 機能　：一次元配列に新たな要素を追加する
' *********************************************************************************************************************
'
Function 一次元配列に値を追加(ByRef valueList As Variant, ByVal 追加設定値 As String)

    ' ファイル名を取得する
    Select Case IsArrayEx(valueList)
        Case 1
            ReDim Preserve valueList(UBound(valueList) + 1)
        Case 0
            ReDim Preserve valueList(0)
    End Select
    
    ' 追加したリストに、設定値を格納。
    valueList(UBound(valueList)) = 追加設定値
    
End Function

' *********************************************************************************************************************
' 機能　：二次元配列の行と列を入れ替える
' *********************************************************************************************************************
'
Function 二次元配列行列逆転(ByRef var二次元配列 As Variant)

    Dim var逆転後配列 As Variant
    
    ReDim var逆転後配列( _
        LBound(var二次元配列, 2) To UBound(var二次元配列, 2), _
        LBound(var二次元配列) To UBound(var二次元配列))
        
    Dim i, j As Long
    
    For i = LBound(var二次元配列) To UBound(var二次元配列, 2)
        
        For j = LBound(var二次元配列) To UBound(var二次元配列)
            
            var逆転後配列(i, j) = var二次元配列(j, i)
            
        Next
    Next
    
    二次元配列行列逆転 = var逆転後配列
        
    
End Function


' #####################################################################################################################
' #
' # 装飾系ユーティリティ
' #
' #####################################################################################################################
    
' *********************************************************************************************************************
' 機能　：対象セル範囲内で検索文字列に該当した文字列を赤太文字にする
' *********************************************************************************************************************
'
Function 検索該当文字の赤太文字化(prmRange As Range, prmTargetString As String)

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
' # シート情報取得系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' 機能　：タイトル名指定でリスト値を取得
'         ※リスト値がなかった場合、配列の要素数1（値は空）が返却されます。
' *********************************************************************************************************************
'
Function タイトル名指定でリスト値を取得(titleName As String, targetSheet As Worksheet) As Variant

    Dim targetRangeList As Range
    Dim targetVariantList As Variant
    
    Set targetRangeList = タイトル名指定でリスト値のRange情報を取得(titleName, targetSheet)
    ' 配列か判定
    If targetRangeList.Count = 1 Then
        targetVariantList = Array(targetRangeList.Item(1).Value)
    Else
        targetVariantList = targetRangeList.Value
    End If
    
    タイトル名指定でリスト値を取得 = targetVariantList
    
End Function

' *********************************************************************************************************************
' 機能　：引数で指定された行が選択状態であるか判定する
' *********************************************************************************************************************
'
Function is選択状態(ByVal lng対象行 As Long)

    Dim rng As Range
    
    For Each rng In Selection.Rows
    
        If rng.Row = lng対象行 Then
        
            is選択状態 = True
            Exit Function
            
        End If
        
    Next rng
        
    is選択状態 = False


End Function

' *********************************************************************************************************************
' 機能：列の値を数字に変換する
' *********************************************************************************************************************
'
Function CAlp2Num(txtAlphabet As String) As Long
  
    CAlp2Num = ActiveSheet.Range(txtAlphabet & "1").Column
    
End Function


' *********************************************************************************************************************
' 機能：列(A:B形式)の値を数字に変換する
' *********************************************************************************************************************
'
Function CAlpxAlp2Num(txtAlpxAlp As String) As Variant
  
    Dim var結果 As Variant

    var結果 = Split(txtAlpxAlp, ":")

    var結果(0) = CAlp2Num(CStr(var結果(0)))
    
    If UBound(var結果) >= 1 Then
        var結果(1) = CAlp2Num(CStr(var結果(1)))
    Else
        Call 一次元配列に値を追加(var結果, var結果(0))
    End If

    CAlpxAlp2Num = var結果
    
End Function

' *********************************************************************************************************************
' 機能：セル座標を数字に変換する
' *********************************************************************************************************************
'
Function CAlpNum2Num(txt座標 As String) As Variant

    Dim var結果 As Variant
    ReDim var結果(1)

    Dim objReg As Object, objMatch As Object
    
    Set objReg = CreateObject("VBScript.RegExp")
    objReg.Pattern = "^([A-Z]+)([0-9]+)$"

    Set objMatch = objReg.Execute(txt座標)

    Dim txt行 As String, txt列 As String
    
    var結果(0) = CLng(CAlp2Num(objMatch(0).SubMatches(0)))
    var結果(1) = CLng(objMatch(0).SubMatches(1))

    CAlpNum2Num = var結果

End Function

' #####################################################################################################################
' #
' # 文字列ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' 機能：引数で指定された文字列の共通部を返す
' *********************************************************************************************************************
'
Public Function f_共通部取得(txt文字列1 As String, txt文字列2 As String) As String

    Dim lng文字数 As Long
    
    If Len(txt文字列1) <= Len(txt文字列2) Then
    
        lng文字数 = Len(txt文字列1)
    Else
        lng文字数 = Len(txt文字列2)
    
    End If
    
    Dim i As Long
    
    For i = 1 To lng文字数
    
        If Left(txt文字列1, i) <> Left(txt文字列2, i) Then
        
            Exit For
            
        End If
        
    Next i
    
    If i > 1 Then
    
        f_共通部取得 = Left(txt文字列1, i - 1)
    
    Else
        f_共通部取得 = ""
    
    End If
    
End Function

' *********************************************************************************************************************
' 機能：指定文字のRTRIM
' *********************************************************************************************************************
'
Function f_RTRIM(txt対象文字列 As String, txt指定文字 As String) As String

    If txt対象文字列 <> "" Then

        If Right(txt対象文字列, 1) = txt指定文字 Then
    
            f_RTRIM = Left(txt対象文字列, Len(txt対象文字列) - 1)
            Exit Function
    
        End If
        
    End If
    
    f_RTRIM = txt対象文字列
    
End Function

' #####################################################################################################################
' #
' # Dictionary系ユーティリティ
' #
' #####################################################################################################################

' *********************************************************************************************************************
' 機能：Dictionary文字列の結合
' *********************************************************************************************************************
'
Function f_Dictonary結合(ByRef dic接続情報 As Object) As String

    Dim txt接続文字列
    
    Dim var設定値 As Variant
    
    For Each var設定値 In dic接続情報
    
        txt接続文字列 = txt接続文字列 & var設定値 & "=" & dic接続情報.Item(var設定値) & ";"
        
    Next var設定値

    f_Dictonary結合 = txt接続文字列

End Function
