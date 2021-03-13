Attribute VB_Name = "テンプレート処理"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' 定数
' ---------------------------------------------------------------------------------------------------------------------

Public Const TEMPLATE_SHEET_NAME = "雛形"

' 対象の拡張子
Private Const FILE_EXTENSION = "xls,xlsx,xlsm"

Private Const KEY_パス = "パス"
Private Const KEY_対象ブックシート = "対象ブックシート名"
Private Const KEY_対象外ブックシート = "対象外ブックシート名"

' ---------------------------------------------------------------------------------------------------------------------
' 変数
' ---------------------------------------------------------------------------------------------------------------------

' 設定値リスト
Public obj設定値シート As cls設定値シート

' 雛形最終列
Public lng雛形最終列 As Long

' 雛形開始行
Public lng雛形開始行 As Long


' *********************************************************************************************************************
' * 機能　：マクロ呼び出し時（シートからの指定用）
' *********************************************************************************************************************

Sub マクロ開始()

    Call init開始時刻

    log ("----------------------------------------------------------------------------------------------------")
    log ("マクロ開始")
    log ("----------------------------------------------------------------------------------------------------")
    
    Set obj設定値シート = New cls設定値シート
    obj設定値シート.ロード (ActiveSheet.Name)
    
    ' -----------------------------------------------------------------------------------------------------------------
    ' 初期化処理
    ' -----------------------------------------------------------------------------------------------------------------
    
    ' 固有処理（マクロ呼び出し元）側のシート情報を取得する。
    Dim wsマクロ呼び出し元シート As Worksheet
    Set wsマクロ呼び出し元シート = ActiveSheet

    ' ★ConcreateProcess側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
    Call 全体前処理(wsマクロ呼び出し元シート)

    lng雛形開始行 = 最終行取得(ThisWorkbook.Sheets(TEMPLATE_SHEET_NAME), False) + 1
    lng雛形最終列 = 最終列取得(ThisWorkbook.Sheets(TEMPLATE_SHEET_NAME))

    ' -----------------------------------------------------------------------------------------------------------------
    ' パスの存在チェック
    ' -----------------------------------------------------------------------------------------------------------------

    Dim txtパス As Variant

    With wsマクロ呼び出し元シート

        ' 対象ディレクトリ分ループ
        For Each txtパス In obj設定値シート.設定値リスト.Item(KEY_パス)
            
            ' ディレクトリまたは、ファイルの存在チェック
            If isDirectoryExist(CStr(txtパス)) < 0 Then
                
                MsgBox "以下のパスは存在しません。" + Chr(10) + "「" + txtパス + "」"
                End
            End If
        Next
    End With

    ' -----------------------------------------------------------------------------------------------------------------
    ' ファイル名の収集
    ' -----------------------------------------------------------------------------------------------------------------

    Call setステータスバー("対象ファイル集計中...")
    
    ' 処理対象の拡張子を設定する。
    Dim varファイル拡張子 As Variant
    varファイル拡張子 = Split(FILE_EXTENSION, ",")
    
    ' 処理対象のファイル名一覧（フルパス＆ファイル名）
    Dim txtパス一覧() As String
    
    '対象ディレクトリ分ループ
    For Each txtパス In obj設定値シート.設定値リスト.Item(KEY_パス)
            
        '指定の値がファイルの場合、その値をリストに追加し、
        ' ディレクトリの場合は、ファイル名の一覧を動的に取得して追加する。
        If isDirectoryExist(CStr(txtパス)) = 2 Then
            
            ' 指定の値がファイルだった場合、その値をリストに追加
            ' フルパス＆ファイル名を追加格納。
            Call 一次配列に値を追加(txtパス一覧, CStr(txtパス))
        Else
                    
            ' 現在のディレクトリ配下のファイル名を取得
            Call doRepeat(txtパス, varファイル拡張子, txtパス一覧, True)
                
        End If
    Next

    ' -----------------------------------------------------------------------------------------------------------------
    ' ファイル処理メソッドの呼び出し
    ' -----------------------------------------------------------------------------------------------------------------
    
    Call ファイル処理(txtパス一覧)
    
    ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
    Call 全体後処理(wsマクロ呼び出し元シート)
    
    MsgBox "処理が終了しました。（処理時間：" & get処理時刻() & ")"

    log ("----------------------------------------------------------------------------------------------------")
    log ("マクロ終了")
    log ("----------------------------------------------------------------------------------------------------")

End Sub

' *********************************************************************************************************************
' * 機能　：対象ファイルの処理を行う。
' * 引数　：varArray 配列
' * 戻り値：判定結果（1:配列/0:空の配列/-1:配列ではない）
' *********************************************************************************************************************
'
Function ファイル処理(txtパス一覧() As String)

    ' ファイル名の一覧が空だった場合、当Functionを中断する。
    If IsArrayEx(txtパス一覧) < 1 Then
        MsgBox "処理対象ファイルが存在しません。"
        Exit Function
    End If
    
    Dim defaultSaveFormat As Long
    defaultSaveFormat = Application.defaultSaveFormat
    
    Application.DisplayAlerts = False ' ファイルを開く際の警告を無効
    Application.ScreenUpdating = False ' 画面表示更新を無効
    
    ' 処理結果保持用
    Dim results() As Variant
    
    Dim index As Long, total As Long
        
    index = 1
    total = UBound(txtパス一覧) + 1
    
    Dim txtパス As Variant
    
    For Each txtパス In txtパス一覧
    
        ' -------------------------------------------------------------------------------------------------------------
        ' 対象ブックを開いて、全シート分の処理を行う。
        ' -------------------------------------------------------------------------------------------------------------

        Call setステータスバー("(" & index & "/" & total & ")" & ファイル名取得(CStr(txtパス)))
        index = index + 1
        
        Dim wb対象ブック As Workbook
        Set wb対象ブック = Workbooks.Open(txtパス, UpdateLinks:=0, IgnoreReadOnlyRecommended:=False)
        
        ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
        If ブックOPEN後処理(txtパス, wb対象ブック, results) Then
        
            ' ブックOPEN後処理の返り値がTrueの場合、シート毎の処理を続行する
        
            Dim var対象ブックシート As Variant
            var対象ブックシート = obj設定値シート.設定値リスト.Item(KEY_対象ブックシート)
            
            Dim var対象外ブックシート As Variant
            var対象外ブックシート = obj設定値シート.設定値リスト.Item(KEY_対象外ブックシート)
        
            Dim i As Long
            For i = 1 To wb対象ブック.Worksheets.Count
            
                Dim ws対象シート As Worksheet
                Set ws対象シート = wb対象ブック.Worksheets(i)
                
                Dim txtブックシート名 As String
                txtブックシート名 = "[" & wb対象ブック.Name & "]" & ws対象シート.Name
                
                ' 対象のブック／シートかチェック
                If Not f_配列含まれているかチェック(var対象ブックシート, txtブックシート名, False, True) Then
                
                    log ("対象ブックシートに不一致：" & txtブックシート名)
                
                ' 対象外のブック／シートでないかチェック
                ElseIf f_配列含まれているかチェック(var対象外ブックシート, txtブックシート名, False, True) Then
                
                    log ("対象外ブックシートに一致：" & txtブックシート名)
                Else
                
                    ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
                    Call シート毎処理(txtパス, ws対象シート, results)
                End If
                
            Next i
            
        End If
        
        ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
        Dim ファイルCLOSE方法区分値 As Long
        ファイルCLOSE方法区分値 = ブックCLOSE前処理(txtパス, wb対象ブック, results)
        
        If ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存しないで閉じる Then
            wb対象ブック.Close
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存して閉じる Then
            wb対象ブック.Save
            wb対象ブック.Close
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存しないで閉じない Then
            
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.保存して閉じない Then
            wb対象ブック.Save
        ElseIf ファイルCLOSE方法区分値 = ファイルCLOSE方法区分.処理中断 Then
            End
        End If
    Next
        
    ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
    ' 実行結果の編集（結果のマージ、並び替え、フィルタリング当）
    Call 実行結果内容編集処理(results)
    
    Dim wb結果ブック As Workbook
    
    If Not Not results Then
    
        If UBound(results, 2) <> 0 Then
        
            ' ファイルの保存形式をexcel2007形式（.xlsx)に変更
            Application.defaultSaveFormat = xlOpenXMLWorkbook
            
            Set wb結果ブック = Workbooks.Add
            
            ' 当ブックにシート「雛形」が用意されている場合、指定ブックの先頭にコピーした後、
            ' シート名を「処理結果」に変更する。（ない場合は新規作成ブックのsheet1を利用）
            Call 雛形シートコピー(wb結果ブック)
            
            ' 結果貼り付け行の取得。
            ' A列に値が設定されている行を、表題欄としてその行数を取得する
            Dim lng最大行 As Long
            With wb結果ブック.ActiveSheet.UsedRange
                lng最大行 = .Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
            End With
            
            ' 結果貼り付け行の設定。
            lng最大行 = lng最大行 + 1
            
            ' 結果貼り付け
            wb結果ブック.ActiveSheet.Range( _
                Cells(lng最大行, 1), _
                Cells(UBound(results, 2) + lng雛形開始行, UBound(results) + 1)) = 二次元配列行列逆転(results)
            
            Dim lng最大列 As Long
            ' 書式コピー
            With wb結果ブック.ActiveSheet
                lng最大行 = 最終行取得(wb結果ブック.ActiveSheet, False)
                lng最大列 = 最終列取得(wb結果ブック.ActiveSheet, False)
                
                .Range(.Cells(lng雛形開始行, 1), .Cells(lng雛形開始行, lng最大列)).Copy
                .Range(.Cells(lng雛形開始行 + 1, 1), .Cells(lng最大行, lng最大列)).PasteSpecial (xlPasteFormats)
                
                For i = lng雛形開始行 To lng最大行
                
                    ' ハイパーリンク設定
                    Dim strHyperLink As String
                    strHyperLink = "=HYPERLINK(""[""&A" & i & "&""\""&B" & i & "&""]""&" & _
                        "C" & i & "&""!" & .Cells(i, 4) & """,""" & .Cells(i, 4) & """)"
            
                    .Range(.Cells(i, 4), .Cells(i, 4)).Value = strHyperLink
                Next
            End With
            
            ' ★実装処理側の処理の呼び出し（呼び出し先のProcedure側ではツールごとの固有の実装を行う）
            Call 実行結果書式編集処理(wb結果ブック.ActiveSheet)
            
            ' "A1"を選択状態にする
            wb結果ブック.ActiveSheet.Cells(1, 1).Select
            
            ' シート名「処理結果」以外のシートを削除する
            Call 不要シート削除(wb結果ブック, RESULT_SHEET_NAME)
            
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
            wb結果ブック.Activate
        End If
    End If

End Function


' *********************************************************************************************************************
' * 機能　：当ブックのシート「雛形」を指定ブックの先頭にコピーした後、
' * 　　　　シート名を「処理結果」に変更する
' *********************************************************************************************************************
'
Sub 雛形シートコピー(wbコピー先ブック As Workbook)

    ' マクロ実行時のブックをアクティブにする
    ThisWorkbook.Activate
    
    ' シート「雛形」があった場合、指定ブックにコピー（一番前に挿入）する
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count ' シートの数分ループする
    
        Dim targetSheet As Worksheet
        Set targetSheet = ThisWorkbook.Worksheets(i)
        
        If TEMPLATE_SHEET_NAME = ThisWorkbook.Worksheets(i).Name Then
        
            ThisWorkbook.Sheets(TEMPLATE_SHEET_NAME).Copy Before:=wbコピー先ブック.Sheets(1)
        End If
        
    Next i
    
    ' マクロを実行中のブックをアクティブにする
    Workbooks(wbコピー先ブック.Name).Sheets(TEMPLATE_SHEET_NAME).Activate
    
    ' シート名を「処理結果」に変更する
    Workbooks(wbコピー先ブック.Name).Sheets(TEMPLATE_SHEET_NAME).Name = RESULT_SHEET_NAME
    
End Sub
