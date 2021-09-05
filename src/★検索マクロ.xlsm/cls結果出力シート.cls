VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls結果出力シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************************************************************
' 定数
' *********************************************************************************************************************
'

' *********************************************************************************************************************
' 変数
' *********************************************************************************************************************
'
' 雛型シート
Private ws雛型シート As Worksheet

' 結果シート名
Private txt結果シート名 As String

' 結果ブック
Private wb結果ブック As Workbook

' サイレントモード
Private flgサイレントモード As Boolean

' 雛形最終列
Private lng雛形最終列 As Long

' 雛形開始行
Private lng雛形開始行 As Long

' 処理結果保持用
Dim var出力配列() As Variant

Public Property Get 雛形開始列() As Long

    雛形開始列 = lng雛形開始列

End Property

Public Property Get 雛形開始行() As Long

    雛形開始行 = lng雛形開始行

End Property


Public Property Get 出力配列() As Variant()

    出力配列 = var出力配列

End Property


' *********************************************************************************************************************
' 機能：コンストラクタ ※ 引数は指定できない
' *********************************************************************************************************************
'
Public Sub Class_Initialize()

End Sub

' *********************************************************************************************************************
' 機能：指定されたシートの内容をロードし、Dictionaryオブジェクトに設定する。
' *********************************************************************************************************************
'
Public Sub 初期化(arg雛型シート名 As String, arg結果シート名 As String, argサイレントモード As Boolean)

    Set ws雛型シート = ThisWorkbook.Sheets(arg雛型シート名)

    txt結果シート名 = arg結果シート名
    flgサイレントモード = argサイレントモード

    lng雛形開始行 = 最終行取得(ws雛型シート, False) + 1
    lng雛形最終列 = 最終列取得(ws雛型シート)

End Sub

' *********************************************************************************************************************
' * 機能　：結果出力の共通部分の処理
' *********************************************************************************************************************
'
Public Sub 結果記録(rng対象セル As Range, var出力内容 As Variant)

    Call reDimResult(lng雛形最終列, var出力配列)

    Dim lng列 As Long: lng列 = UBound(var出力配列, 2)

    ' フォルダ名
    var出力配列(0, lng列) = rng対象セル.Parent.Parent.Path
    ' ファイル名
    var出力配列(1, lng列) = rng対象セル.Parent.Parent.Name
    ' シート名
    var出力配列(2, lng列) = rng対象セル.Parent.Name
    ' セル座標
    var出力配列(3, lng列) = rng対象セル.Address(False, False)
    
    Dim i As Long
    
    For i = LBound(var出力内容) To UBound(var出力内容)
    
        If lng雛形最終列 < 4 + i Then
        
            log ("出力内容が雛形のサイズを超えています。" & _
                "出力内容：" & 4 + UBound(var出力内容) & "。" & _
                "雛形：" & lng雛形最終列 & "。")
                
        Else
            var出力配列(4 + i, lng列) = var出力内容(i)
            
        End If
    Next i
    
End Sub

' *********************************************************************************************************************
' * 機能　：結果ブックの作成
' *********************************************************************************************************************
'
Public Function 結果ブック作成(txtパス共通部 As String)
    
    If Not Not var出力配列 Then
    
        If UBound(var出力配列, 2) <> 0 Then
        
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
            wb結果ブック.ActiveSheet.Range("B1") = txtパス共通部

            wb結果ブック.ActiveSheet.Range( _
                Cells(lng最大行, 1), _
                Cells(UBound(var出力配列, 2) + lng雛形開始行, UBound(var出力配列) + 1)) = 二次元配列行列逆転(var出力配列)
            
            Dim lng最大列 As Long
            ' 書式コピー
            With wb結果ブック.ActiveSheet
                lng最大行 = 最終行取得(wb結果ブック.ActiveSheet, False)
                lng最大列 = 最終列取得(wb結果ブック.ActiveSheet, False)
                
                .Range(.Cells(lng雛形開始行, 1), .Cells(lng雛形開始行, lng最大列)).Copy
                .Range(.Cells(lng雛形開始行 + 1, 1), .Cells(lng最大行, lng最大列)).PasteSpecial (xlPasteFormats)
                
                Dim i As Long
                
                For i = lng雛形開始行 To lng最大行
                
                    ' 共通部を関数化
                    .Cells(i, 1) = "=B1 & """ & Replace(.Cells(i, 1), txtパス共通部, "") & """"

                    ' ハイパーリンク設定
                    Dim strHyperLink As String
                    strHyperLink = "=HYPERLINK(""[""&A" & i & "&""\""&B" & i & "&""]'""&" & _
                        "C" & i & "&""'!" & .Cells(i, 4) & """,""" & .Cells(i, 4) & """)"
            
                    .Range(.Cells(i, 4), .Cells(i, 4)).Value = strHyperLink
                Next
            End With
            
            
            ' "A1"を選択状態にする
            wb結果ブック.ActiveSheet.Cells(1, 1).Select
            
            ' シート名「処理結果」以外のシートを削除する
            Call 不要シート削除(wb結果ブック, txt結果シート名)
            
        ElseIf flg結果0件は異常 Then
            
            Call s_メッセージ通知("処理結果は0件です。", flgサイレントモード)

        End If
        
    ElseIf flg結果0件は異常 Then
    
        Call s_メッセージ通知("処理結果は0件です。", flgサイレントモード)

    End If
    
    Set 結果ブック作成 = wb結果ブック

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
        
        If ws雛型シート.Name = ThisWorkbook.Worksheets(i).Name Then
        
            ws雛型シート.Copy Before:=wbコピー先ブック.Sheets(1)
        End If
        
    Next i
    
    ' マクロを実行中のブックをアクティブにする
    Workbooks(wbコピー先ブック.Name).Sheets(ws雛型シート.Name).Activate
    
    ' シート名を「処理結果」に変更する
    Workbooks(wbコピー先ブック.Name).Sheets(ws雛型シート.Name).Name = txt結果シート名
    
End Sub

' *********************************************************************************************************************
' * 機能　：結果が設定されているか否かを返す
' *********************************************************************************************************************
'
Public Function 出力あり() As Boolean

    If Not Not var出力配列 Then
    
        If UBound(var出力配列, 2) <> 0 Then
        
            出力あり = True
            Exit Function
    
        End If
        
    End If

    出力あり = False

End Function

' *********************************************************************************************************************
' * 機能　：保存先が設定されている場合、保存して閉じる
' *********************************************************************************************************************
'
Public Sub 必要に応じて保存(obj設定値シート As cls設定値シート)

    If obj設定値シート.設定値リスト.Item("保存先") <> "" Then
            
        wb結果ブック.SaveAs (f_日時サフィックス付与( _
            obj設定値シート.設定値リスト.Item("保存先"), _
            obj設定値シート.設定値リスト.Item("保存時サフィックス")))

        ' サイレントモードの場合はファイルを閉じる
        If flgサイレントモード Then
    
            wb結果ブック.Close
    
        End If
           
    End If

End Sub


