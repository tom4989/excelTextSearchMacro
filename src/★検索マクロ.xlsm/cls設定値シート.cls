VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls設定値シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************************************************************
' 定数
' *********************************************************************************************************************
'
Const cnstシート名 = "★設定値"


' *********************************************************************************************************************
' 変数
' *********************************************************************************************************************
'
Private dic設定値リスト As Object

Public Property Get 設定値リスト() As Object

    If dic設定値リスト Is Nothing Then
    
        Call MsgBox("cls設定値シート：設定値がロードされていません。", vbExclamation)
        End
    
    End If
    
    Set 設定値リスト = dic設定値リスト
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
Public Sub ロード(Optional txt設定値シート As String = cnstシート名)

    ' 辞書初期化
    Set dic設定値リスト = CreateObject("Scripting.Dictionary")

    Dim varシートの値全体() As Variant
    varシートの値全体 = シート内容取得(ThisWorkbook.Sheets(txt設定値シート))
    
    Dim i As Long
    
    For i = LBound(varシートの値全体) To UBound(varシートの値全体)
            
        ' A列に★が設定されている行のみ対象
        If varシートの値全体(i, 1) = "★" Then
        
            Dim j As Long, lng開始列 As Long, lng最終列 As Long
            lng開始列 = -1
            lng最終列 = -1

            ' 要素数を数える
            For j = LBound(varシートの値全体, 2) + 1 To UBound(varシートの値全体, 2)
        
                If varシートの値全体(i, j) <> "" Then
                
                    If lng開始列 = -1 Then
                    
                        lng開始列 = j
                    
                    End If
                
                    lng最終列 = j
                End If
            Next j
    
            ' 設定値行が2階層以上設定されている場合
            If lng開始列 < lng最終列 Then
        
                Dim objDic As Object
                Set objDic = dic設定値リスト
        
                ' 1行に記載された情報を解析する
                For j = lng開始列 To lng最終列 - 1
        
                    ' 処理中の列が、最後の列でない場合
                    If j < lng最終列 - 1 Then
                
                        If objDic.Exists(varシートの値全体(i, j)) Then
                    
                            Set objDic = objDic.Item(varシートの値全体(i, j))
                    
                        Else
                            Dim objNewDic As Object
                            Set objNewDic = CreateObject("Scripting.Dictionary")
                            Call objDic.Add(varシートの値全体(i, j), objNewDic)
                        
                            Set objDic = objNewDic

                        End If
                    
                    ' 処理中の列が、最後の列の1つ前の場合
                    Else
                    
                        Dim var設定値() As Variant
                    
                        ' 既に登録済の場合は、文字列のリストに設定値を追加する
                        If objDic.Exists(varシートの値全体(i, j)) Then
                            
                            var設定値 = objDic.Item(varシートの値全体(i, j))
                            ReDim Preserve var設定値(UBound(var設定値) + 1)
                            var設定値(UBound(var設定値)) = varシートの値全体(i, j + 1)
                    
                            objDic.Item(varシートの値全体(i, j)) = var設定値
                    
                        ' 次の行を確認し、同一階層の場合、設定値を文字列でなく文字列のリストにする
                        ElseIf 階層一致確認(varシートの値全体, i, j) Then
                        
                            ReDim var設定値(0)
                            var設定値(0) = varシートの値全体(i, j + 1)
                            Call objDic.Add(varシートの値全体(i, j), var設定値)
                            
                        Else
                    
                            Call objDic.Add(varシートの値全体(i, j), varシートの値全体(i, j + 1))
                    
                        End If
                        
                        Exit For
                    
                    End If
                Next j
            End If
        End If
    Next i
    
End Sub

' *********************************************************************************************************************
' 機能：現在の設定値と、次の行の設定値が同一の項目か確認する
' *********************************************************************************************************************
'
Private Function 階層一致確認(ByRef var検証対象() As Variant, lng行数 As Long, lng列数 As Long) As Boolean

    ' 先の行がない場合、Falseを返す
    If lng行数 = UBound(var検証対象) Then
    
        階層一致確認 = False
    End If

    Dim i, j As Long
    
    For j = 1 To lng列数
    
        ' 一致しない要素が存在した時点でFalse
        If var検証対象(lng行数, j) <> var検証対象(lng行数 + 1, j) Then
        
            階層一致確認 = False
            Exit Function
            
        End If
    Next j
    
    階層一致確認 = True

End Function
