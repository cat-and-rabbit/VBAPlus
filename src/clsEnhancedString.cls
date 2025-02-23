VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnhancedString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----- Private Member Variables -----
' 内部で保持する文字列の変数
Private mValue As String
' メソッド呼び出し時にインプレース更新するか、新規インスタンスを生成するかのフラグ
Private mInPlaceUpdate As Boolean

' ----- Class LifeCycle Methods -----
' クラスの初期化処理
Private Sub Class_Initialize()
    mValue = ""   ' 初期値を空文字列に設定
    mInPlaceUpdate = False
End Sub

' クラス終了時の後処理（ここでは文字列をクリア）
Private Sub Class_Terminate()
    mValue = ""
End Sub

' ----- Public Properties and Methods -----
' プロパティ Get/Let と外部利用可能な各種文字列操作メソッド

' 外部から初期テキストとインプレース更新フラグを設定するメソッド
Public Sub Initialize(ByVal pInitialText As String, Optional ByVal pInPlaceUpdate As Boolean = False)
    mValue = pInitialText
    mInPlaceUpdate = pInPlaceUpdate
End Sub

' プロパティ Get: 現在の文字列の値を返す
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

' プロパティ Let: 文字列の値を更新する
Public Property Let Value(ByVal pNewText As String)
    mValue = pNewText
End Property

' プロパティ Get: 文字列の長さ（文字数）を返す
Public Property Get Length() As Long
    Length = VBA.Len(mValue)
End Property

' メソッド: 文字列を大文字に変換する
Public Function ToUpperCase() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    ' インプレース更新かどうかで、返すインスタンスを選択
    Set lvClass = GetInstance
    
    ' 変換処理（VBAのUCase関数）
    lvClass.Value = VBA.UCase(mValue)
    
    Set ToUpperCase = lvClass
End Function

' メソッド: 文字列を小文字に変換する
Public Function ToLowerCase() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    ' インプレース更新の判定
    Set lvClass = GetInstance
    
    ' 変換処理（VBAのLCase関数）
    lvClass.Value = VBA.LCase(mValue)
    
    Set ToLowerCase = lvClass
End Function

' メソッド: 文字列の前後の空白を除去する（Trim）
Public Function Trim() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' VBAのTrim関数で空白を除去
    lvClass.Value = VBA.Trim(mValue)
    
    Set Trim = lvClass
End Function

' メソッド: 文字列の先頭の空白を除去する（LTrim）
Public Function TrimStart() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    lvClass.Value = VBA.LTrim(mValue)
    
    Set TrimStart = lvClass
End Function

' メソッド: 文字列の末尾の空白を除去する（RTrim）
Public Function TrimEnd() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    lvClass.Value = VBA.RTrim(mValue)
    
    Set TrimEnd = lvClass
End Function

' メソッド: 文字列の一部を抽出する（Slice）
Public Function Slice(ByVal pStart As Long, Optional ByVal pEnd As Variant) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' インデックスの正規化（負の値や範囲外を調整）
    pStart = NormalizeIndex(pStart, VBA.Len(mValue))
    pEnd = NormalizeOptionalEnd(pEnd, VBA.Len(mValue))
    
    ' 開始位置が終了位置を超えている場合は空の結果を返す
    If pStart > pEnd Then
        Set Slice = lvClass
        Exit Function
    End If

    ' Mid関数を使って指定部分を抽出（VBAでは1オリジンのため+1）
    lvClass.Value = Mid(mValue, pStart + 1, pEnd - pStart)
    
    Set Slice = lvClass
End Function

' メソッド: 文字列の一部を置換または削除する（Splice）
Public Function Splice(ByVal pStart As Long, Optional ByVal pEnd As Variant, Optional ByVal pText As String = "") As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' インデックスの正規化
    pStart = NormalizeIndex(pStart, Len(mValue))
    pEnd = NormalizeOptionalEnd(pEnd, Len(mValue))
    
    ' 範囲が不正な場合はそのまま返す
    If pStart > pEnd Then
        Set Splice = lvClass
        Exit Function
    End If

    ' 指定範囲をpTextで置換
    lvClass.Value = Mid(mValue, 1, pStart) & pText & Mid(mValue, pEnd + 1)
    
    Set Splice = lvClass
End Function

' メソッド: 指定した文字列が含まれているかを確認する
Public Function Includes(ByVal pFindText As String) As Boolean
    Includes = (InStr(1, mValue, pFindText, vbTextCompare) > 0)
End Function

' メソッド: 指定した文字列の位置を返す（0オリジン）
Public Function IndexOf(ByVal pFindText As String) As Long
    IndexOf = InStr(1, mValue, pFindText, vbTextCompare) - 1
End Function

' メソッド: 文字列が指定した文字列で始まるかを判定する
Public Function StartsWith(ByVal pFindText As String) As Boolean
    If Len(pFindText) > Len(mValue) Then
        StartsWith = False
    Else
        StartsWith = (Left(mValue, VBA.Len(pFindText)) = pFindText)
    End If
End Function

' メソッド: 文字列が指定した文字列で終わるかを判定する
Public Function EndsWith(ByVal pFindText As String) As Boolean
    If Len(pFindText) > Len(mValue) Then
        EndsWith = False
    Else
        EndsWith = (Right(mValue, VBA.Len(pFindText)) = pFindText)
    End If
End Function

' メソッド: 最初に見つかった指定文字列を置換する
Public Function Replace(ByVal pFindText As String, ByVal pReplaceText As String) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' VBAのReplace関数で、1回だけ置換を実施
    lvClass.Value = VBA.Replace(mValue, pFindText, pReplaceText, 1, 1)
    
    Set Replace = lvClass
End Function

' メソッド: 全ての一致する指定文字列を置換する
Public Function ReplaceAll(ByVal pFindText As String, ByVal pReplaceText As String) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' VBAのReplace関数で、全ての一致部分を置換
    lvClass.Value = VBA.Replace(mValue, pFindText, pReplaceText)
    
    Set ReplaceAll = lvClass
End Function

' メソッド: 指定したデリミタで文字列を分割して配列を返す
Public Function Split(ByVal pDelimiter As String) As Variant
    Split = VBA.Split(mValue, pDelimiter)
End Function

' メソッド: 指定した長さになるよう、前方にパディング文字を追加する
Public Function PadStart(ByVal pLength As Long, Optional ByVal pPadString As String = " ") As clsEnhancedString
    Dim lvPadCount As Long
    Dim lvRepeat As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    ' パディング用のクラスインスタンスを生成
    Set lvRepeat = New clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' パディング文字が空文字の場合はスペースを利用
    If Len(pPadString) = 0 Then
        pPadString = " "
    End If
    
    ' 必要なパディング回数を算出（Ceil関数で端数を切り上げ）
    lvPadCount = Ceil((pLength - Len(mValue)) / Len(pPadString))
    lvRepeat.Value = pPadString
    ' 指定回数分繰り返したパディング文字列を生成し、余分な部分をSliceで調整して先頭に連結
    lvClass.Value = lvRepeat.Repeat(lvPadCount).Slice(0, pLength - Len(mValue)) & mValue
    
    Set PadStart = lvClass
End Function

' メソッド: 指定した長さになるよう、後方にパディング文字を追加する
Public Function PadEnd(ByVal pLength As Long, Optional ByVal pPadString As String = " ") As clsEnhancedString
    Dim lvPadCount As Long
    Dim lvRepeat As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvRepeat = New clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' パディング文字が空文字の場合はスペースを利用
    If Len(pPadString) = 0 Then
        pPadString = " "
    End If
    
    lvPadCount = Ceil((pLength - Len(mValue)) / Len(pPadString))
    lvRepeat.Value = pPadString
    ' 元の文字列に対して、必要なパディングを追加（余分な部分はSliceで調整）
    lvClass.Value = mValue & lvRepeat.Repeat(lvPadCount).Slice(0, pLength - Len(mValue))
    
    Set PadEnd = lvClass
End Function

' メソッド: 現在の文字列を指定回数繰り返す
Public Function Repeat(ByVal pCount As Long) As clsEnhancedString
    ' 負の回数の場合はエラーを発生させる
    If pCount < 0 Then
        Err.Raise 9, "clsEnhancedString.Repeat"
    End If
    
    Dim lvItems() As Variant
    Dim lvClass As clsEnhancedString
    Dim i As Long
    
    ReDim lvItems(0 To pCount - 1)
    
    Set lvClass = GetInstance
    
    ' 配列に同じ文字列を格納してJoinで連結する
    For i = 0 To pCount - 1
        lvItems(i) = mValue
    Next i
    
    lvClass.Value = VBA.Join(lvItems, "")
    
    Set Repeat = lvClass
End Function

' メソッド: テンプレート文字列内のプレースホルダーを指定の引数で置換する
Public Function Template(ParamArray pTexts() As Variant) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    Dim i As Long
    
    Set lvClass = GetInstance
    
    lvClass.Value = mValue
    
    ' {0}, {1}, ... の形式でプレースホルダーを置換
    For i = LBound(pTexts) To UBound(pTexts)
        lvClass.Value = VBA.Replace(lvClass.Value, "{" & i & "}", CStr(pTexts(i)))
    Next i
    
    Set Template = lvClass
End Function

' メソッド: 文字列を逆順にする
Public Function Reverse() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    Dim i As Long
    
    Set lvClass = GetInstance
    
    ' VBAのStrReverse関数を使用
    lvClass.Value = VBA.StrReverse(mValue)
    
    Set Reverse = lvClass
End Function

' メソッド: 正規表現パターンに対してマッチするかテストする
Public Function Test(ByVal pPattern As String, Optional ByVal pIgnoreCase As Boolean = True) As Boolean
    Dim lvRegExp As Object
    
    ' VBScript.RegExp オブジェクトの生成
    Set lvRegExp = CreateObject("VBScript.RegExp")
    
    With lvRegExp
        .Pattern = pPattern
        .IgnoreCase = pIgnoreCase
        .Global = True
    End With
    
    Test = lvRegExp.Test(mValue)
End Function

' メソッド: 正規表現による置換を実施する
Public Function ReplaceRegex(ByVal pPattern As String, ByVal pReplacement As String, Optional ByVal pIgnoreCase As Boolean = True, Optional ByVal pGlobal As Boolean = True) As clsEnhancedString
    Dim lvRegExp As Object
    Dim lvClass As clsEnhancedString
    
    Set lvRegExp = CreateObject("VBScript.RegExp")
    
    Set lvClass = GetInstance
    
    With lvRegExp
        .Pattern = pPattern
        .IgnoreCase = pIgnoreCase
        .Global = pGlobal
    End With
    
    ' 正規表現を使って置換
    lvClass.Value = lvRegExp.Replace(mValue, pReplacement)
    
    Set ReplaceRegex = lvClass
End Function

' メソッド: 正規表現にマッチする部分を取得する
Public Function Match(ByVal pPattern As String, Optional ByVal pIgnoreCase As Boolean = True, Optional ByVal pGlobal As Boolean = True) As Object
    Dim lvRegExp As Object
    Dim i As Long
    
    Set lvRegExp = CreateObject("VBScript.RegExp")
    
    With lvRegExp
        .Pattern = pPattern
        .IgnoreCase = pIgnoreCase
        .Global = pGlobal
    End With
    
    ' Executeでマッチしたコレクションを返す
    Set Match = lvRegExp.Execute(mValue)
End Function

' ----- Private Helper Methods -----
' 内部処理用の補助メソッド

' プライベートメソッド: 数値を切り上げする（Ceiling関数）
Private Function Ceil(ByVal pNumber As Double, Optional ByVal pSignificance As Double = 1) As Double
    If pSignificance = 0 Then
        Ceil = pNumber
    Else
        ' 負の数に対しても切り上げ処理を適用
        Ceil = -Int(-pNumber / pSignificance) * pSignificance
    End If
End Function

' プライベートメソッド: インデックスを正規化する
Private Function NormalizeIndex(ByVal pIndex As Long, ByVal pFullLength As Long) As Long
    ' 負のインデックスは文字列長を加算して正のインデックスに変換
    If pIndex < 0 Then
        pIndex = pIndex + pFullLength
    End If

    ' 範囲外の調整
    If pIndex < 0 Then
        pIndex = 0
    End If
    
    If pIndex >= pFullLength Then
        pIndex = pFullLength
    End If
    
    NormalizeIndex = pIndex
End Function

' プライベートメソッド: 終了インデックス（省略時は文字列末尾）を正規化する
Private Function NormalizeOptionalEnd(ByVal pEnd As Variant, ByVal pFullLength As Long) As Long
    If IsMissing(pEnd) Or Not IsNumeric(pEnd) Then
        NormalizeOptionalEnd = pFullLength
    Else
        NormalizeOptionalEnd = NormalizeIndex(CLng(pEnd), pFullLength)
    End If
End Function

' プライベートメソッド: pInPlaceUpdate の値に応じたインスタンスを返す
Private Function GetInstance() As clsEnhancedString
    If mInPlaceUpdate Then
        ' インプレース更新の場合、現在のインスタンスを返す
        Set GetInstance = Me
    Else
        ' それ以外は新たなインスタンスを生成して返す
        Set GetInstance = New clsEnhancedString
    End If
End Function

