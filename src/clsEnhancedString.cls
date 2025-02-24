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

'==============================================
' /**
'  * クラス clsEnhancedString
'  * 文字列操作のための拡張機能を提供するクラス
'  */
'==============================================

' 内部で保持する文字列の変数
Private mValue As String
' メソッド呼び出し時にインプレース更新するか、新規インスタンスを生成するかのフラグ
Private mInPlaceUpdate As Boolean

'==============================================
' /**
'  * クラスの初期化処理（コンストラクタ）
'  *
'  * @constructor
'  */
'==============================================
Private Sub Class_Initialize()
    mValue = ""   ' 初期値を空文字列に設定
    mInPlaceUpdate = False
End Sub

'==============================================
' /**
'  * インスタンスの初期化を行う
'  *
'  * @param {String} pInitialText - 初期テキスト
'  * @param {Boolean} [pInPlaceUpdate=false] - インプレース更新フラグ（既存インスタンス更新の場合はTrue）
'  */
'==============================================
Public Sub Initialize(ByVal pInitialText As String, Optional ByVal pInPlaceUpdate As Boolean = False)
    mValue = pInitialText
    mInPlaceUpdate = pInPlaceUpdate
End Sub

'==============================================
' /**
'  * クラス終了時の後処理（デストラクタ）
'  * 文字列をクリアする
'  */
'==============================================
Private Sub Class_Terminate()
    mValue = ""
End Sub

'==============================================
' /**
'  * 現在の文字列の値を取得するプロパティ
'  *
'  * @return {String} 現在の文字列
'  */
'==============================================
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

'==============================================
' /**
'  * 文字列の値を設定するプロパティ
'  *
'  * @param {String} pNewText - 新しい文字列
'  */
'==============================================
Public Property Let Value(ByVal pNewText As String)
    mValue = pNewText
End Property

'==============================================
' /**
'  * 文字列の長さ（文字数）を返すプロパティ
'  *
'  * @return {Long} 文字列の長さ
'  */
'==============================================
Public Property Get Length() As Long
    Length = VBA.Len(mValue)
End Property

'==============================================
' /**
'  * 文字列を大文字に変換する
'  *
'  * @return {clsEnhancedString} 変換後の文字列を保持するインスタンス
'  */
'==============================================
Public Function ToUpperCase() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    ' インプレース更新かどうかで、返すインスタンスを選択
    Set lvClass = GetInstance
    
    ' 変換処理（VBAのUCase関数）
    lvClass.Value = VBA.UCase(mValue)
    
    Set ToUpperCase = lvClass
End Function

'==============================================
' /**
'  * 文字列を小文字に変換する
'  *
'  * @return {clsEnhancedString} 変換後の文字列を保持するインスタンス
'  */
'==============================================
Public Function ToLowerCase() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    ' インプレース更新の判定
    Set lvClass = GetInstance
    
    ' 変換処理（VBAのLCase関数）
    lvClass.Value = VBA.LCase(mValue)
    
    Set ToLowerCase = lvClass
End Function

'==============================================
' /**
'  * 文字列の前後の空白を除去する（Trim）
'  *
'  * @return {clsEnhancedString} トリム後の文字列を保持するインスタンス
'  */
'==============================================
Public Function Trim() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' VBAのTrim関数で空白を除去
    lvClass.Value = VBA.Trim(mValue)
    
    Set Trim = lvClass
End Function

'==============================================
' /**
'  * 文字列の先頭の空白を除去する（LTrim）
'  *
'  * @return {clsEnhancedString} 先頭トリム後の文字列を保持するインスタンス
'  */
'==============================================
Public Function TrimStart() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    lvClass.Value = VBA.LTrim(mValue)
    
    Set TrimStart = lvClass
End Function

'==============================================
' /**
'  * 文字列の末尾の空白を除去する（RTrim）
'  *
'  * @return {clsEnhancedString} 末尾トリム後の文字列を保持するインスタンス
'  */
'==============================================
Public Function TrimEnd() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    lvClass.Value = VBA.RTrim(mValue)
    
    Set TrimEnd = lvClass
End Function

'==============================================
' /**
'  * 文字列の一部を抽出する（Slice）
'  *
'  * @param {Long} pStart - 開始インデックス（0オリジン）
'  * @param {Variant} [pEnd] - 終了インデックス（省略時は文字列末尾）
'  * @return {clsEnhancedString} 抽出された文字列を保持するインスタンス
'  */
'==============================================
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

'==============================================
' /**
'  * 文字列の一部を置換または削除する（Splice）
'  *
'  * @param {Long} pStart - 置換開始インデックス（0オリジン）
'  * @param {Variant} [pEnd] - 置換終了インデックス（省略時は文字列末尾）
'  * @param {String} [pText=""] - 置換後の文字列（空文字の場合は削除）
'  * @return {clsEnhancedString} 置換結果の文字列を保持するインスタンス
'  */
'==============================================
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

'==============================================
' /**
'  * 指定した文字列が含まれているかを確認する
'  *
'  * @param {String} pFindText - 検索する文字列
'  * @return {Boolean} 含まれていればTrue、含まれていなければFalse
'  */
'==============================================
Public Function Includes(ByVal pFindText As String) As Boolean
    Includes = (InStr(1, mValue, pFindText, vbTextCompare) > 0)
End Function

'==============================================
' /**
'  * 指定した文字列の位置を返す（0オリジン）
'  *
'  * @param {String} pFindText - 検索する文字列
'  * @return {Long} 検索文字列の開始位置（0オリジン）。見つからなければ -1 を返す
'  */
'==============================================
Public Function IndexOf(ByVal pFindText As String) As Long
    IndexOf = InStr(1, mValue, pFindText, vbTextCompare) - 1
End Function

'==============================================
' /**
'  * 文字列が指定した文字列で始まるかを判定する
'  *
'  * @param {String} pFindText - 判定に使用する文字列
'  * @return {Boolean} 指定した文字列で始まる場合はTrue、そうでなければFalse
'  */
'==============================================
Public Function StartsWith(ByVal pFindText As String) As Boolean
    If Len(pFindText) > Len(mValue) Then
        StartsWith = False
    Else
        StartsWith = (Left(mValue, VBA.Len(pFindText)) = pFindText)
    End If
End Function

'==============================================
' /**
'  * 文字列が指定した文字列で終わるかを判定する
'  *
'  * @param {String} pFindText - 判定に使用する文字列
'  * @return {Boolean} 指定した文字列で終わる場合はTrue、そうでなければFalse
'  */
'==============================================
Public Function EndsWith(ByVal pFindText As String) As Boolean
    If Len(pFindText) > Len(mValue) Then
        EndsWith = False
    Else
        EndsWith = (Right(mValue, VBA.Len(pFindText)) = pFindText)
    End If
End Function

'==============================================
' /**
'  * 最初に見つかった指定文字列を置換する
'  *
'  * @param {String} pFindText - 置換対象の文字列
'  * @param {String} pReplaceText - 置換後の文字列
'  * @return {clsEnhancedString} 置換結果の文字列を保持するインスタンス
'  */
'==============================================
Public Function Replace(ByVal pFindText As String, ByVal pReplaceText As String) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' VBAのReplace関数で、1回だけ置換を実施
    lvClass.Value = VBA.Replace(mValue, pFindText, pReplaceText, 1, 1)
    
    Set Replace = lvClass
End Function

'==============================================
' /**
'  * 全ての一致する指定文字列を置換する
'  *
'  * @param {String} pFindText - 置換対象の文字列
'  * @param {String} pReplaceText - 置換後の文字列
'  * @return {clsEnhancedString} 置換結果の文字列を保持するインスタンス
'  */
'==============================================
Public Function ReplaceAll(ByVal pFindText As String, ByVal pReplaceText As String) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    Set lvClass = GetInstance
    
    ' VBAのReplace関数で、全ての一致部分を置換
    lvClass.Value = VBA.Replace(mValue, pFindText, pReplaceText)
    
    Set ReplaceAll = lvClass
End Function

'==============================================
' /**
'  * 指定したデリミタで文字列を分割して配列を返す
'  *
'  * @param {String} pDelimiter - 分割に使用する区切り文字
'  * @return {Variant} 分割された文字列の配列
'  */
'==============================================
Public Function Split(ByVal pDelimiter As String) As Variant
    Split = VBA.Split(mValue, pDelimiter)
End Function

'==============================================
' /**
'  * 指定した長さになるよう、前方にパディング文字を追加する
'  *
'  * @param {Long} pLength - 最終的な文字列の長さ
'  * @param {String} [pPadString=" "] - パディングに使用する文字列
'  * @return {clsEnhancedString} パディング後の文字列を保持するインスタンス
'  */
'==============================================
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

'==============================================
' /**
'  * 指定した長さになるよう、後方にパディング文字を追加する
'  *
'  * @param {Long} pLength - 最終的な文字列の長さ
'  * @param {String} [pPadString=" "] - パディングに使用する文字列
'  * @return {clsEnhancedString} パディング後の文字列を保持するインスタンス
'  */
'==============================================
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

'==============================================
' /**
'  * 現在の文字列を指定回数繰り返す
'  *
'  * @param {Long} pCount - 繰り返す回数（負の値はエラー）
'  * @return {clsEnhancedString} 繰り返し後の文字列を保持するインスタンス
'  * @throws エラー番号 9（pCountが負の場合）
'  */
'==============================================
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

'==============================================
' /**
'  * テンプレート文字列内のプレースホルダーを指定の引数で置換する
'  *
'  * @param {...Variant} pTexts - プレースホルダーに置換する引数群（{0}, {1}, ...）
'  * @return {clsEnhancedString} 置換後の文字列を保持するインスタンス
'  */
'==============================================
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

'==============================================
' /**
'  * 文字列を逆順にする
'  *
'  * @return {clsEnhancedString} 逆順にされた文字列を保持するインスタンス
'  */
'==============================================
Public Function Reverse() As clsEnhancedString
    Dim lvClass As clsEnhancedString
    Dim i As Long
    
    Set lvClass = GetInstance
    
    ' VBAのStrReverse関数を使用
    lvClass.Value = VBA.StrReverse(mValue)
    
    Set Reverse = lvClass
End Function

'==============================================
' /**
'  * 正規表現パターンに対してマッチするかテストする
'  *
'  * @param {String} pPattern - 正規表現パターン
'  * @param {Boolean} [pIgnoreCase=true] - 大文字小文字の区別を無視するかどうか
'  * @return {Boolean} パターンにマッチすればTrue、そうでなければFalse
'  */
'==============================================
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

'==============================================
' /**
'  * 正規表現による置換を実施する
'  *
'  * @param {String} pPattern - 正規表現パターン
'  * @param {String} pReplacement - 置換後の文字列
'  * @param {Boolean} [pIgnoreCase=true] - 大文字小文字の区別を無視するかどうか
'  * @param {Boolean} [pGlobal=true] - 全体で置換するかどうか
'  * @return {clsEnhancedString} 置換後の文字列を保持するインスタンス
'  */
'==============================================
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

'==============================================
' /**
'  * 正規表現にマッチする部分を取得する
'  *
'  * @param {String} pPattern - 正規表現パターン
'  * @param {Boolean} [pIgnoreCase=true] - 大文字小文字の区別を無視するかどうか
'  * @param {Boolean} [pGlobal=true] - 全体で検索するかどうか
'  * @return {Object} マッチした部分のコレクション（VBScript.RegExpの結果）
'  */
'==============================================
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

'==============================================
' /**
'  * 数値を切り上げする（Ceiling関数）
'  *
'  * @private
'  * @param {Double} pNumber - 対象の数値
'  * @param {Double} [pSignificance=1] - 切り上げの単位
'  * @return {Double} 切り上げ後の数値
'  */
'==============================================
Private Function Ceil(ByVal pNumber As Double, Optional ByVal pSignificance As Double = 1) As Double
    If pSignificance = 0 Then
        Ceil = pNumber
    Else
        ' 負の数に対しても切り上げ処理を適用
        Ceil = -Int(-pNumber / pSignificance) * pSignificance
    End If
End Function

'==============================================
' /**
'  * インデックスを正規化する
'  *
'  * @private
'  * @param {Long} pIndex - 元のインデックス（0オリジン）
'  * @param {Long} pFullLength - 文字列の全長
'  * @return {Long} 正規化されたインデックス
'  */
'==============================================
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

'==============================================
' /**
'  * 終了インデックス（省略時は文字列末尾）を正規化する
'  *
'  * @private
'  * @param {Variant} pEnd - 終了インデックスまたは省略値
'  * @param {Long} pFullLength - 文字列の全長
'  * @return {Long} 正規化された終了インデックス
'  */
'==============================================
Private Function NormalizeOptionalEnd(ByVal pEnd As Variant, ByVal pFullLength As Long) As Long
    If IsMissing(pEnd) Or Not IsNumeric(pEnd) Then
        NormalizeOptionalEnd = pFullLength
    Else
        NormalizeOptionalEnd = NormalizeIndex(CLng(pEnd), pFullLength)
    End If
End Function

'==============================================
' /**
'  * pInPlaceUpdate の値に応じたインスタンスを返す
'  *
'  * @private
'  * @return {clsEnhancedString} インプレース更新の場合は現在のインスタンス、そうでなければ新規インスタンス
'  */
'==============================================
Private Function GetInstance() As clsEnhancedString
    If mInPlaceUpdate Then
        ' インプレース更新の場合、現在のインスタンスを返す
        Set GetInstance = Me
    Else
        ' それ以外は新たなインスタンスを生成して返す
        Set GetInstance = New clsEnhancedString
    End If
End Function
