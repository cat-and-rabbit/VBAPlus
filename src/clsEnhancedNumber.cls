VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnhancedNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================
' /**
'  * clsEnhancedNumber クラス
'  * 数値を拡張し、四則演算や各種数学的操作を提供するクラス
'  *
'  * @class clsEnhancedNumber
'  */
'==============================================

' 内部で保持する数値（Variant型）
Private mValue As Variant
' メソッド呼び出し時にインプレース更新するか、新規インスタンスを生成するかのフラグ
Private mInPlaceUpdate As Boolean

'==============================================
' /**
'  * クラス初期化処理
'  * 内部の数値を 0 に設定する
'  *
'  * @constructor
'  */
'==============================================
Private Sub Class_Initialize()
    mValue = 0
End Sub

'==============================================
' /**
'  * 数値の取得・設定を行うプロパティ
'  *
'  * @property {Variant} Value
'  */
'==============================================
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

Public Property Let Value(ByVal pValue As Variant)
    If Not IsNumeric(pValue) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Value", "Value is not a number"
    End If

    mValue = pValue
End Property

'==============================================
' /**
'  * クラスの初期化を行う
'  * 内部の数値に指定された初期値を設定する
'  *
'  * @param {Variant} pInitialNumber - 初期化する数値
'  * @param {Boolean} [pInPlaceUpdate=false] - インプレース更新フラグ（既存インスタンス更新の場合はTrue）
'  */
'==============================================
Public Sub Initialize(ByVal pInitialNumber As Variant, Optional ByVal pInPlaceUpdate As Boolean = False)
    If Not IsNumeric(pInitialNumber) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Initialize", "Initial number is not a number"
    End If

    mValue = pInitialNumber
    mInPlaceUpdate = pInPlaceUpdate
End Sub

'==============================================
' /**
'  * Add メソッド
'  * 引数の数値と内部の数値を加算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Variant} pNumber - 加算する数値
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Add(ByVal pNumber As Variant) As clsEnhancedNumber
    If Not IsNumeric(pNumber) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Add", "pNumber is not a number"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = mValue + pNumber

    Set Add = lvResult
End Function

'==============================================
' /**
'  * Subtract メソッド
'  * 引数の数値を内部の数値から減算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Variant} pNumber - 減算する数値
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Subtract(ByVal pNumber As Variant) As clsEnhancedNumber
    If Not IsNumeric(pNumber) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Subtract", "pNumber is not a number"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = mValue - pNumber

    Set Subtract = lvResult
End Function

'==============================================
' /**
'  * Multiply メソッド
'  * 引数の数値と内部の数値を乗算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Variant} pNumber - 乗算する数値
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Multiply(ByVal pNumber As Variant) As clsEnhancedNumber
    If Not IsNumeric(pNumber) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Multiply", "pNumber is not a number"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = mValue * pNumber

    Set Multiply = lvResult
End Function

'==============================================
' /**
'  * Divide メソッド
'  * 引数の数値で内部の数値を除算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Variant} pNumber - 除算する数値
'  * @throws {Error} 除数が数値ではないか0の場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Divide(ByVal pNumber As Variant) As clsEnhancedNumber
    If Not IsNumeric(pNumber) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Divide", "pNumber is not a number"
    End If

    If pNumber = 0 Then
        Err.Raise vbObjectError + 1001, "clsEnhancedNumber.Divide", "pNumber is zero"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = mValue / pNumber

    Set Divide = lvResult
End Function

'==============================================
' /**
'  * Pow メソッド
'  * 引数を指数として、内部の数値の累乗を計算する
'  *
'  * @param {Variant} pExponent - 指数
'  * @throws {Error} 指数が数値ではない場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Pow(ByVal pExponent As Variant) As clsEnhancedNumber
    If Not IsNumeric(pExponent) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Pow", "pExponent is not a number"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = mValue ^ pExponent

    Set Pow = lvResult
End Function

'==============================================
' /**
'  * Sqrt メソッド
'  * 内部の数値の平方根を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @throws {Error} 負の数の場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Sqrt() As clsEnhancedNumber
    If mValue < 0 Then
        Err.Raise vbObjectError + 1002, "clsEnhancedNumber.Sqrt", "Square root defined for non-negative numbers only"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Sqr(mValue)

    Set Sqrt = lvResult
End Function

'==============================================
' /**
'  * Round メソッド
'  * 指定された桁数で内部の数値を四捨五入し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Integer} pNumDigits - 四捨五入する小数点以下の桁数
'  * @throws {Error} 桁数が負の場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Round(ByVal pNumDigits As Integer) As clsEnhancedNumber
    If pNumDigits < 0 Then
        Err.Raise vbObjectError + 1003, "clsEnhancedNumber.Round", "Number of digits must be non-negative"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Round(mValue, pNumDigits)

    Set Round = lvResult
End Function

'==============================================
' /**
'  * Absolute メソッド
'  * 内部の数値の絶対値を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Absolute() As clsEnhancedNumber
    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Abs(mValue)

    Set Absolute = lvResult
End Function

'==============================================
' /**
'  * ToString メソッド
'  * 内部の数値を文字列に変換した clsEnhancedString オブジェクトを返す
'  *
'  * @return {clsEnhancedString} 内部の数値を文字列に変換した clsEnhancedString オブジェクト
'  */
'==============================================
Public Function ToString() As clsEnhancedString
    Dim lvResult As clsEnhancedString

    Set lvResult = New clsEnhancedString
    lvResult.Initialize CStr(mValue), mInPlaceUpdate

    Set ToString = lvResult
End Function

'==============================================
' /**
'  * Sin メソッド
'  * 内部の数値のサイン値を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Sin() As clsEnhancedNumber
    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Sin(mValue)

    Set Sin = lvResult
End Function

'==============================================
' /**
'  * Cos メソッド
'  * 内部の数値のコサイン値を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Cos() As clsEnhancedNumber
    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Cos(mValue)

    Set Cos = lvResult
End Function

'==============================================
' /**
'  * Tan メソッド
'  * 内部の数値のタンジェント値を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Tan() As clsEnhancedNumber
    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Tan(mValue)

    Set Tan = lvResult
End Function

'==============================================
' /**
'  * LogE メソッド
'  * 内部の数値の自然対数を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function LogE() As clsEnhancedNumber
    If mValue <= 0 Then
        Err.Raise vbObjectError + 1004, "clsEnhancedNumber.LogE", "Logarithm defined for positive numbers only"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Log(mValue)

    Set LogE = lvResult
End Function

'==============================================
' /**
'  * Log10 メソッド
'  * 内部の数値の常用対数（底10）を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @throws {Error} 数値が0以下の場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Log10() As clsEnhancedNumber
    If mValue <= 0 Then
        Err.Raise vbObjectError + 1004, "clsEnhancedNumber.Log10", "Logarithm defined for positive numbers only"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Log(mValue) / VBA.Log(10)

    Set Log10 = lvResult
End Function

'==============================================
' /**
'  * Exp メソッド
'  * 内部の数値の指数関数（e^x）を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Exp() As clsEnhancedNumber
    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = VBA.Exp(mValue)

    Set Exp = lvResult
End Function

'==============================================
' /**
'  * Modulo メソッド
'  * 内部の数値を引数で割った剰余を計算し、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Variant} pNumber - 剰余演算に使用する数値
'  * @throws {Error} 除数が0の場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Modulo(ByVal pNumber As Variant) As clsEnhancedNumber
    If Not IsNumeric(pNumber) Then
        Err.Raise vbObjectError + 1000, "clsEnhancedNumber.Modulo", "pNumber is not a number"
    End If

    If pNumber = 0 Then
        Err.Raise vbObjectError + 1005, "clsEnhancedNumber.Modulo", "pNumber is zero"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = mValue - pNumber * VBA.Int(mValue / pNumber)

    Set Modulo = lvResult
End Function

'==============================================
' /**
'  * Floor メソッド
'  * 任意の丸め単位で内部の数値を切り捨て、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Double} [pRoundingUnit=1] - 丸め単位（デフォルトは1）
'  * @throws {Error} 丸め単位が0の場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Floor(Optional ByVal pRoundingUnit As Double = 1) As clsEnhancedNumber
    If pRoundingUnit = 0 Then
        Err.Raise vbObjectError + 1010, "clsEnhancedNumber", "Rounding unit cannot be zero"
    End If

    Dim lvResult As clsEnhancedNumber

    Set lvResult = GetInstance()
    lvResult.Value = pRoundingUnit * VBA.Int(mValue / pRoundingUnit)

    Set Floor = lvResult
End Function

'==============================================
' /**
'  * Ceiling メソッド
'  * 任意の丸め単位で内部の数値を切り上げ、新しい clsEnhancedNumber オブジェクトを返す
'  *
'  * @param {Double} [pRoundingUnit=1] - 丸め単位（デフォルトは1）
'  * @throws {Error} 丸め単位が0の場合、エラーを発生させる
'  * @return {clsEnhancedNumber} 結果のオブジェクト
'  */
'==============================================
Public Function Ceiling(Optional ByVal pRoundingUnit As Double = 1) As clsEnhancedNumber
    If pRoundingUnit = 0 Then
        Err.Raise vbObjectError + 1011, "clsEnhancedNumber", "Rounding unit cannot be zero"
    End If

    Dim lvResult As clsEnhancedNumber
    Dim lvTemp As Double

    Set lvResult = GetInstance()
    lvTemp = mValue / pRoundingUnit

    If lvTemp = VBA.Int(lvTemp) Then
        lvResult.Value = mValue
    Else
        lvResult.Value = pRoundingUnit * (VBA.Int(lvTemp) + 1)
    End If

    Set Ceiling = lvResult
End Function

'==============================================
' /**
'  * pInPlaceUpdate の値に応じたインスタンスを返す
'  *
'  * @private
'  * @return {clsEnhancedNumber} インプレース更新の場合は現在のインスタンス、そうでなければ新規インスタンス
'  */
'==============================================
Private Function GetInstance() As clsEnhancedNumber
    If mInPlaceUpdate Then
        ' インプレース更新の場合、現在のインスタンスを返す
        Set GetInstance = Me
    Else
        ' それ以外は新たなインスタンスを生成して返す
        Set GetInstance = New clsEnhancedNumber
    End If
End Function
