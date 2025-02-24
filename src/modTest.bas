Attribute VB_Name = "modTest"
Option Explicit

'==============================================
' /**
'  * Test_All
'  *
'  * 全テストを実行する
'  *
'  * @function Test_All
'  */
'==============================================
Public Sub Test_All()
    Call Test_clsEnhancedString
    Call Test_clsEnhancedNumber
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString
'  *
'  * clsEnhancedString クラスの全テストを実行する
'  *
'  * @function Test_clsEnhancedString
'  */
'==============================================
Public Sub Test_clsEnhancedString()
    Call Test_clsEnhancedString_Initialize
    Call Test_clsEnhancedString_Value
    Call Test_clsEnhancedString_Length
    Call Test_clsEnhancedString_ToUpperCase
    Call Test_clsEnhancedString_ToLowerCase
    Call Test_clsEnhancedString_Trim
    Call Test_clsEnhancedString_TrimStart
    Call Test_clsEnhancedString_TrimEnd
    Call Test_clsEnhancedString_Slice
    Call Test_clsEnhancedString_Splice
    Call Test_clsEnhancedString_Includes
    Call Test_clsEnhancedString_IndexOf
    Call Test_clsEnhancedString_StartsWith
    Call Test_clsEnhancedString_EndsWith
    Call Test_clsEnhancedString_Replace
    Call Test_clsEnhancedString_ReplaceAll
    Call Test_clsEnhancedString_Split
    Call Test_clsEnhancedString_PadStart
    Call Test_clsEnhancedString_PadEnd
    Call Test_clsEnhancedString_Repeat
    Call Test_clsEnhancedString_Template
    Call Test_clsEnhancedString_Reverse
    Call Test_clsEnhancedString_Test
    Call Test_clsEnhancedString_ReplaceRegex
    Call Test_clsEnhancedString_Match
    Call Test_clsEnhancedString_InPlaceUpdate
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Initialize
'  *
'  * 初期化処理のテスト。
'  * インスタンス生成時に Value プロパティが空文字であること、及び
'  * Initialize メソッドで値が設定されることを確認する。
'  *
'  * @function Test_clsEnhancedString_Initialize
'  */
'==============================================
Private Sub Test_clsEnhancedString_Initialize()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    Debug.Assert lvStr.Value = ""
    
    lvStr.Initialize "Test"
    Debug.Assert lvStr.Value = "Test"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Value
'  *
'  * Value プロパティの設定と取得のテスト。
'  *
'  * @function Test_clsEnhancedString_Value
'  */
'==============================================
Private Sub Test_clsEnhancedString_Value()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Length
'  *
'  * Length プロパティが文字列の長さを正しく返すかテストする。
'  *
'  * @function Test_clsEnhancedString_Length
'  */
'==============================================
Private Sub Test_clsEnhancedString_Length()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Length = 5
    
    lvStr.Value = ""
    Debug.Assert lvStr.Length = 0
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_ToUpperCase
'  *
'  * ToUpperCase メソッドが文字列を大文字に変換するかテストする。
'  *
'  * @function Test_clsEnhancedString_ToUpperCase
'  */
'==============================================
Private Sub Test_clsEnhancedString_ToUpperCase()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "abc"
    Debug.Assert lvStr.ToUpperCase.Value = "ABC"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_ToLowerCase
'  *
'  * ToLowerCase メソッドが文字列を小文字に変換するかテストする。
'  *
'  * @function Test_clsEnhancedString_ToLowerCase
'  */
'==============================================
Private Sub Test_clsEnhancedString_ToLowerCase()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "ABC"
    Debug.Assert lvStr.ToLowerCase.Value = "abc"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Trim
'  *
'  * Trim メソッドが前後の空白を正しく除去するかテストする。
'  *
'  * @function Test_clsEnhancedString_Trim
'  */
'==============================================
Private Sub Test_clsEnhancedString_Trim()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "  Hello  "
    Debug.Assert lvStr.Trim.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_TrimStart
'  *
'  * TrimStart メソッドが先頭の空白を正しく除去するかテストする。
'  *
'  * @function Test_clsEnhancedString_TrimStart
'  */
'==============================================
Private Sub Test_clsEnhancedString_TrimStart()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "  Hello"
    Debug.Assert lvStr.TrimStart.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_TrimEnd
'  *
'  * TrimEnd メソッドが末尾の空白を正しく除去するかテストする。
'  *
'  * @function Test_clsEnhancedString_TrimEnd
'  */
'==============================================
Private Sub Test_clsEnhancedString_TrimEnd()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello  "
    Debug.Assert lvStr.TrimEnd.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Slice
'  *
'  * Slice メソッドが指定された範囲の文字列を正しく抽出するかテストする。
'  *
'  * @function Test_clsEnhancedString_Slice
'  */
'==============================================
Private Sub Test_clsEnhancedString_Slice()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Slice(1, 4).Value = "ell"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Splice
'  *
'  * Splice メソッドが指定範囲の文字列を正しく置換または削除するかテストする。
'  *
'  * @function Test_clsEnhancedString_Splice
'  */
'==============================================
Private Sub Test_clsEnhancedString_Splice()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Splice(1, 4, "XX").Value = "HXXo"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Includes
'  *
'  * Includes メソッドが指定文字列の存在を正しく判定するかテストする。
'  *
'  * @function Test_clsEnhancedString_Includes
'  */
'==============================================
Private Sub Test_clsEnhancedString_Includes()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Includes("ll") = True
    Debug.Assert lvStr.Includes("XX") = False
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_IndexOf
'  *
'  * IndexOf メソッドが指定文字列の位置（0オリジン）を正しく返すかテストする。
'  *
'  * @function Test_clsEnhancedString_IndexOf
'  */
'==============================================
Private Sub Test_clsEnhancedString_IndexOf()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.IndexOf("ll") = 2
    Debug.Assert lvStr.IndexOf("XX") = -1
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_StartsWith
'  *
'  * StartsWith メソッドが文字列の先頭一致を正しく判定するかテストする。
'  *
'  * @function Test_clsEnhancedString_StartsWith
'  */
'==============================================
Private Sub Test_clsEnhancedString_StartsWith()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.StartsWith("He") = True
    Debug.Assert lvStr.StartsWith("XX") = False
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_EndsWith
'  *
'  * EndsWith メソッドが文字列の末尾一致を正しく判定するかテストする。
'  *
'  * @function Test_clsEnhancedString_EndsWith
'  */
'==============================================
Private Sub Test_clsEnhancedString_EndsWith()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.EndsWith("lo") = True
    Debug.Assert lvStr.EndsWith("XX") = False
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Replace
'  *
'  * Replace メソッドが文字列内の最初の一致を正しく置換するかテストする。
'  *
'  * @function Test_clsEnhancedString_Replace
'  */
'==============================================
Private Sub Test_clsEnhancedString_Replace()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Replace("ll", "XX").Value = "HeXXo"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_ReplaceAll
'  *
'  * ReplaceAll メソッドが文字列内の全一致部分を正しく置換するかテストする。
'  *
'  * @function Test_clsEnhancedString_ReplaceAll
'  */
'==============================================
Private Sub Test_clsEnhancedString_ReplaceAll()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello Hello"
    Debug.Assert lvStr.ReplaceAll("Hello", "Hi").Value = "Hi Hi"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Split
'  *
'  * Split メソッドが指定デリミタで文字列を分割できるかテストする。
'  *
'  * @function Test_clsEnhancedString_Split
'  */
'==============================================
Private Sub Test_clsEnhancedString_Split()
    Dim lvStr As clsEnhancedString
    Dim lvResult As Variant
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello,World,Test"
    lvResult = lvStr.Split(",")
    
    Debug.Assert lvResult(0) = "Hello"
    Debug.Assert lvResult(1) = "World"
    Debug.Assert lvResult(2) = "Test"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_PadStart
'  *
'  * PadStart メソッドが文字列の先頭に指定文字でパディングできるかテストする。
'  *
'  * @function Test_clsEnhancedString_PadStart
'  */
'==============================================
Private Sub Test_clsEnhancedString_PadStart()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.PadStart(10, "*").Value = "*****Hello"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_PadEnd
'  *
'  * PadEnd メソッドが文字列の末尾に指定文字でパディングできるかテストする。
'  *
'  * @function Test_clsEnhancedString_PadEnd
'  */
'==============================================
Private Sub Test_clsEnhancedString_PadEnd()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.PadEnd(10, "*").Value = "Hello*****"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Repeat
'  *
'  * Repeat メソッドが文字列を指定回数繰り返すかテストする。
'  *
'  * @function Test_clsEnhancedString_Repeat
'  */
'==============================================
Private Sub Test_clsEnhancedString_Repeat()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "A"
    Debug.Assert lvStr.Repeat(5).Value = "AAAAA"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Template
'  *
'  * Template メソッドがテンプレート内のプレースホルダーを正しく置換するかテストする。
'  *
'  * @function Test_clsEnhancedString_Template
'  */
'==============================================
Private Sub Test_clsEnhancedString_Template()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello {0}, welcome to {1}!"
    Debug.Assert lvStr.Template("John", "VBA").Value = "Hello John, welcome to VBA!"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Reverse
'  *
'  * Reverse メソッドが文字列を逆順に変換するかテストする。
'  *
'  * @function Test_clsEnhancedString_Reverse
'  */
'==============================================
Private Sub Test_clsEnhancedString_Reverse()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Reverse.Value = "olleH"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Test
'  *
'  * Test メソッドが正規表現で文字列をテストできるか確認する。
'  *
'  * @function Test_clsEnhancedString_Test
'  */
'==============================================
Private Sub Test_clsEnhancedString_Test()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello123"
    Debug.Assert lvStr.Test("\d+") = True
    Debug.Assert lvStr.Test("^\D+$") = False
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_ReplaceRegex
'  *
'  * ReplaceRegex メソッドが正規表現による置換を正しく実施するかテストする。
'  *
'  * @function Test_clsEnhancedString_ReplaceRegex
'  */
'==============================================
Private Sub Test_clsEnhancedString_ReplaceRegex()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "123-456-789"
    Debug.Assert lvStr.ReplaceRegex("\d{3}", "XXX").Value = "XXX-XXX-XXX"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_Match
'  *
'  * Match メソッドが正規表現でマッチした部分を正しく取得するかテストする。
'  *
'  * @function Test_clsEnhancedString_Match
'  */
'==============================================
Private Sub Test_clsEnhancedString_Match()
    Dim lvStr As clsEnhancedString
    Dim lvMatches As Object
    Dim lvMatch As Object
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "abc 123 def 456"
    Set lvMatches = lvStr.Match("\d+")
    
    Debug.Assert lvMatches.Count = 2
    Debug.Assert lvMatches.Item(0) = "123"
    Debug.Assert lvMatches.Item(1) = "456"
    
    Set lvStr = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedString_InPlaceUpdate
'  *
'  * mInPlaceUpdate フラグが正しく動作するかテストする。
'  * インプレース更新が False の場合は元のインスタンスは更新されず、
'  * True の場合は元のインスタンスが更新されることを確認する。
'  *
'  * @function Test_clsEnhancedString_InPlaceUpdate
'  */
'==============================================
Private Sub Test_clsEnhancedString_InPlaceUpdate()
    Dim lvStr As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' インプレース更新が False の場合（デフォルト）
    Set lvStr = New clsEnhancedString
    lvStr.Initialize "Hello"
    Set lvResult = lvStr.ToUpperCase
    Debug.Assert lvStr.Value = "Hello"
    Debug.Assert lvResult.Value = "HELLO"
    
    ' インプレース更新が True の場合
    Set lvStr = New clsEnhancedString
    lvStr.Initialize "Hello", True
    Set lvResult = lvStr.ToUpperCase
    Debug.Assert lvStr.Value = "HELLO"
    Debug.Assert lvResult.Value = "HELLO"
    
    Set lvStr = Nothing
    Set lvResult = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber
'  *
'  * clsEnhancedNumber クラスの全テストを実行する
'  *
'  * @function Test_clsEnhancedNumber
'  */
'==============================================
Public Sub Test_clsEnhancedNumber()
    Call Test_clsEnhancedNumber_Initialize
    Call Test_clsEnhancedNumber_Value
    Call Test_clsEnhancedNumber_Add
    Call Test_clsEnhancedNumber_Subtract
    Call Test_clsEnhancedNumber_Multiply
    Call Test_clsEnhancedNumber_Divide
    Call Test_clsEnhancedNumber_Pow
    Call Test_clsEnhancedNumber_Sqrt
    Call Test_clsEnhancedNumber_Round
    Call Test_clsEnhancedNumber_Absolute
    Call Test_clsEnhancedNumber_ToString
    Call Test_clsEnhancedNumber_Sin
    Call Test_clsEnhancedNumber_Cos
    Call Test_clsEnhancedNumber_Tan
    Call Test_clsEnhancedNumber_LogE
    Call Test_clsEnhancedNumber_Log10
    Call Test_clsEnhancedNumber_Exp
    Call Test_clsEnhancedNumber_Modulo
    Call Test_clsEnhancedNumber_Floor
    Call Test_clsEnhancedNumber_Ceiling
    Call Test_clsEnhancedNumber_InPlaceUpdate
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Initialize
'  *
'  * 初期化処理のテスト。インスタンス生成時の初期値（0）と
'  * Initialize メソッドで設定した値が反映されるか確認する。
'  *
'  * @function Test_clsEnhancedNumber_Initialize
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Initialize()
    Dim num As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    Debug.Assert num.Value = 0
    
    num.Initialize 42
    Debug.Assert num.Value = 42
    
    Set num = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Value
'  *
'  * Value プロパティの設定と取得のテスト。
'  *
'  * @function Test_clsEnhancedNumber_Value
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Value()
    Dim num As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    
    num.Value = 100
    Debug.Assert num.Value = 100
    
    Set num = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Add
'  *
'  * Add メソッドが加算を正しく行うかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Add
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Add()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 10
    
    Set result = num.Add(5)
    Debug.Assert result.Value = 15
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Subtract
'  *
'  * Subtract メソッドが減算を正しく行うかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Subtract
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Subtract()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 10
    
    Set result = num.Subtract(3)
    Debug.Assert result.Value = 7
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Multiply
'  *
'  * Multiply メソッドが乗算を正しく行うかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Multiply
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Multiply()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 5
    
    Set result = num.Multiply(3)
    Debug.Assert result.Value = 15
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Divide
'  *
'  * Divide メソッドが除算を正しく行うかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Divide
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Divide()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 20
    
    Set result = num.Divide(4)
    Debug.Assert result.Value = 5
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Pow
'  *
'  * Pow メソッドが累乗を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Pow
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Pow()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 2
    
    Set result = num.Pow(3)
    Debug.Assert result.Value = 8 ' 2^3 = 8
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Sqrt
'  *
'  * Sqrt メソッドが平方根を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Sqrt
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Sqrt()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 16
    
    Set result = num.Sqrt()
    Debug.Assert result.Value = 4
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Round
'  *
'  * Round メソッドが指定桁数で四捨五入を正しく行うかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Round
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Round()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 3.14159
    
    Set result = num.Round(2)
    Debug.Assert result.Value = 3.14
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Absolute
'  *
'  * Absolute メソッドが絶対値を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Absolute
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Absolute()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize -10
    
    Set result = num.Absolute()
    Debug.Assert result.Value = 10
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_ToString
'  *
'  * ToString メソッドが数値を文字列に変換できるかテストする。
'  *
'  * @function Test_clsEnhancedNumber_ToString
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_ToString()
    Dim num As clsEnhancedNumber
    Dim strObj As clsEnhancedString
    Set num = New clsEnhancedNumber
    num.Initialize 12345
    
    Set strObj = num.ToString()
    Debug.Assert strObj.Value = "12345"
    
    Set num = Nothing
    Set strObj = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Sin
'  *
'  * Sin メソッドがサイン値を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Sin
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Sin()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 0
    
    Set result = num.Sin()
    Debug.Assert Abs(result.Value) < 0.0001 ' sin(0) = 0
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Cos
'  *
'  * Cos メソッドがコサイン値を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Cos
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Cos()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 0
    
    Set result = num.Cos()
    Debug.Assert Abs(result.Value - 1) < 0.0001 ' cos(0) = 1
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Tan
'  *
'  * Tan メソッドがタンジェント値を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Tan
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Tan()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 0
    
    Set result = num.Tan()
    Debug.Assert Abs(result.Value) < 0.0001 ' tan(0) = 0
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_LogE
'  *
'  * LogE メソッドが自然対数を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_LogE
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_LogE()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 1
    
    Set result = num.LogE()
    Debug.Assert Abs(result.Value) < 0.0001 ' ln(1) = 0
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Log10
'  *
'  * Log10 メソッドが常用対数を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Log10
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Log10()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 100
    
    Set result = num.Log10()
    Debug.Assert Abs(result.Value - 2) < 0.0001 ' log10(100) = 2
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Exp
'  *
'  * Exp メソッドが指数関数を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Exp
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Exp()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 1
    
    Set result = num.Exp()
    Debug.Assert Abs(result.Value - 2.71828) < 0.001
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Modulo
'  *
'  * Modulo メソッドが剰余を正しく計算するかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Modulo
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Modulo()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 10
    
    Set result = num.Modulo(3)
    Debug.Assert result.Value = 10 - 3 * Int(10 / 3) ' 10 modulo 3 = 1
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Floor
'  *
'  * Floor メソッドが切り捨てを正しく行うかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Floor
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Floor()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 3.7
    
    Set result = num.Floor()
    Debug.Assert result.Value = 3
    
    ' 丸め単位を指定した場合（例: 0.5 単位）
    Set result = num.Floor(0.5)
    Debug.Assert result.Value = 3.5
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_Ceiling
'  *
'  * Ceiling メソッドが切り上げを正しく行うかテストする。
'  *
'  * @function Test_clsEnhancedNumber_Ceiling
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_Ceiling()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    Set num = New clsEnhancedNumber
    num.Initialize 3.2
    
    Set result = num.Ceiling()
    Debug.Assert result.Value = 4
    
    ' 丸め単位を指定した場合（例: 0.5 単位）
    Set result = num.Ceiling(0.5)
    Debug.Assert result.Value = 3.5
    
    Set num = Nothing
    Set result = Nothing
End Sub

'==============================================
' /**
'  * Test_clsEnhancedNumber_InPlaceUpdate
'  *
'  * mInPlaceUpdate フラグが正しく動作するかテストする。
'  * インプレース更新が False の場合は元のインスタンスは更新されず、
'  * True の場合は元のインスタンスが更新されることを確認する。
'  *
'  * @function Test_clsEnhancedNumber_InPlaceUpdate
'  */
'==============================================
Private Sub Test_clsEnhancedNumber_InPlaceUpdate()
    Dim num As clsEnhancedNumber, result As clsEnhancedNumber
    
    ' インプレース更新が False の場合（デフォルト）
    Set num = New clsEnhancedNumber
    num.Initialize 10, False
    Set result = num.Add(5)
    Debug.Assert num.Value = 10        ' 元のインスタンスは変更されない
    Debug.Assert result.Value = 15
    
    ' インプレース更新が True の場合
    Set num = New clsEnhancedNumber
    num.Initialize 10, True
    Set result = num.Add(5)
    Debug.Assert num.Value = 15        ' 元のインスタンスが更新される
    Debug.Assert result.Value = 15
    
    Set num = Nothing
    Set result = Nothing
End Sub
