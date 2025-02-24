Attribute VB_Name = "modTest"
Option Explicit

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
