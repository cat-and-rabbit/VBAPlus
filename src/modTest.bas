Attribute VB_Name = "modTest"
Option Explicit

' clsEnhancedString のテストモジュール
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

' 初期化のテスト
Private Sub Test_clsEnhancedString_Initialize()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    Debug.Assert lvStr.Value = ""
    
    lvStr.Initialize "Test"
    Debug.Assert lvStr.Value = "Test"
    
    Set lvStr = Nothing
End Sub

' Value プロパティのテスト
Private Sub Test_clsEnhancedString_Value()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

' Length プロパティのテスト
Private Sub Test_clsEnhancedString_Length()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Length = 5
    
    lvStr.Value = ""
    Debug.Assert lvStr.Length = 0
    
    Set lvStr = Nothing
End Sub

' ToUpperCase メソッドのテスト
Private Sub Test_clsEnhancedString_ToUpperCase()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "abc"
    Debug.Assert lvStr.ToUpperCase.Value = "ABC"
    
    Set lvStr = Nothing
End Sub

' ToLowerCase メソッドのテスト
Private Sub Test_clsEnhancedString_ToLowerCase()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "ABC"
    Debug.Assert lvStr.ToLowerCase.Value = "abc"
    
    Set lvStr = Nothing
End Sub

' Trim メソッドのテスト
Private Sub Test_clsEnhancedString_Trim()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "  Hello  "
    Debug.Assert lvStr.Trim.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

' TrimStart メソッドのテスト
Private Sub Test_clsEnhancedString_TrimStart()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "  Hello"
    Debug.Assert lvStr.TrimStart.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

' TrimEnd メソッドのテスト
Private Sub Test_clsEnhancedString_TrimEnd()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello  "
    Debug.Assert lvStr.TrimEnd.Value = "Hello"
    
    Set lvStr = Nothing
End Sub

' Slice メソッドのテスト
Private Sub Test_clsEnhancedString_Slice()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Slice(1, 4).Value = "ell"
    
    Set lvStr = Nothing
End Sub

' Splice メソッドのテスト
Private Sub Test_clsEnhancedString_Splice()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Splice(1, 4, "XX").Value = "HXXo"
    
    Set lvStr = Nothing
End Sub

' Includes メソッドのテスト
Private Sub Test_clsEnhancedString_Includes()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Includes("ll") = True
    Debug.Assert lvStr.Includes("XX") = False
    
    Set lvStr = Nothing
End Sub

' IndexOf メソッドのテスト
Private Sub Test_clsEnhancedString_IndexOf()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.IndexOf("ll") = 2
    Debug.Assert lvStr.IndexOf("XX") = -1
    
    Set lvStr = Nothing
End Sub

' StartsWith メソッドのテスト
Private Sub Test_clsEnhancedString_StartsWith()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.StartsWith("He") = True
    Debug.Assert lvStr.StartsWith("XX") = False
    
    Set lvStr = Nothing
End Sub

' EndsWith メソッドのテスト
Private Sub Test_clsEnhancedString_EndsWith()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.EndsWith("lo") = True
    Debug.Assert lvStr.EndsWith("XX") = False
    
    Set lvStr = Nothing
End Sub

' Replace メソッドのテスト
Private Sub Test_clsEnhancedString_Replace()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Replace("ll", "XX").Value = "HeXXo"
    
    Set lvStr = Nothing
End Sub

' ReplaceAll メソッドのテスト
Private Sub Test_clsEnhancedString_ReplaceAll()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello Hello"
    Debug.Assert lvStr.ReplaceAll("Hello", "Hi").Value = "Hi Hi"
    
    Set lvStr = Nothing
End Sub

' Split メソッドのテスト
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

' PadStart メソッドのテスト
Private Sub Test_clsEnhancedString_PadStart()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.PadStart(10, "*").Value = "*****Hello"
    
    Set lvStr = Nothing
End Sub

' PadEnd メソッドのテスト
Private Sub Test_clsEnhancedString_PadEnd()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.PadEnd(10, "*").Value = "Hello*****"
    
    Set lvStr = Nothing
End Sub

' Repeat メソッドのテスト
Private Sub Test_clsEnhancedString_Repeat()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "A"
    Debug.Assert lvStr.Repeat(5).Value = "AAAAA"
    
    Set lvStr = Nothing
End Sub

' Template メソッドのテスト
Private Sub Test_clsEnhancedString_Template()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello {0}, welcome to {1}!"
    Debug.Assert lvStr.Template("John", "VBA").Value = "Hello John, welcome to VBA!"
    
    Set lvStr = Nothing
End Sub

' Reverse メソッドのテスト
Private Sub Test_clsEnhancedString_Reverse()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello"
    Debug.Assert lvStr.Reverse.Value = "olleH"
    
    Set lvStr = Nothing
End Sub

' Test メソッドのテスト（正規表現）
Private Sub Test_clsEnhancedString_Test()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "Hello123"
    Debug.Assert lvStr.Test("\d+") = True
    Debug.Assert lvStr.Test("^\D+$") = False
    
    Set lvStr = Nothing
End Sub

' ReplaceRegex メソッドのテスト
Private Sub Test_clsEnhancedString_ReplaceRegex()
    Dim lvStr As clsEnhancedString
    Set lvStr = New clsEnhancedString
    
    lvStr.Value = "123-456-789"
    Debug.Assert lvStr.ReplaceRegex("\d{3}", "XXX").Value = "XXX-XXX-XXX"
    
    Set lvStr = Nothing
End Sub

' Match メソッドのテスト（正規表現）
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

' mInPlaceUpdate のテスト
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
