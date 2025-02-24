Attribute VB_Name = "modTest"
Option Explicit

'==============================================
' /**
'  * Test_clsEnhancedString
'  * 
'  * clsEnhancedString �N���X�̑S�e�X�g�����s����
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
'  * �����������̃e�X�g�B
'  * �C���X�^���X�������� Value �v���p�e�B���󕶎��ł��邱�ƁA�y��
'  * Initialize ���\�b�h�Œl���ݒ肳��邱�Ƃ��m�F����B
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
'  * Value �v���p�e�B�̐ݒ�Ǝ擾�̃e�X�g�B
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
'  * Length �v���p�e�B��������̒����𐳂����Ԃ����e�X�g����B
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
'  * ToUpperCase ���\�b�h���������啶���ɕϊ����邩�e�X�g����B
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
'  * ToLowerCase ���\�b�h����������������ɕϊ����邩�e�X�g����B
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
'  * Trim ���\�b�h���O��̋󔒂𐳂����������邩�e�X�g����B
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
'  * TrimStart ���\�b�h���擪�̋󔒂𐳂����������邩�e�X�g����B
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
'  * TrimEnd ���\�b�h�������̋󔒂𐳂����������邩�e�X�g����B
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
'  * Slice ���\�b�h���w�肳�ꂽ�͈͂̕�����𐳂������o���邩�e�X�g����B
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
'  * Splice ���\�b�h���w��͈͂̕�����𐳂����u���܂��͍폜���邩�e�X�g����B
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
'  * Includes ���\�b�h���w�蕶����̑��݂𐳂������肷�邩�e�X�g����B
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
'  * IndexOf ���\�b�h���w�蕶����̈ʒu�i0�I���W���j�𐳂����Ԃ����e�X�g����B
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
'  * StartsWith ���\�b�h��������̐擪��v�𐳂������肷�邩�e�X�g����B
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
'  * EndsWith ���\�b�h��������̖�����v�𐳂������肷�邩�e�X�g����B
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
'  * Replace ���\�b�h����������̍ŏ��̈�v�𐳂����u�����邩�e�X�g����B
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
'  * ReplaceAll ���\�b�h����������̑S��v�����𐳂����u�����邩�e�X�g����B
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
'  * Split ���\�b�h���w��f���~�^�ŕ�����𕪊��ł��邩�e�X�g����B
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
'  * PadStart ���\�b�h��������̐擪�Ɏw�蕶���Ńp�f�B���O�ł��邩�e�X�g����B
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
'  * PadEnd ���\�b�h��������̖����Ɏw�蕶���Ńp�f�B���O�ł��邩�e�X�g����B
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
'  * Repeat ���\�b�h����������w��񐔌J��Ԃ����e�X�g����B
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
'  * Template ���\�b�h���e���v���[�g���̃v���[�X�z���_�[�𐳂����u�����邩�e�X�g����B
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
'  * Reverse ���\�b�h����������t���ɕϊ����邩�e�X�g����B
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
'  * Test ���\�b�h�����K�\���ŕ�������e�X�g�ł��邩�m�F����B
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
'  * ReplaceRegex ���\�b�h�����K�\���ɂ��u���𐳂������{���邩�e�X�g����B
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
'  * Match ���\�b�h�����K�\���Ń}�b�`���������𐳂����擾���邩�e�X�g����B
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
'  * mInPlaceUpdate �t���O�����������삷�邩�e�X�g����B
'  * �C���v���[�X�X�V�� False �̏ꍇ�͌��̃C���X�^���X�͍X�V���ꂸ�A
'  * True �̏ꍇ�͌��̃C���X�^���X���X�V����邱�Ƃ��m�F����B
'  *
'  * @function Test_clsEnhancedString_InPlaceUpdate
'  */
'==============================================
Private Sub Test_clsEnhancedString_InPlaceUpdate()
    Dim lvStr As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' �C���v���[�X�X�V�� False �̏ꍇ�i�f�t�H���g�j
    Set lvStr = New clsEnhancedString
    lvStr.Initialize "Hello"
    Set lvResult = lvStr.ToUpperCase
    Debug.Assert lvStr.Value = "Hello"
    Debug.Assert lvResult.Value = "HELLO"
    
    ' �C���v���[�X�X�V�� True �̏ꍇ
    Set lvStr = New clsEnhancedString
    lvStr.Initialize "Hello", True
    Set lvResult = lvStr.ToUpperCase
    Debug.Assert lvStr.Value = "HELLO"
    Debug.Assert lvResult.Value = "HELLO"
    
    Set lvStr = Nothing
    Set lvResult = Nothing
End Sub
