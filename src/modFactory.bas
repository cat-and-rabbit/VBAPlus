Attribute VB_Name = "modFactory"
Option Explicit

' CreateEnhancedString �֐�
' �V���� clsEnhancedString �C���X�^���X�𐶐����A�����l�ƃC���v���[�X�X�V�t���O��ݒ肷��
' pInitialValue: �����̕�����l�i�ȗ��\�A�f�t�H���g�͋󕶎���j
' pInPlaceUpdate: �C���v���[�X�X�V�t���O�i�ȗ��\�A�f�t�H���g�� False�j
' �߂�l: ���������ꂽ clsEnhancedString �C���X�^���X
Public Function CreateEnhancedString(Optional pInitialValue As String = "", Optional ByVal pInPlaceUpdate As Boolean = False) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    ' �V���� clsEnhancedString �C���X�^���X�𐶐�
    Set lvClass = New clsEnhancedString
    
    ' �����l�ƃC���v���[�X�X�V�t���O��ݒ�
    lvClass.Initialize pInitialValue, pInPlaceUpdate
    
    ' ���������ꂽ�C���X�^���X��Ԃ�
    Set CreateEnhancedString = lvClass
End Function
