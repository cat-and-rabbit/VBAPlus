Attribute VB_Name = "modFactory"
Option Explicit

'==============================================
' /**
'  * CreateEnhancedString 関数
'  * 新しい clsEnhancedString インスタンスを生成し、初期値とインプレース更新フラグを設定する
'  *
'  * @param {String} [pInitialValue=""] - 初期の文字列値（省略可能、デフォルトは空文字列）
'  * @param {Boolean} [pInPlaceUpdate=false] - インプレース更新フラグ（省略可能、デフォルトは False）
'  * @return {clsEnhancedString} 初期化された clsEnhancedString インスタンス
'  */
'==============================================
Public Function CreateEnhancedString(Optional pInitialValue As String = "", Optional ByVal pInPlaceUpdate As Boolean = False) As clsEnhancedString
    Dim lvClass As clsEnhancedString
    
    ' 新しい clsEnhancedString インスタンスを生成
    Set lvClass = New clsEnhancedString
    
    ' 初期値とインプレース更新フラグを設定
    lvClass.Initialize pInitialValue, pInPlaceUpdate
    
    ' 初期化されたインスタンスを返す
    Set CreateEnhancedString = lvClass
End Function
