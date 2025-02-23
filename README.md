# VBAPlus

VBAPlus は、VBA の基本的な型の利便性向上を目的として開発され、特に文字列操作における作業を効率化するための機能を提供します。

## clsEnhancedString
`clsEnhancedString` は、VBA での文字列操作を拡張するためのカスタムクラスです。  
このクラスは、以下の機能を提供しています:

- **初期化およびプロパティ**
  - `Initialize`: 初期文字列を設定し、必要に応じてインプレース更新を有効化します。
  - `Value`: 文字列の取得および設定。
  - `Length`: 文字列の長さ（文字数）の取得。

- **変換機能**
  - `ToUpperCase`: 文字列を大文字に変換します。
  - `ToLowerCase`: 文字列を小文字に変換します。

- **トリミング機能**
  - `Trim`: 両端の空白を削除します。
  - `TrimStart`: 左側の空白を削除します。
  - `TrimEnd`: 右側の空白を削除します。

- **抽出・置換機能**
  - `Slice`: 文字列の一部を抽出します。
  - `Splice`: 文字列の一部を置換または削除、挿入します。

- **検索機能**
  - `Includes`: 指定した文字列が含まれているかチェックします。
  - `IndexOf`: 指定した文字列の位置（0ベース）を返します。
  - `StartsWith`: 文字列が指定した文字列で始まるかチェックします。
  - `EndsWith`: 文字列が指定した文字列で終わるかチェックします。

- **置換機能**
  - `Replace`: 最初に見つかった部分文字列を置換します。
  - `ReplaceAll`: 該当するすべての部分文字列を置換します。
  - `ReplaceRegex`: 正規表現に基づいて文字列を置換します。

- **その他の機能**
  - `Split`: 指定した区切り文字で文字列を分割します。
  - `PadStart`: 指定の長さになるよう左側を埋めます。
  - `PadEnd`: 指定の長さになるよう右側を埋めます。
  - `Repeat`: 文字列を指定した回数繰り返します。
  - `Template`: テンプレートとして文字列を適用します。
  - `Reverse`: 文字列を反転させます。
  - `Test`: 正規表現パターンにマッチするかテストします。
  - `Match`: 正規表現パターンに該当する部分を抽出します。
  - **InPlaceUpdate**: 文字列の更新方法に関する機能を提供します。

なお、VBAは本来1ベースのインデックスを持つ言語ですが、  
`clsEnhancedString` は0ベースの思想に従って設計されているため、  
インデックスや位置指定の動作が一般的な0開始となっています。

また、デフォルトでは新規インスタンスを返すオブジェクト指向の設計となっていますが、  
`Initialize` メソッドのオプションとしてインプレース更新（内部変更）を有効にすることで、  
既存インスタンス内で値の更新を行うことも可能です。