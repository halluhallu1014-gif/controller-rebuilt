# 新CONTROLLER
作成日：2026-11-04

**2026-11-05：以前の青・白・グレーの画面を基準にUIを改修しました。実ExcelでVBAコンパイル、画面生成、人物／関連事案のページ操作、結合入力欄、日付操作、設定・入力・追加項目の保持を確認しました。正式管理簿の登録・更新・SharePoint・移動・ZIPの実機試験は未完了です。**

既存ブックへの反映は[UI_UPDATE.md](UI_UPDATE.md)をご覧ください。差し替えはUi・Configuration・Utilitiesの3本、実行はApplyUiLayoutです。SetupControllerの再実行は不要です。

[検索画面](previews/検索.png) / [詳細画面](previews/詳細.png) / [取込画面](previews/取込.png) は、架空データを表示した実Excelから出力したプレビューです。

## 内容
- src：標準7モジュール＋CaseRecordクラス。すべてコピー用のUTF-8テキストです。
- [INSTALL.md](INSTALL.md)：初期構築、会社PCへの導入、差し替え、操作方法。
- [CONFIGURATION.md](CONFIGURATION.md)：Settings・Fields・画面配置・架空データ例。
- [DESIGN.md](DESIGN.md)：構造、保存、Masked Number、旧版調査、変更箇所。
- [TEST_RESULTS.md](TEST_RESULTS.md)：実施結果と実Excelで残っている受入確認。

旧CONTROLLER、既存管理簿、参照GitHubリポジトリは変更していません。このリポジトリには新CONTROLLERの配布コードと説明書だけを登録します。職場のURL・認証情報・実データは含みません。

## 会社PCでコードをコピーする

初めて導入する場合は、下の全8コードを新しいCONTROLLER.xlsmへコピーします。**旧CONTROLLER 3.xへUIの3本だけを貼る手順ではありません。** すでにこの再構築版を導入済みでUIだけ更新するときは[UI_UPDATE.md](UI_UPDATE.md)を使用します。

各Rawリンクを開き、全文をコピーして、表の種類・名前のモジュールへ貼り付けてください。.vbaは全文コピー用テキストです。シートモジュールやThisWorkbookへの貼付は不要です。

|名前|貼付先|コード|全文コピー用Raw|
|---|---|---|---|
|Core|標準モジュール|[開く](src/Core.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/Core.vba)|
|Configuration|標準モジュール|[開く](src/Configuration.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/Configuration.vba)|
|CaseService|標準モジュール|[開く](src/CaseService.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/CaseService.vba)|
|Storage|標準モジュール|[開く](src/Storage.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/Storage.vba)|
|Ui|標準モジュール|[開く](src/Ui.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/Ui.vba)|
|Utilities|標準モジュール|[開く](src/Utilities.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/Utilities.vba)|
|Operations|標準モジュール|[開く](src/Operations.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/Operations.vba)|
|CaseRecord|クラスモジュール|[開く](src/CaseRecord.vba)|[Raw](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/src/CaseRecord.vba)|

初期導入は [INSTALL.md](INSTALL.md)、年度別アドレス・設定セルは [CONFIGURATION.md](CONFIGURATION.md)、職場で確認する項目は [TEST_RESULTS.md](TEST_RESULTS.md) に記載しています。職場のURL・認証情報・個人情報はGitHubへ登録しないでください。

リポジトリ全体のダウンロードはGitHubの「Code → Download ZIP」を使用できます。会社でファイルを持ち込めない場合は、上のRawページからコードをコピーしてください。

## 最初の操作
1. 空の別ブックをCONTROLLER.xlsmで保存します。
2. srcの7ファイルを同名の標準モジュールへ、CaseRecord.vbaだけを同名のクラスモジュールへ全文貼り付けます。
3. VBEの「デバッグ→VBAProjectのコンパイル」を実行します。
4. SetupControllerを1回実行します。SettingsとFieldsを設定し、StartControllerで読込を開始します。

初期状態は年度Enabled=No、AllowWrite=No、SharePointExclusive=Noです。保存先を未確認のまま保存しない設定です。業務開始・終了・バックアップ復元・移動結果管理はStaffだけが実行します。通常の登録・更新は、旧版と同様にStaffとGeneralの両方を扱い、年度AllowWriteで保存可否を制御します。

## 維持した運用
業務開始は常勤用→非常勤用、業務終了は非常勤用→常勤用へ、最新の有効年度の管理簿1ファイルだけを移動します。最新はSettingsでEnabled=Yesの年度のうち数値最大の年度です。過年度は移動せず、設定された所在から検索・許可された更新を続けます。開始時は移動前に、その最新年度ファイルの全シートの型付きCSVをZIP化し、展開・復元照合を行います。新しいバックアップと移動を確認した後だけ、直近5作成日より古い自システム形式ZIPを整理します。年度間移動は追加していません。

## 未確定事項
詳細レイアウトはFieldsとExcelの名前定義で変更できます。外部システムの実文字列は未確認です。取込は明示的なラベルとコロン／タブの形式に限定した暫定実装で、実サンプルによる規則確定が必要です。Paletteのカレンダー資産は確認できていないため再利用していません。カレンダーは新規のシート実装です。
