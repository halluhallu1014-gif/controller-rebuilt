# コントローラー2の会社導入

[Raw貼り付け用コードを開く](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/main/workplace/controller2/InstallController2.txt)

Windows版のデスクトップExcel用です。会社側で行うVBAの貼り付けは、上記のRaw本文を標準モジュールに1回です。アドイン、`.frm`、`.frx`、PowerShell、ZIPの持込みやVBEのインポートは不要です。

## 導入前に確認すること

- 会社で許可されたマクロ実行方法と、VBAプロジェクト オブジェクト モデルへのプログラムによるアクセスが利用できることを管理者に確認します。後者が禁止されている場合は、この1回貼り付け方式を実行できません。制限を回避せず、管理者に承認済みの導入手段を確認してください。
- VBA、Excel、OLE Automation（stdole）、Microsoft Office、Microsoft Forms 2.0の標準ライブラリを使用します。UserForm作成時にFormsの参照が自動で入ります。Scripting.DictionaryとVBScript.RegExpは遅延バインディングです。COMコンポーネントが組織ポリシーで無効な場合は、IT管理者と利用可否を確認してください。
- 会社の管理簿・検証用管理簿のURL、シート名、見出し行、データ開始行を用意します。導入の動作試験には複製と専用の保存先を使用します。

## 新しいブックから作る

1. Excelで空のブックを開きます。シートが1枚の場合は **Shift + F11** を2回押し、空のワークシート3枚にします。ブックをExcelマクロ有効ブック（`.xlsm`）として保存します。
2. Rawをブラウザーで開いて本文全体をコピーします。Excelで **Alt + F11** を押し、対象ブックのVBEで **挿入 → 標準モジュール** を選びます。**F4** を押し、プロパティの **(Name)** を `Controller2Setup` に変更して本文を貼り付けます。
3. Excelの **Alt + F8** で `InstallController2` を実行します。ステータスバーに **Controller2 install complete** と表示されたことを確認します。失敗が表示された場合はその内容を控え、組み立て途中のブックを保存せずに閉じます。
4. VBEで `Controller2Setup` を右クリックして削除します。エクスポートの確認では「いいえ」を選択します。**デバッグ → VBAProjectのコンパイル** を実行します。コンポーネントが18件であることと、`Address`、`Fields`、`Operations` の3シートがあることを確認します。
5. 下記の設定を記入して保存します。まず試験用の管理簿で動作を確認してください。開き直すと日次バックアップを処理し、フォームがモードレスで開きます。フォームを開き直すには **Alt + F8 → OpenLedgerController** を使います。

## 設定

| シート | 入力項目 |
|---|---|
| Address A、E、F | 識別名、役割（Source／Master／Branch）、対象の署名。Masterは1件とし、Branchは利用する署ごとに設定します。 |
| Address B、C、D | Bはフォルダの直接HTTPS URL、Cは拡張子付き実ファイル名、Dは実シート名。週次取込元のCは空欄で、実行時に選択します。行を追加する場合は役割を含めて設定してください。 |
| Address G、H | 見出し行・データ開始行。初期値は試験用ファイルに合わせた1・2です。会社ファイルの実際の行番号を指定してください。 |
| Fields A–M | 管理簿列・見出し・型・編集可否・照合キー順・同期可否・週次取込元の見出し／列位置／必須指定。照合キーは初期値で共通番号、枝番号の順です。同期可否の未確定39項目は空欄です。業務承認後に○／×を指定してください。 |
| Operations B2–B5 | B2は操作者ID。B3は操作ログ用の専用.xlsx、B4は日次バックアップ用の専用.xlsxの直接HTTPS URL。B5は保持日数で初期値5。B3とB4は書込み前に専用の空ブックとして用意します。 |
| Operations B7–B8 | B7は処理未済判定に使う完了日項目のFieldsのID。B8は新規復旧ジャーナルの保存フォルダURL。B8が空欄なら、画面で用意済みの空ジャーナルを選択します。 |

初期状態に会社パス、署名、保存先、操作者を含めていません。転記の直前には既存エンジンが要求する排他運用を会社側で確保してください。画面の確認ボタンは排他ロックを取得する処理ではありません。

## SharePoint確認事項

- エンジンはWindows版Excelの認証を使い、管理簿の直接HTTPS URLを開きます。共有リンクではなくドキュメントライブラリ上の直接URLを使います。
- 操作ログと日次バックアップはHTTPS指定時にExcelのチェックアウト・チェックインを使用します。管理簿は排他運用を事前に確保し、既存エンジンが書込み可能と判定できることを確認します。対象のライブラリで、チェックアウト、チェックイン、Excelデスクトップでの直接保存、ジャーナル新規作成が許可されている必要があります。
- 会社テナントでの動作は未確認です。本番前に専用の試験ライブラリで認証、読込み、排他、保存・再読込み、ジャーナル復旧、操作ログ、バックアップを順に確認してください。

## コンポーネント

| 種類 | 名前 |
|---|---|
| ブック／シート（4） | ThisWorkbook、3枚のシートの文書モジュール |
| 標準モジュール（8） | LedgerTransfer、LedgerTransferFlows、OperationAudit、DailyBackup、ControllerData、ControllerActions、modCalendar、modUITheme |
| クラス（4） | ExtendedListBox、ExtendedListBoxLabel、CCalendarDayButton、ControllerDateButton |
| UserForm（2） | LedgerControllerForm、frmCalendar |

[個別コンポーネントのコード](components/) は内容照合用です。フォームのコードだけを貼り付けても画面は作成されません。画面のコントロールとプロパティは、Raw本文内の `DesignLedgerControllerForm` と `DesignfrmCalendar` が作ります。


## 検証結果

Windows版Excel 16.0の新規 .xlsm に配布ファイルの本文を挿入し、18件のコンポーネントを再構築しました。全体コンパイル、元ソースとの照合、保存後の別Excelでの再読込み・参照検証は合格しました。

複製環境では検索、詳細、カレンダー、個別登録と再読込み、先頭ゼロ、操作ログ、日次バックアップと同日重複防止が合格しました。会社のSharePointライブラリとITポリシーは未検証です。
