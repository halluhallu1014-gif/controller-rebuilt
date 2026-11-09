# CONTROLLER 職場導入

Windows版Excelで、本体と設定ツールをそれぞれ空ブックから作ります。下のRawリンクは、検証したソースのコミットに固定しています。各ページで **Ctrl+A → Ctrl+C** を押して全文をコピーします。

## 本体（貼付8回）

| 順 | 貼り付け先 | Raw |
|---:|---|---|
| 1 | 標準モジュール `Core` | [Core.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/Core.txt) |
| 2 | 標準モジュール `Configuration` | [Configuration.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/Configuration.txt) |
| 3 | 標準モジュール `Presentation` | [Presentation.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/Presentation.txt) |
| 4 | 標準モジュール `Storage` | [Storage.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/Storage.txt) |
| 5 | 標準モジュール `CaseComparison` | [CaseComparison.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/CaseComparison.txt) |
| 6 | 標準モジュール `DeployController` | [DeployController.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/DeployController.txt) |
| 7 | クラスモジュール `CaseRecord` | [CaseRecord.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/CaseRecord.txt) |
| 8 | 既存の `ThisWorkbook` | [ThisWorkbook.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/main/ThisWorkbook.txt) |

## 設定ツール（貼付5回）

| 順 | 貼り付け先 | Raw |
|---:|---|---|
| 1 | 標準モジュール `Module1` | [Module1.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/setup/Module1.txt) |
| 2 | 空のUserForm `UserForm1` | [UserForm1.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/setup/UserForm1.txt) |
| 3 | 空のUserForm `UserForm2` | [UserForm2.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/setup/UserForm2.txt) |
| 4 | 空のUserForm `UserForm3` | [UserForm3.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/setup/UserForm3.txt) |
| 5 | 既存の `ThisWorkbook` | [ThisWorkbook.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/setup/ThisWorkbook.txt) |

## 構築と確認

1. 空ブック2冊を、Excel マクロ有効ブック形式の `CONTROLLER.xlsm` と `CONTROLLER設定.xlsm` として保存します。
2. Alt+F11でVBAエディターを開きます。表の種類に従って［挿入］から部品を追加し、F4のプロパティで名前を表と同じにします。UserFormは空のままF7でコード画面を開きます。既存の `ThisWorkbook` はプロジェクト一覧から開きます。
3. 各コード画面に対応するRawの全文を貼り付けます。既に `Option Explicit` がある場合は、コード画面を全選択して置き換えます。各ブックで［デバッグ］→［VBAProjectのコンパイル］を実行します。
4. Alt+F8で、本体は `InstallControllerMain`、設定ツールは `InstallControllerSettings` を1回実行します。「構築と構成確認が完了しました」を確認し、Ctrl+Sで保存します。
5. Excelを終了し、起動し直して両方のブックを開きます。本体の検索画面と設定ツールのフォームが表示されることを確認します。エラー時は表示と直前の操作を記録して、その工程で停止します。

## 管理簿と運用の設定

- 設定ツールで本体を選び、年度・管理簿・対象シート・見出しを登録します。最初は架空管理簿で、完了日付きの登録・更新・再読込を試してください。
- 本体の［整理簿と照合］でその回のファイルとシートを選びます。設定ツールの照合設定で13桁番号と、年分・文書名・税目などの照合項目を対応付けます。対象事案が明確な行の訂正欄を入力し、［一括更新］で新管理簿へ保存します。複数事案にまたがる共通項目は1件に集約し、詳細で更新対象を確認します。
- ［抽出・印刷］では日付項目、条件、年度・署・状態、並べ替えを選びます。申請日は、設定ツールで収受日と別の登録済み日付項目を指定した場合に使えます。別ブックに結果が出るので、印刷プレビューで確認できます。
- 詳細の［複数項目を取込］または［選択欄へ取込］で、コピー元画面を指定します。入力欄への反映前に対象項目と既存値の置換件数を確認します。クリップボードが使えない場合は、別の空Excelブックの文字列形式のセルに貼り、範囲を指定します。設定ツールの選択肢・表示形式の設定画面から取込ルールを編集し、架空文字列でテストできます。
- オーナーを設定ツールで登録してから、本体で Alt+F8 → `OpenOwnerWorkspace` を実行します。現在のWindowsアカウントが登録された `DOMAIN\ユーザー名` に一致する場合に表示できます。常勤の管理には既存の常勤権限も必要です。正式ファイルのアクセス制御は既存のSharePoint権限で行います。

## 旧管理簿を移行する場合（別ブックに貼付1回）

[LedgerMigration.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/00ffc03582820aba7614cfe2a679d51fc7011cc2/workplace/migration/LedgerMigration.txt) の全文を、別の空マクロ有効ブックの標準モジュールに貼ります。コンパイル後、`MigrateLedger` を実行します。ローカルの旧 `.xlsx` と対象シートを選び、新しいファイル名で保存します。終了画面の移行件数と、元・新ファイルの番号と枝番を確認してください。

## 検証範囲

自宅Excelで、本体と設定ツールの空ブック構築・保存・再起動、架空管理簿での照合・競合を含む一括更新・完了日・抽出・取込ルール・オーナー権限を確認しました。今回の空ブック検証ではソース登録を開発PCで自動化しました。職場の操作ではVBEへ貼り付けるため、VBAProject自動操作の権限は使いません。職場のSharePointと既存のOffice標準COMの利用可否は、職場環境での確認が必要です。
