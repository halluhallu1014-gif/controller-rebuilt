# CONTROLLER 職場導入用VBA

このページのRawリンクは、手動導入テストに使ったVBAソースとSHA-256が一致するコミットに固定しています。各Rawページで **Ctrl+A → Ctrl+C** を押して全文をコピーします。

職場では、Excelマクロと使用するWindows／Office標準COMの利用可否を管理担当に確認してから実施してください。このページは自宅で確認した手順の公開用です。職場のブラウザー・Excel・SharePointでの動作は未確認です。

## 1. 本体を作る

1. Excelで空白のブックを作ります。空シート1枚の状態で **F12** を押し、保存先のフォルダを開きます。ファイル名欄には新規ファイル名だけを入れ、種類を **Excel マクロ有効ブック (*.xlsm)** にして保存します。自宅試験の名前は **手動導入テスト_本体.xlsm** でした。既存の業務ブックに上書きしないでください。
2. **Alt+F11** でVBAエディターを開き、今保存したブックのプロジェクトを選びます。
3. 下表の順に貼り付けます。標準は **［挿入］→［標準モジュール］**、クラスは **［挿入］→［クラス モジュール］** を選びます。**F4** の **(オブジェクト名)** で表どおりに命名します。既存の **ThisWorkbook** は **Microsoft Excel Objects** の下から開きます。
4. 各Rawページで全文をコピーし、対象の **(コード)** 欄で **Ctrl+A → Ctrl+V** を押して全文を置き換えます。

| 順 | 貼り付け先の種類 | 貼り付け先 | Raw |
|---:|---|---|---|
| 1 | 標準 | `Core` | [Core.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/Core.txt) |
| 2 | 標準 | `Configuration` | [Configuration.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/Configuration.txt) |
| 3 | 標準 | `Presentation` | [Presentation.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/Presentation.txt) |
| 4 | 標準 | `Storage` | [Storage.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/Storage.txt) |
| 5 | 標準 | `CaseComparison` | [CaseComparison.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/CaseComparison.txt) |
| 6 | 標準 | `DeployController` | [DeployController.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/DeployController.txt) |
| 7 | クラス | `CaseRecord` | [CaseRecord.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/CaseRecord.txt) |
| 8 | 既存のブックモジュール | `ThisWorkbook` | [ThisWorkbook.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/main/ThisWorkbook.txt) |

5. **［デバッグ］→［VBAProject のコンパイル］** を選びます。エラーがなければ **Alt+F11** でExcelに戻り、**Alt+F8** から **InstallControllerMain** を一度だけ実行します。
6. **「CONTROLLERの構築と構成確認が完了しました。」** と表示されたら、メッセージを閉じて **Ctrl+S** で保存します。Excelを終了して起動し直し、同じブックを開きます。**検索** と **詳細** のタブ、およびVBEの **ThisWorkbook** に **Workbook_Open** が残っていることを確認します。
7. 本体を閉じます。設定ツールからの設定保存は、本体を閉じた状態で行います。

本体は新規部品7個、貼付8回、構築マクロ1本です。

## 2. 設定ツールを作る

1. 別の空白ブックを作ります。本体と同じ方法で、フォルダとファイル名を別々に指定し **.xlsm** で保存します。自宅試験の名前は **手動導入テスト_設定.xlsm** でした。
2. **Alt+F11** を押して設定ツール側のプロジェクトを選び、下表の順に貼り付けます。**Module1** は **［挿入］→［標準モジュール］** で作ります。**UserForm1〜3** は **［挿入］→［ユーザーフォーム］** で順に作り、名前は既定のままにします。各フォームを選び **F7** で **(コード)** 欄を開きます。フォーム上の部品は構築コードが用意します。
3. 各Rawページの全文をコピーし、対象の **(コード)** 欄で **Ctrl+A → Ctrl+V** を押します。最後の **ThisWorkbook** は設定ツール側の既存の項目です。

| 順 | 貼り付け先の種類 | 貼り付け先 | Raw |
|---:|---|---|---|
| 1 | 標準 | `Module1` | [Module1.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/setup/Module1.txt) |
| 2 | UserForm | `UserForm1` | [UserForm1.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/setup/UserForm1.txt) |
| 3 | UserForm | `UserForm2` | [UserForm2.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/setup/UserForm2.txt) |
| 4 | UserForm | `UserForm3` | [UserForm3.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/setup/UserForm3.txt) |
| 5 | 既存のブックモジュール | `ThisWorkbook` | [ThisWorkbook.txt](https://raw.githubusercontent.com/halluhallu1014-gif/controller-rebuilt/10211c5012de6be9f4cf3cddc834903a90cd212d/workplace/setup/ThisWorkbook.txt) |

4. **［デバッグ］→［VBAProject のコンパイル］** を選びます。エラーがなければ **Alt+F11** でExcelに戻り、**Alt+F8** から **InstallControllerSettings** を一度だけ実行します。
5. **「CONTROLLER設定の構築と構成確認が完了しました。」** と表示されたら、メッセージを閉じて **Ctrl+S** で保存します。Excelを終了して起動し直し、設定ツールを開きます。**「CONTROLLER ／ 設定」** フォームが起動し、**［対象を選択］** が見えれば構築確認は成功です。
6. **［対象を選択］** から今回構築した本体を選びます。自宅試験では **［管理簿を追加］** から架空管理簿のコピー2年度分を追加し、**［設定を保存］** で **「設定を保存し、本体で読み込めることを確認しました。」** と表示されるところまで確認しました。税目列を追加した2026年度のコピーへのファイル変更後も、同じ保存確認が出ました。

設定ツールは新規部品4個、貼付5回、構築マクロ1本です。途中でエラーが出たら操作を止め、表示内容と工程を記録します。構築マクロの再実行や途中のコード差し替えはせず、再試験は空白ブックから始めます。

## 自宅で確認した範囲と未検証項目

本体は構築、保存、Excel終了・再起動、再オープン後のVBA保持を確認しました。設定ツールは構築、保存、再起動後の設定フォーム起動、架空2年度の設定と税目列付きコピーへの切替後の設定保存を確認しました。

設定ツールの最初の保存で、保存ダイアログがファイル名欄のフルパスを受け付けない警告が出ました。その後同じ空白ブックの保存に成功し、手動で構築確認を続けました。この記録は警告なしで全工程を完了した試験や、全機能の検証完了を示すものではありません。

- 本体での登録・更新後の再読込
- 署ボタンでの絞り込み
- 選択肢設定の本体への反映
- 職場のSharePoint環境
- 職場のCOM利用可否
- 職場PCでのRaw閲覧とコピー、マクロ実行ポリシー

再構築した本体と設定ツールを組み合わせた過去の外部自動試験は、Excelとの接続切断で中断しました。自宅で手動確認できた工程と区別して扱います。
