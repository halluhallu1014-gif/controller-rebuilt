# CONTROLLER SharePoint運用候補の導入と復旧

現行の正式版と分けた候補です。職場の管理簿に接続する前に、専用の検証フォルダと架空データで以下を確認します。ローカルでの合格範囲と職場で未実測の操作は末尾に記載しています。

## 空ブックからの構築

Excelデスクトップ版で本体用と設定ツール用に空ブックを1冊ずつ作ります。VBEで表の順にコード全体を貼ります。「デバッグ」からVBAプロジェクトをコンパイルし、Excelへ戻って Alt+F8 で指定のマクロを1回実行します。構築完了のダイアログを確認し、マクロ有効ブック（.xlsm）形式で保存します。

| 順序 | 本体の貼り付け先 | 貼り付けるファイル |
|---|---|---|
| 1 | 標準モジュール Core | [Core](main/Core.txt) |
| 2 | 標準モジュール Configuration | [Configuration](main/Configuration.txt) |
| 3 | 標準モジュール Presentation | [Presentation](main/Presentation.txt) |
| 4 | 標準モジュール Storage | [Storage](main/Storage.txt) |
| 5 | 標準モジュール CaseComparison | [CaseComparison](main/CaseComparison.txt) |
| 6 | 標準モジュール DeployController | [本体構築](main/DeployController.txt) |
| 7 | クラスモジュール CaseRecord | [CaseRecord](main/CaseRecord.txt) |
| 8 | 既存の ThisWorkbook | [ThisWorkbook](main/ThisWorkbook.txt) |

本体の構築マクロは `InstallControllerMain`。保存名は `CONTROLLER.xlsm`。構築後は「検索」「詳細」が表示されます。

| 順序 | 設定ツールの貼り付け先 | 貼り付けるファイル |
|---|---|---|
| 1 | 標準モジュール Module1 | [設定標準](setup/Module1.txt) |
| 2 | 空のユーザーフォーム UserForm1 | [SettingsForm](setup/UserForm1.txt) |
| 3 | 空のユーザーフォーム UserForm2 | [PickForm](setup/UserForm2.txt) |
| 4 | 空のユーザーフォーム UserForm3 | [ChoicesForm](setup/UserForm3.txt) |
| 5 | 既存の ThisWorkbook | [設定ThisWorkbook](setup/ThisWorkbook.txt) |

設定ツールの構築マクロは `InstallControllerSettings`。保存名は `CONTROLLER設定.xlsm`。Excelを閉じて再起動し、本体の検索画面と設定ツールの「設定を開く」を確認します。

## 管理簿と初期設定

1. 検証用の常勤・非常勤フォルダに、同じ構築済みの `CONTROLLER.xlsm` を1冊ずつ配置します。管理簿とバックアップのフォルダは同じ SharePoint サイト内で用意します。
2. 管理簿の初期位置は常勤側です。移行する場合は、常勤側に空の `管理簿.xlsm` を保存して閉じます。別の検証用 .xlsm の標準モジュールに [LedgerMigration.vba](migration/LedgerMigration.txt) 全体を貼り、`MigrateLedger` を実行します。旧ブックの直接URL、旧シート名、移行先の直接URL、西暦4桁の年度を1年度ずつ指定します。移行先の同名年度は上書きされません。旧「文書名」「税目」「収受日」などは、新項目と同義であることを確認してから移行結果に補記します。原本は照合用に保持します。
3. 設定ツールで「対象を選択」を押し、設定対象の `CONTROLLER.xlsm` の直接URLを入力します。「運用設定」から「共通アドレス・フォルダ」を開き、下表の値を登録します。
4. 「年度シート追加」で、西暦4桁のシートを選びます。年度とシート名は同じ値です。複数年度を使うときは年度ごとに登録します。更新する年度は「運用設定」の「選択年度の更新権限」で許可します。設定保存時は、本体を他のExcelで開いたままにしないでください。
5. 「設定を保存」を押し、本体で読めたことを示す完了メッセージを確認します。常勤・非常勤の両方で設定と年度をそろえます。設定完了後、本体だけを起動して検索します。
6. 架空データで登録・更新して再読込します。詳細の「発送文書登録」を押し、カレンダーで発送日をダブルクリックし、文書名とお問い合わせ番号を入力します。内容は詳細画面に入り、「登録・更新」で管理簿に保存されます。常勤側から朝・夕の操作を行い、バックアップ内容と移動先を確認します。非常勤側から移動操作が拒否されることも確認します。

下表の URL は例です。実際のサイトとライブラリ名に置き換えます。Excelで対象を開いたときの `Workbook.FullName` を確認できる場合は、その直接URLを使います。共有リンク（例 `/:x:/r/`、`?web=1`）は直接URLとして使えません。

| 設定 | 値の例 |
|---|---|
| SharePointSite | `https://example.sharepoint.com/sites/Team` |
| BackupFolder | `https://example.sharepoint.com/sites/Team/Shared%20Documents/Backup` |
| StaffFolder | `https://example.sharepoint.com/sites/Team/Shared%20Documents/Staff` |
| GeneralFolder | `https://example.sharepoint.com/sites/Team/Shared%20Documents/General` |
| StaffLedger | `https://example.sharepoint.com/sites/Team/Shared%20Documents/Staff/管理簿.xlsm` |
| GeneralLedger | `https://example.sharepoint.com/sites/Team/Shared%20Documents/General/管理簿.xlsm` |

## エラー別の復旧

| 表示・症状 | 確認と復旧 |
|---|---|
| 両方に管理簿がある、両方にない、または所在確認エラー | 更新と移動を止めます。常勤・非常勤のフォルダをSharePointで開き、版履歴・更新者・内容から正式データを確定します。必要に応じて管理者が重複側を別の退避名に移し、1冊に確定した後に本体を再読込します。所在不明のまま新規作成や再移動をしないでください。 |
| 保存結果が不明 | 保存操作を行った方の CONTROLLER を開き、そのブックでオーナー認証して「保存結果確認」または `ConfirmSaveResult` を実行します。保存予定の行と再読込した行が一致すれば確認済みになります。不一致の場合は記録が残ります。対象年度・共通番号・SharePointの版履歴を突き合わせ、確認前に再保存しないでください。 |
| バックアップ確認待ち | `ReleaseUnsentMove` は、移動要求が未送信で、元ファイルのIDとETagが同じ、移動先が空と確認できた場合だけ記録を解除します。退避候補は残ります。解除後、原因を解消して朝夕操作を行います。 |
| 移動結果が不明 | 常勤側で `ConfirmMoveResult` を実行します。元の消失、同じファイルIDでの移動先の存在、バックアップと内容の一致を確認します。確定しない場合は再送せず、SharePointの履歴と両フォルダを管理者が確認します。確認待ち記録はそのまま保持します。 |
| バックアップから内容を確認したい | 常勤側でオーナー認証し、`RestoreBackup` を実行します。退避元と、同じ退避フォルダ内の未使用の `復旧確認_001.xlsm` など の直接URLを指定します。作成後、内容一致の表示を確認します。正式な差し替えが必要なら版履歴と比較し、管理者が正式データを確定します。 |
| 設定保存結果が不明 | 画面に出た退避ファイルを記録し、本体の設定を読み直します。SharePoint版履歴と退避内容を比較し、保存された値を確定してから作業を再開します。 |

## 確認範囲

自宅のExcelでVBAコンパイル、空ブック構築と再オープン、2年度の読込、登録・更新と再取得、18桁の先頭ゼロ、状態と完了日、発送文字列、二重所在の安全停止、朝夕移動と退避内容の照合、移行、設定保存を確認しました。

職場の SharePoint では、Excelの直接URLによる読込・保存、設定退避、サーバー側のコピー・移動をまだ実測していません。移動・退避および不在確認用のRESTは現在のWindows認証で応答を得られる環境が必要です。OAuth専用のREST認証には未対応です。Excelの保存前後に照合しますが、共同編集中の再照合から保存までをサーバーで一体の処理として保証する仕組みは未実装です。発送カレンダーの画面操作と旧機能全体の回帰も未確認です。

