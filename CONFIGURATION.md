# 設定例

## Settings
同じシートのA:Bにシステム設定、D:Iに年度表、J:Lに電話番号の判定表を置きます。行1は見出し、行2以降はデータです。表は列ごとに独立して増やせます。

|A Setting|B Value（初期例）|意味|
|---|---|---|
|ControllerRole|General|常勤=Staff、非常勤=General。管理操作はStaffだけ|
|SharePointSite|空欄|正式運用のサイト直接URL。末尾スラッシュなし|
|SharePointExclusive|No|必須チェックアウトを実環境で確認後のみYes|
|BackupFolder|空欄|バックアップ専用フォルダの絶対パス／同一サイトURL|
|BusinessState|空欄|開始・終了処理の結果表示用|
|Transport|Local|自宅はLocal、正式運用はSharePoint|

年度表の列は次のとおりです。保存先を列順以外の別シートへ重複登録しません。

|D Year|E StaffAddress|F Sheet|G Enabled|H AllowWrite|I GeneralAddress|
|---|---|---|---|---|---|
|2025|C:\ControllerDemo\Staff\\Ledger2025.xlsx|管理簿|Yes|No|C:\ControllerDemo\General\\Ledger2025.xlsx|
|2026|C:\ControllerDemo\Staff\\Ledger2026.xlsx|管理簿|Yes|No|C:\ControllerDemo\General\\Ledger2026.xlsx|

上記は存在しない架空パスです。自分で作ったテスト用の絶対パスへ置き換えてください。フォルダは先に作ります。両候補のどちらか片側にだけファイルを置きます。年度は文字列で保持し、表示にもそのまま使います。6年目・7年目もD:Iの行追加だけです。

業務開始・終了で移動するのは、Enabled=Yesの年度のうち数値最大の1年度だけです。例の2025・2026がともにYesなら2026だけを移動します。2027をYesで追加すると対象は2027へ切り替わります。行順・AllowWrite・PCの現在年では決めません。最新年度の所在や設定が不正なら、過年度へ代替せず移動を中止します。過年度の検索・更新用にEnabled=Yesを維持して構いません。

Settingsを編集したらStartControllerで読み直します。各年度の不正はその年度の読込エラーになり、正常年度の検索を妨げません。

電話判定表の例：

|J PhonePrefix|K AreaLength|L ExchangeLength|
|---|---|---|
|03|2|4|
|06|2|4|

Prefixは文字列です。最長一致で判定し、10桁固定電話のみ区切ります。全市外局番の一覧を同梱したものではありません。03/06以外で正しい市外局番が未登録の番号は数字のまま表示します。正しい地域別のプレフィックスと長さを確認できた場合だけ行を追加してください。単純な03等の市外局番だけで判定できない地域は、市内局番を含む識別可能なPrefixを設定します。根拠のない一律区切りをしません。070/080/090の11桁は3-4-4、050の11桁は独立ルールで3-4-4です。

## UI改修後の配置について
以下は2026-11-05改修後の標準配置です。実際に配置を変更した場合は、実ブックのFieldsを正として確認してください。Header・Type・Required・年度例外・年度設定値は保持します。

新しい名前定義はSearchName=検索!C3、SearchYear=検索!E3、SearchField=検索!C5、SearchValue=検索!E5、SearchResults=検索!B10:F21、SearchState=検索!B6、DetailYear=詳細!C5、DetailState=詳細!B2、NumberPaste=取込!C7、ImportPaste=取込!C13、RelatedStart=詳細!B9、RelatedViewport=詳細!B9:I13です。結合入力欄は左上セルを参照します。関連一覧は5件ずつ前後に切り替えます。

## Fields
A:Fが項目表、H:Jが例外表です。列位置を変える場合はConfigurationの表読取部分の変更が必要ですが、各項目の入力セルや管理簿の通常列順変更にはコード修正は不要です。

|FieldId|Header|Type|Sheet|Cell|Required|
|---|---|---|---|---|---|
|MaskedNumber|MaskedNumber|Text|詳細|C20|Yes|
|Branch|枝番|Text|詳細|G20|Yes|
|Office|署|Text|詳細|G18|No|
|PersonName|氏名|Text|詳細|C17|Yes|
|Furigana|氏名(フリガナ)|Text|詳細|G17|No|
|DocumentName|文書名|Text|詳細|C22|No|
|TaxType|税目|Text|詳細|C23|No|
|TaxYear|年分|Text|詳細|G22|No|
|Status|状態|Text|詳細|G23|No|
|ReceivedDate|収受日|Date|詳細|C24|No|
|DraftDate|起案日|Date|詳細|G24|No|
|ApprovedDate|決裁日|Date|詳細|C25|No|
|IssuedDate|施行日|Date|詳細|G25|No|
|Memo1|メモ1|Text|詳細|C28|No|
|Memo2|メモ2|Text|詳細|C29|No|
|Memo3|メモ3|Text|詳細|C30|No|
|ApprovalNumber|決裁番号|Text|詳細|C26|No|
|InquiryNumber|お問い合わせ番号|Text|詳細|G26|No|
|PostalCode|郵便番号|Text|詳細|C19|No|
|Address|住所|Text|詳細|C18|No|
|PhoneNumber|電話番号|Text|詳細|G19|No|

Headerは管理簿1行目の文字と完全一致させます。氏名(フリガナ)の括弧、全半角、空白を自動変換しません。A列MaskedNumber、B列枝番は固定です。他の項目は管理簿列を並べ替えても見出しで解決します。

TypeはTextまたはDate、RequiredはYesまたはNoです。Cellは単一アドレス・数式なしのセルです。結合入力欄は左上セルを指定します。Text欄を数値から文字列へ推測補完しません。数値化で失われた先頭0はエラーとして確認を求めます。

### 追加項目の例
管理簿へ「担当者名」列を作り、Fieldsへ次の1行を追加します。

|FieldId|Header|Type|Sheet|Cell|Required|
|---|---|---|---|---|---|
|AssigneeName|担当者名|Text|詳細|C35|No|

B35に表示ラベルを配置し、C35の入力欄を文字列形式にします。再読込後、詳細表示・新規登録・更新の共通処理へ入ります。検索画面の追加FieldIdにAssigneeNameと入力すれば、その値も部分一致検索できます。クラス・登録処理・更新処理は変更しません。

2025年度に追加列がない場合は、同じFieldsのH:Jへ例外を1行設定します。

|H Year|I FieldId|J HeaderOverride|
|---|---|---|
|2025|AssigneeName|[Absent]|
|2025|Furigana|氏名（フリガナ）|

最初の行は当該年度でその項目を欠落扱いにします。次の行は例外例であり、実際の見出しが全角括弧の場合だけ使います。必須項目・MaskedNumber・Branchは[Absent]不可です。Absent項目に入力値がある場合、黙って捨てず保存を中止します。

### 入力欄の移動
例えばAddressのCellをC18からC37へ変更し、Excelでラベルと入力欄を移し、StartControllerで再読込します。VBA本体の修正は不要です。入力中に設定を動かさず、先に保存または内容を控えてください。

## 名前定義（操作画面の配置）
入力項目以外の操作領域はExcelの「数式→名前の管理」で位置を変更します。全てブック範囲の名前です。表全体を移す場合は列数・行数を維持します。

|名前|初期参照|
|---|---|
|SearchName|検索!C3|
|SearchYear|検索!E3|
|SearchField|検索!C5|
|SearchValue|検索!E5|
|SearchResults|検索!B10:F21（12行×5列）|
|SearchState|検索!B6|
|DetailYear|詳細!C5|
|DetailState|詳細!B2|
|NumberPaste|取込!C7|
|ImportPaste|取込!C13|
|RelatedStart|詳細!B9（5件ページ）|

関連一覧の表示領域はRelatedViewport=詳細!B9:I13です。この範囲に入力項目を置きません。5件ずつ前後ページで全年度の関連事案を表示します。Calendarシート内の日付グリッドはB5:H10の共通部品です。通常の詳細入力欄の配置変更はFieldsのみで済みます。ボタンはExcel標準のフォームコントロールなので、手動で移動できます。

## 架空データ例
以下はすべて開発用に合成した値です。人物の実番号として扱いません。

|本来の番号（試験用）|A列保存値|B列枝番|
|---|---|---|
|0000000000000|MN1:7777777777777:00|00001|
|0123456789012|MN1:9876543210987:77|00001|
|9999999999999|MN1:6666666666666:43|00001|
|1234567890123|MN1:0987654321098:68|00001|

2025と2026の両管理簿に2行目としてA列=MN1:9876543210987:77、B列=00001、氏名=架空 一郎、郵便番号=0010001、電話番号=09000000000を文字列で入れます。2026の3行目にA列=MN1:6666666666666:43、B列=00001、氏名=架空 二郎を入れます。年分・署・文書名・税目等も文字列にします。日付欄にはExcelで2026/11/4等を入力します。空欄可の項目は空欄で構いません。

期待する検索結果：氏名「架空」で2人、架空 一郎の関連事案は2件です。同じ氏名で番号または枝番が異なる別行を足すと、別人物として増えます。

## 暫定テキスト取込例
ラベル名はFieldsのHeaderまたはFieldIdに完全一致させます。番号だけは専用ラベル「識別番号」に18桁を指定します。

```text
氏名：架空 三郎
氏名(フリガナ)：カクウ サブロウ
郵便番号：001-0001
住所：架空県架空市架空町1
電話番号：090-0000-0000
識別番号：123456789012300001
```

ラベルと値は同一行のコロン（:／：）またはタブで区切ります。行順は不問です。見つからない項目の既存入力は保持します。重複ラベルや18桁不正は解析を止めます。実際のコピー文字列が別形式の場合はUtilities.ParseImportのラベル抽出部分を実サンプルで確定してください。未知のレイアウトから推測して自動保存しません。
