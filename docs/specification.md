# CIVIL HUB 仕様書

## 目的

CIVIL HUB は、建設工事における施工体制台帳作成と添付書類確認を支援する Windows デスクトップアプリ風の業務アプリケーションである。

工事、施工体制の業者階層、必要添付書類、添付ファイルを一元管理し、施工体制台帳 Excel からの工事情報取込と反映コピー作成を行う。

## 対象範囲

### 対象

- 工事情報の管理
- 施工体制の業者階層管理
- 業者ごとの必要添付書類チェック
- 添付ファイルの登録と保存
- 不足、未確認、期限切れ書類の確認
- 工事情報 CSV 取込
- 施工体制台帳 Excel 取込
- 施工体制台帳 Excel への工事情報反映コピー作成

### 対象外

- 添付ファイル内容の自動解析
- 帳票全項目の自動反映
- 複数ユーザー同時編集
- 権限管理
- ネットワーク同期

## アプリケーション構成

| 項目 | 内容 |
| --- | --- |
| 実行ファイル | `app.py` |
| GUI | tkinter / ttk |
| データベース | SQLite |
| DBファイル | `civilhub.db` |
| Excel処理 | openpyxl |
| 添付初期保存先 | `storage/attachments` |

## 画面構成

### ヘッダー

- アプリ名: `CIVIL HUB`
- サブタイトル: `施工体制台帳作成・添付書類チェックシステム`
- 操作: `不足一覧`

### 左ペイン: 工事一覧

表示内容:

- 工事名
- 工事番号

操作:

- `+ 工事登録`
- `工事編集`
- `CSV取込`
- `台帳Excel取込`
- `台帳反映`
- `工事削除`
- `初期化`
- `工事再読込`

### 中央ペイン: 施工体制ツリー

表示内容:

- 会社名
- 階層
- 担当工種

操作:

- `+ 元請`
- `+ 子業者`

階層表示:

- `元請`
- `一次下請`
- `二次下請`
- `三次下請以降`

### 右ペイン: 選択業者の詳細

選択中の業者情報をテキスト表示する。

操作:

- `書類状態更新`
- `リネーム候補`

### 下部タブ: 添付書類チェック

業者ごとの必要添付書類を表示する。

列:

- 書類名
- 必須
- 状態
- 有効期限
- 添付数
- 備考

操作:

- `状態更新`
- 行ダブルクリックによる状態更新

### 下部タブ: ファイル添付

選択書類に紐づく添付ファイルを表示する。

列:

- 元ファイル名
- 保存ファイル名
- 種類
- 登録日時

操作:

- `+ 添付`
- `保存場所を開く`
- `ファイルを開く`

## データ仕様

### projects

工事情報を管理する。

| カラム | 内容 |
| --- | --- |
| id | 工事ID |
| name | 工事名 |
| construction_no | 工事番号 |
| client_name | 発注者 |
| start_date | 工期開始日 |
| end_date | 工期終了日 |
| site_agent | 現場代理人 |
| managing_engineer | 監理技術者 |
| chief_engineer | 主任技術者 |
| base_folder | 保存フォルダ |
| created_at | 作成日時 |
| updated_at | 更新日時 |

必須入力:

- 工事名
- 工事番号
- 発注者

### companies

施工体制の業者情報を管理する。

| カラム | 内容 |
| --- | --- |
| id | 業者ID |
| project_id | 工事ID |
| parent_company_id | 親業者ID |
| level | 階層 |
| name | 会社名 |
| kana | 会社名カナ |
| representative | 代表者名 |
| address | 所在地 |
| phone | 電話番号 |
| license_no | 建設業許可番号 |
| license_expiry | 許可有効期限 |
| work_type | 担当工種 |
| contract_date | 契約日 |
| planned_start_date | 施工開始予定日 |
| planned_end_date | 施工終了予定日 |
| chief_engineer_name | 主任技術者名 |
| chief_engineer_license | 主任技術者資格 |
| safety_manager | 安全衛生責任者 |
| created_at | 作成日時 |
| updated_at | 更新日時 |

必須入力:

- 会社名

### required_documents

業者ごとの必要添付書類を管理する。

| カラム | 内容 |
| --- | --- |
| id | 必要書類ID |
| company_id | 業者ID |
| document_name | 書類名 |
| required | 必須フラグ |
| status | 状態 |
| expiry_date | 有効期限 |
| note | 備考 |
| created_at | 作成日時 |
| updated_at | 更新日時 |

状態:

- `未確認`
- `不足`
- `添付済み`
- `期限切れ`
- `不要`

業者登録時に、次の書類を自動生成する。

- 再下請通知書
- 建設業許可証
- 主任技術者資格証
- 主任技術者の雇用確認資料
- 社会保険加入確認資料
- 労働保険関係資料
- 安全衛生責任者選任書
- 作業員名簿
- 外国人就労関係書類
- 契約書または注文請書

### attachments

添付ファイルの登録情報を管理する。

| カラム | 内容 |
| --- | --- |
| id | 添付ID |
| required_document_id | 必要書類ID |
| original_path | 取込元ファイルパス |
| stored_path | 保存先ファイルパス |
| original_filename | 元ファイル名 |
| stored_filename | 保存ファイル名 |
| file_type | 拡張子 |
| created_at | 登録日時 |

### project_imports

外部ファイル取込履歴を管理する。

| カラム | 内容 |
| --- | --- |
| id | 取込履歴ID |
| project_id | 工事ID |
| import_kind | 取込種別 |
| source_path | 取込元ファイルパス |
| source_sheet | 取込元シート名 |
| imported_at | 初回取込日時 |
| updated_at | 更新日時 |

## CSV取込仕様

### 標準CSV

対応列:

| CSV列名 | 保存先 |
| --- | --- |
| 工事名 | projects.name |
| 工事番号 | projects.construction_no |
| 発注者 | projects.client_name |
| 工期開始日 | projects.start_date |
| 工期終了日 | projects.end_date |
| 現場代理人 | projects.site_agent |
| 監理技術者 | projects.managing_engineer |
| 主任技術者 | projects.chief_engineer |
| 保存フォルダ | projects.base_folder |

取込条件:

- 工事名と工事番号が空の行は取り込まない。
- 工事番号で既存工事を照合する。
- 既存工事がある場合は更新する。
- 既存工事がない場合は新規登録する。
- 取込前に確認ダイアログを表示する。

### 設定CSV

`project_name` と `excel_output_dir` を含む CSV は設定 CSV として扱う。

| CSV列名 | 保存先 |
| --- | --- |
| project_name | projects.name |
| excel_output_dir | projects.base_folder |

工事番号は `CFG_<工事名>` 形式で仮生成する。

## 施工体制台帳 Excel 取込仕様

対象:

- `.xlsx`
- 先頭シート

読取項目:

| 項目 | 読取方法 |
| --- | --- |
| 工事名 | `工事名称及び工事内容` または `事業所名・現場ID` の右側値 |
| 発注者 | `発注者名及び住所` の右側値 |
| 工期開始日 | `工期` の右側値 |
| 工期終了日 | `工期` の1行下の右側値 |
| 現場代理人 | `現場代理人名` の右側値 |
| 監理技術者 | `監理技術者名主任技術者名` の右側2番目の値 |

日付は `YYYY年M月D日` 形式の場合、`YYYY-MM-DD` に変換する。

照合:

- 工事名で既存工事を照合する。
- 既存工事がある場合は更新する。
- 既存工事がない場合は新規登録する。
- 取込元ファイルパスとシート名を `project_imports` に保存する。

## 台帳反映仕様

実行条件:

- 工事が選択されている。
- 対象工事に `施工体制台帳` の取込履歴がある。
- 取込元 Excel ファイルが存在する。

出力:

- 取込元ファイルは変更しない。
- 同じフォルダに反映済みコピーを作成する。
- 初回出力名は `<元ファイル名>_civilhub.xlsx` とする。
- 同名ファイルがある場合は `<元ファイル名>_civilhub_v2.xlsx` のように連番を付ける。

反映セル:

| projects項目 | セル |
| --- | --- |
| name | `I6`, `G21` |
| client_name | `G25` |
| start_date | `H29` |
| end_date | `H30` |
| site_agent | `G55` |
| managing_engineer | `J57` |

日付が `YYYY-MM-DD` の場合、開始日は `自 YYYY年M月D日`、終了日は `至 YYYY年M月D日` に変換する。

## 添付ファイル仕様

登録条件:

- 工事、業者、書類が選択されている。

選択可能ファイル:

- `.pdf`
- `.xlsx`
- `.xls`
- `.docx`
- `.doc`
- `.png`
- `.jpg`
- `.jpeg`
- `.bmp`
- その他すべてのファイル

保存先:

```text
<工事保存フォルダまたはstorage/attachments/project_ID>/CIVIL_HUB_添付/<会社名>/<書類名>/
```

保存ファイル名:

```text
<会社名>_<書類名>_v<番号><拡張子>
```

登録後処理:

- ファイルを保存先へコピーする。
- `attachments` に登録する。
- 対象書類の状態を `添付済み` に更新する。

## 不足一覧仕様

選択中の工事に紐づく書類のうち、次の状態を一覧表示する。

- `不足`
- `未確認`
- `期限切れ`

表示形式:

```text
<会社名>：<書類名> が<状態>
```

## 削除・初期化仕様

### 工事削除

工事削除時、関連する業者、必要書類、添付情報、取込履歴も削除する。

添付ファイルの実ファイル削除は行わない。

### 初期化

デバッグ用初期化として、次のデータを全削除する。

- 工事
- 業者
- 必要書類
- 添付情報
- 取込履歴

`storage/attachments` 配下も削除して再作成する。

## 既存参照資料

- `docs/ledger-sample-input-mapping.md`
- `docs/ui-operation-screen-reference.md`
