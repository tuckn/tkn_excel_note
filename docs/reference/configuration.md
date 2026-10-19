# 設定リファレンス

`tkn-excel-note` の設定ファイルの場所、読み込む順序、すべての設定キーと既定値をまとめます。
初めて設定ファイルを作る手順は、[README の「設定ファイルを作る」](../../README.md#41-設定ファイルを作る)を参照してください。

## 設定ファイルの場所と優先順位

設定は次の順に読み込み、後に読み込んだ値を優先します。

1. 組み込みの既定値。
2. ユーザー共通の設定 `~/.tkn/excel_note/config.yaml`。`config init` で作成します。
3. コマンドを実行したフォルダの `.tkn/config.yaml`。
4. `--config` で指定したファイル。

`generation` のような項目の集まりは、ファイルをまたいで項目単位で上書きします。
一方、`sources` と、`include` / `ignore` / `profile_dirs` のような一覧は、後のファイルに書くと全体を置き換えます。
同じ source の設定を複数のファイルに分けて書くことはできません。

設定ファイル内の相対パスは、原則としてコマンドを実行したフォルダを基準に解決します。
補足ファイル（`reference.files`）だけは、その値を書いた設定ファイルのフォルダが基準です。

読み込まれた設定ファイルと、有効な値は次のコマンドで確認できます。
`git config --list` と同様、1行に1つの `key=value` を表示します。
入れ子の項目は `.`、配列は `[0]` などで表し、空の配列・項目は `[]` / `{}`、真偽値は `true` / `false`、未設定値は `null` になります。
文字列は引用符で囲まず、Windows パスの `\` は重ねません。改行などの制御文字だけエスケープします。

```shell
tkn-excel-note config list
tkn-excel-note config list --json
```

`--json` は前置きの説明を付けず、1行の JSON を標準出力に出します。
どちらの形式でも進捗は標準エラーに `[INFO] Showing resolved configuration` と表示します。`--quiet` はコマンドの前に指定します。
設定・state・cache・report は作成・更新せず、Bridge の認証情報や補足ファイルの本文も読み取りません。

| 出力項目 | 内容 |
| --- | --- |
| `config` | 解決済みの設定。`loadedConfigFiles` に読み込んだファイルを優先順位の順で列挙します。 |
| `loaded_sources` | 各ファイルの絶対パス、読み込んだ `schema_version`、`effective_schema_version`、読み取り時の正規化の有無（`migrated`）。版が省略された入力は `null` で表示します。 |
| `winning_sources` | 各設定値の決定元。ファイルの絶対パスまたは `built-in`。配列の各要素も表示します。 |
| `effective_schema_version` | 内部で使用する設定スキーマ版。ファイルによる上書き対象ではありません。 |
| `selected_generator` / `selected_prompt_profile` / `selected_bridge_profile` | 既定の生成設定と、その設定が選ぶ profile。source 固有の選択は `config.sources` で確認します。 |

## コマンドが使う設定

| コマンド | 使う設定 |
| --- | --- |
| `export` | `generation` だけを使います。入力と出力先は引数で決まり、`sources`・`sync`・`cover` は使いません。 |
| `pull` | `sources`、`sync`、`cover`。`--context` を付けた場合は `generation` も使います。 |
| `push`、`status`、`adopt`、`delete-notes` | `sources` |

`pull` は `--source <id>` または `--all-sources` のどちらかが必須です。`--all-sources` は登録済みのすべての source を処理します。
`pull --source <id> "example.xlsx"` で1ブックに限定できます。相対パスは `workbooks_dir` が基準で、配下の絶対パスも使えます。
ファイルは `recursive`・`include`・`ignore` の対象範囲内に限ります。`--all-sources` との併用はできません。
`push`、`status`、`adopt`、`delete-notes` には `--source <id>` が必須です。

## source の設定

source は、入力ブックのフォルダと代理ノートの保存先を組にした同期対象です。
`sources.<id>` の `<id>` が、`--source` で指定する名前になります。

| 設定キー（`sources.<id>.` 以下） | 既定値 | 意味 |
| --- | --- | --- |
| `workbooks_dir` | なし（必須） | 入力ブックのフォルダです。 |
| `recursive` | `false` | `true` にすると、サブフォルダのブックも対象にします。 |
| `include` | `['*.xlsx', '*.xlsm']` | 対象にするファイルのパターンです。 |
| `ignore` | `[]` | 対象から外すファイルのパターンです。 |
| `notes.dir` | なし（必須） | 代理ノートを保存するフォルダです。 |
| `notes.profile` | `tkn-obsidian-v1` | 代理ノートのテンプレートです。現在は同梱の `tkn-obsidian-v1` だけです。 |
| `notes.frontmatter_term_format` | `obsidian-link` | `keywords` と `categories` の書式です。`obsidian-link` は `[[...]]` のリンク、`plain` は通常の文字列です。 |
| `notes.rename_adapter` | `report-only` | ブックの名前変更・移動の扱いです。`report-only` は報告だけ、`filesystem` はファイルを移動します。条件は「[名前変更と移動](synchronization.md#名前変更と移動)」を参照してください。 |
| `generation.generator` | なし | この source で使う generator です。「[名前付き generator と補足文](#名前付き-generator-と補足文)」を参照してください。 |
| `generation.reference` | なし | この source で使う補足文です。 |

`include` と `ignore` には、入力フォルダからの相対パスのパターンを、区切り文字 `/` で書きます。
たとえば `ignore` に `archive/**` を指定すると、`archive` フォルダ以下を除外します。

`recursive: true` の場合、入力フォルダのフォルダ構成を、代理ノートの保存先にも再現します。
入力フォルダの `2026/example.xlsx` は、`notes.dir` の `2026/example.xlsx.md` になります。

次は、サブフォルダを含めて探索し、一部を除外する `sources` の例です。

```yaml
sources:
  workbooks:
    workbooks_dir: 'C:\path\to\excel-workbooks'
    recursive: true
    include:
      - "*.xlsx"
      - "*.xlsm"
    ignore:
      - "archive/**"
      - "**/Temp/**"
    notes:
      dir: 'C:\path\to\excel-notes'
      profile: tkn-obsidian-v1
      frontmatter_term_format: obsidian-link
      rename_adapter: report-only
```

Windows のパスをシングルクォートで囲むと、バックスラッシュを二重にせずに書けます。

## 同期の設定

| 設定キー | 既定値 | 意味 |
| --- | --- | --- |
| `sync.max_extracted_text_chars` | `12000` | 代理ノートの「Extracted Text」に出力する、1ブックあたりの最大文字数です。`export` には適用しません。 |

`config init` で作成した設定ファイルには、`sync.pull_preserves_user_metadata`、`sync.delete_missing_notes`、`sync.delete_missing_workbooks`、`sync.allow_source_rename` も含まれます。
これらの値は読み込まれますが、処理の動作は変わりません。

- 同期では、未知の Frontmatter 項目と管理マーカー外の本文を常に保持します。
- 同期で、ブックや代理ノートを自動で削除することはありません。元ブックがない代理ノートを削除するには、[`delete-notes`](../guides/catalog-operations.md#元ブックがない代理ノートを削除する)を使います。
- 名前変更・移動は、`notes.rename_adapter` と `push --allow-rename` で制御します。

## cover の設定

トップレベルの `cover` は、`pull` で作るカード表示用の画像の作り方を指定します。
`export` は cover を作りません。

| 設定キー | 既定値 | 意味 |
| --- | --- | --- |
| `cover.mode` | `auto` | `embedded` はブックに埋め込まれたサムネイル、`sheet` はシートの指定範囲の画像です。`auto` は、ブックごとに最後に成功した方式を使い、記録がなければ `embedded` を使います。 |
| `cover.sheet` | `null` | 画像にするシートの名前です。 |
| `cover.range` | `null` | 画像にするセル範囲です。 |
| `cover.width` | `null` | 画像の幅（ピクセル）です。600～4000 の整数を指定します。 |

`null` は、ブックごとに最後に成功した条件を使うという意味です。
記録がないブックでは、先頭の表示中のワークシート、範囲 `A1:Q50`、幅 2400 px を使います。

条件は次の順に決まります。

1. コマンドラインのオプション（`--cover`、`--cover-sheet`、`--cover-range`、`--cover-width`）。
2. 設定ファイルの `null` 以外の値。
3. ブックごとに最後に成功した条件。
4. 記録がない場合の既定値。

設定ファイルで `range` や `width` を指定すると、ブックごとにコマンドラインで選んだ条件より優先されます。
ブックごとの条件を残したい場合は、`mode: sheet` だけを指定します。

```yaml
cover:
  mode: sheet
```

設定ファイルで `mode: embedded` を指定すると、自動で作ったシート画像の cover も埋め込みサムネイルに切り替えます。
手動で設定した cover は、どの方式でも上書きしません。

`null` に戻しても、ブックごとの記録は初期値に戻りません。
すべての条件を既定値に戻すには、`--cover embedded` で一度成功させた後に `--cover sheet` を実行し、設定ファイルの個別の指定も削除します。

シート画像の方式には、Windows とデスクトップ版 Excel が必要です。
AI の設定は不要です。
画像の仕様は「[シートの指定範囲を画像にする](../guides/catalog-operations.md#シートの指定範囲を画像にする)」を参照してください。

## AI 生成の設定

トップレベルの `generation` は、`export --context` と `pull --context` で使います。
AI の接続先・モデル・認証は、このファイルではなく GenAI Bridge の共有設定 `~/.tkn/genai_bridge/config.yaml` で管理します。
使い方は「[AI によるシート説明とブック要約](../guides/sheet-content.md)」を参照してください。

### 名前付き generator と補足文

generator は、AI の接続先（Bridge の profile）、文章の profile、Bridge の設定の上書き、補足文を組にして名前を付けた設定です。
複数の generator を定義し、`--generator <id>` や source ごとの設定で切り替えます。

```yaml
generation:
  default_generator: my-codex-def
  generators:
    my-codex-def:
      bridge_profile: codex-default
      prompt_profile: default-ja
      overrides:
        timeout_seconds: 600
      reference:
        text: |
          技術調査や設計検討のノートを扱います。
        files:
          - ./references/technical-terms.md
    my-codex-en:
      bridge_profile: codex-default
      prompt_profile: default-en
  max_reference_chars: 20000
sources:
  workbooks:
    workbooks_dir: 'C:\path\to\workbooks'
    notes:
      dir: 'C:\path\to\notes'
    generation:
      generator: my-codex-def
      reference:
        text: "この source の案件の背景を記載します。"
        files:
          - ./references/project-background.md
```

| 設定キー（`generation.generators.<id>.` 以下） | 省略時の値 | 意味 |
| --- | --- | --- |
| `bridge_profile` | `codex-default` | GenAI Bridge の共有設定にある接続先の名前です。 |
| `prompt_profile` | `default-ja` | 文章の profile です。 |
| `overrides` | `{}` | 共有設定の値を、このツールだけで上書きする値です。GenAI Bridge の設定仕様に従います。 |
| `reference.text` | なし | 補足文です。 |
| `reference.files` | なし | 補足文のファイル（UTF-8）の一覧です。相対パスは、この設定ファイルのフォルダが基準です。 |

使う設定は次の規則で決まります。

| 対象 | 決まり方 |
| --- | --- |
| generator | `--generator` → `sources.<id>.generation.generator` → `generation.default_generator` の順で、最初に指定されたものを使います。どれもなければ、`generation` 直下の `bridge_profile`・`prompt_profile`・`overrides` を使います。 |
| 文章の profile | `--profile` を最優先します。指定がなければ、generator を選んだ場合はその `prompt_profile`（省略時 `default-ja`）、選ばない場合は `generation.prompt_profile`（既定 `default-ja`）を使います。 |
| 補足文 | 選んだ generator の補足 → source の補足 → `--reference` → `--reference-file` の順に、すべてを合わせて使います。選ばなかった generator の補足は使いません。 |
| 補足の無効化 | `--no-reference` を指定すると、設定ファイルの補足も含めてすべて使いません。`--reference` や `--reference-file` とは併用できません。 |

`export` は source の設定を使わないため、generator は `--generator` または `generation.default_generator` で決まり、source の補足は使いません。
`--generator`、`--profile`、`--sheet`、`--reference`、`--reference-file`、`--no-reference` は、`--context` と一緒に指定します。

補足文の合計文字数は、`generation.max_reference_chars`（既定 20000 文字）までです。
1回の生成に渡すすべての補足文の合計に適用します。
超えた場合は、AI を呼び出す前にエラーになり、補足文を切り詰めません。
補足ファイルは、空でない UTF-8 のテキストにします。
BOM 付きでも読み込めます。

複数の設定ファイルに同じ名前の generator を書いた場合は、項目単位で上書きします。
`reference.text` は後のファイルの値で置き換え、`reference.files` は一覧全体を置き換えます。
`text: ""` や `files: []` を書くと、その項目を空にできます。
未定義の generator 名、未知の項目、型の誤りはエラーになります。

`config list` は補足文の設定を表示しますが、補足ファイルの中身は読み取りません。

### 画像化と抽出の上限

次の項目は、generator ではなく `generation` の直下に書きます。

| 設定キー（`generation.` 以下） | 既定値 | 意味・単位 |
| --- | --- | --- |
| `profile_dirs` | `[]` | 独自の文章 profile を置いた親フォルダの一覧です。記載順に探します。 |
| `max_images` | `24` | 1シートあたりの画像の枚数の上限です。全体の縮小図を含みます。最大 `100` です。 |
| `tile_width_points` / `tile_height_points` | `1200` / `800` | 分割した画像1枚が表す範囲の幅・高さです。Excel のポイント単位です。 |
| `overlap_points` | `80` | 隣り合う画像の重なりです。幅と高さの小さい方の半分未満にします。 |
| `image_dpi` | `150` | 画像の解像度です。`72`～`300` dpi を指定します。 |
| `max_cells` / `max_objects` | `10000` / `10000` | 1シートから抽出するセル・オブジェクトの数の上限です。 |
| `max_input_chars` | `null` | AI に送る抽出の根拠データ（JSON）の文字数の上限です。`null` はこのツールでは制限しません。 |
| `max_workbook_mb` | `100` | 入力ブックのファイルサイズの上限です。単位は MiB（1024 × 1024 バイト）です。 |

`max_input_chars` 以外の数値は、正の整数で指定します。
上限を超えた場合は、対象を切り詰めずにエラーとして停止します。
上限を変えると、`pull` で前回の説明を再利用できなくなる場合があります。

`max_input_chars` は、根拠データの JSON の文字数だけに適用し、画像とプロンプトの固定部分は含みません。
根拠データは、通常の JSON より短くなる場合、列名を一度だけ書く表形式で送ります。
実際に送る文字数は、AI を呼び出す前に標準エラー出力と使用量の記録に出力します。
`null` にしても、GenAI Bridge やモデル側の入力の上限はなくなりません。
`max_input_chars` に数値を指定した場合は、その文字数を上限とします。

### 文章の profile

| 設定キー（`generation.` 以下） | 既定値 | 意味 |
| --- | --- | --- |
| `prompt_profile` | `default-ja` | generator を使わない場合の文章の profile です。英語は `default-en` です。 |
| `profile_dirs` | `[]` | 独自の profile を置いた親フォルダの一覧です。 |

`profile_dirs` に `C:\path\to\profiles` を指定して `prompt_profile: technical-notes` を選ぶと、`C:\path\to\profiles\technical-notes\` を読み込みます。
同名のフォルダがなければ、同梱の `default-ja` / `default-en` から探します。
profile の検証は `--context` を付けた実行で行い、`config list` では行いません。
profile のファイル構成と作り方は「[独自の文章 profile を作る](../guides/sheet-content.md#独自の文章-profile-を作る)」を参照してください。

## 設定ファイルの版

`schema_version` は設定ファイルの形式の版で、現在は `"2.1.0"` です。

## プレビュー画像の設定

`thumbnails` は `pull` の cover 用画像だけに適用します。VLM 用の `img/` と独立し、EXIF・ICC・XMPなどのメタデータを除去します。

```yaml
thumbnails:
  enabled: true
  subdir: thumbnails
  max_edge: 480
  format: webp
  quality: 85
  lossless: false
```

既定値は `max_edge: null`（従来の寸法を維持）、`lossless: true` です。他の既定値は上記と同じです。`max_edge` は64～4000またはnull、`quality` は1～100、`format` はpng/jpeg/webpを指定できます。`lossless: true` はWebP専用で、PNGは常に可逆圧縮です。`subdir` は `notes.dir` 直下の安全な単一フォルダ名です。`enabled: false` は生成を停止し、既存のcoverを保持します。手動指定のcover・旧画像は変更・削除しません。設定を変えたシート方式は再描画しますが、旧形式の同期済みキャッシュは変換して利用できます。
