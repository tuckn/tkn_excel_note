# Changelog

## 4.0.0 - 2026-10-07

- Excel代理ノートを `type: "excel"` / `schemaVersion: "3.0.0"` へ変更。先頭5項目、意味別の中央項目、末尾 `tags`・`created`・`updated`・`noteId` に整列し、フラットなYAMLを出力。
- `date` を `created` へ引き継ぎ、日時を日本時間・秒単位・ダブルクォート付きに統一。通常更新とAI生成後更新に共通適用し、Frontmatter終端と本文の間に空行を保証。
- `python -m excel_catalog_pipeline.note_migration` で既存ノートのYAMLのみを移行可能に。`--dry-run`、全件事前検証、バックアップ、同時編集検出、途中失敗時の復元に対応。本文・識別子・独自項目の値を保持。
- 既存の `date` / `type: Excel` を参照する外部スクリプトやBasesは、`created` / `type: excel` に更新してください。独立した `export` のスキーマは継続。

- `config show` を `config list` に変更。既定出力を `git config --list` と同様の1行ごとの `key=value` にし、Windows パスをそのままコピーできる表記に統一。
- `config list --json` で1行の構造化 JSON を出力。読み込んだ設定の版・正規化の有無、値の決定元、既定の生成 profile を両形式で表示。設定確認のログは標準エラーへ出力。
- 更新後は `uv tool install . --reinstall` を実行し、既存の `config show` 呼び出しを `config list`（JSON 利用時は `config list --json`）に置き換えてください。設定ファイルの変更は不要。

## 3.3.1 - 2026-10-03

- `pull --source <id> "example.xlsx"` でsource内の1ブックを選択可能に。`--sheet`・`--context`・cover・`--dry-run` と組み合わせ、既存の代理ノートと同期記録を引き継ぐ。
- 相対パスをsourceの `workbooks_dir` 基準で解決し、範囲外・除外・存在しないブックと `--all-sources` の併用を実行前に拒否。選択外のノート・同期記録を保持。

## 3.3.0 - 2026-10-02

- ノートの処理を現行構成に一本化し、旧本文・Frontmatter・シート見出し・ブック単位抽出欄の自動変換を削除。
- 旧バージョン向けの移行説明を削除。シート全体のcontext集計、既存本文の保持、生成記録による保護は継続。

## 3.2.2 - 2026-10-02

- Workbook Mapをタイトル直下へ移動し、ブック要約と各シートの説明より先に表示。通常pull・export・context付き出力で順序を統一。

## 3.2.1 - 2026-10-02

- `pull --all-sources` を復元。`--source <id>` との選択式で登録済みの全 source を処理し、`--context`・cover・`--dry-run` にも対応。
- 対象の省略・併用は実行前にエラーにし、複数 source の通常取り込み・dry-run・AI context の対象選択を検証。

## 3.2.0 - 2026-10-02

- シートごとの説明・出典画像の後に `#### Extracted Text` を配置。通常pullと独立exportでも、シート単位の構成と最後のWorkbook Mapを共通化。
- 通常pullでは既存contextと使用profileを保持し、抽出欄だけ更新。旧ブック全体の抽出欄を次回pullで移行し、手書き本文とcontextの保護判定を維持。
- 代理ノートの文字数上限で省略した抽出と、空・抽出非対応を区別して表示。

## 3.1.0 - 2026-10-01

- `generation.generators.<id>` に接続・文章profile・Bridge上書き・共通補足を複数保存。CLI、source、既定の順でgeneratorを選択。
- `--reference` / `--reference-file` の併用・複数指定、source固有の補足、`--no-reference` に対応。Excel由来の証拠を優先する参照ルールを両生成段階と独自profileへ適用。
- 補足ファイルを実行前に検証・読み取り、同じ内容を生成全体で使用。補足の内容と参照ルールを再利用判定へ追加し、本文を含まない補足の出所・文字数・ハッシュを記録。
- 設定スキーマを2.1.0へ更新。従来の単一生成設定と対応する旧スキーマは読み取り時に正規化し、ユーザー設定を自動変更しない。

## 3.0.0 - 2026-10-01

- 単体・フォルダの独立した Markdown 書き出しを `export` に分離。既定は原本の隣の `<Excelファイル名>.md`、フォルダは再帰探索。
- `export --context` は共通のシート解析・ブック要約・profileを使用。同期記録や生成キャッシュを使わず、出力の置き換えは `--force` 指定と全生成の成功後に限定。
- `pull` / `push` / `status` / `adopt` / `delete-notes` の `--source` を必須化。単体 pull/push と `--all-sources` を廃止。
- 独立出力に `type: ExcelExport` と原本の識別値・生成日時を記録し、pushの対象から除外。
- READMEを単体書き出し、フォルダ書き出し、source同期の順に更新。sourceのcover・context再利用・メタデータ同期は継続。


このプロジェクトの主な利用者向け変更を記録します。
既存の履歴は Git の変更内容と `pyproject.toml` のバージョンから再構成しています。
日付はバージョン更新コミットの日付であり、公開リリース日を示すものではありません。
同じバージョンのまま行われた変更は、その版の項目に含めています。
Git 履歴で独立したバージョン更新を確認できない版は掲載していません。

## [Unreleased]

### 2.0.0

- 生成言語をprofileのテンプレートに一本化し、既定profileを `default-ja` に変更。英語は `default-en`、その他の言語は独自profileで指定します。
- `generation.language` と `prompt_profile: auto` / `--profile auto` を廃止。旧設定には移行方法を示すエラーを返し、設定ファイルは自動変更しません。
- 設定スキーマを `"2.0.0"` に更新。廃止項目を除いた従来の `1` / `"1.0.0"`～`"1.3.0"` は引き続き読み込めます。

### 1.4.0

- 単一シートの `pull --context` でも、取得済みの全現存シートのcontextからブック要約を更新。取得済み・未取得の一覧も累積範囲に統一。
- ブック要約、ブック内の順に並ぶシート、補助テキスト、Workbook Mapの順へ整理。シート要約を共通化し、結論・要点・内容・不確実な点は必要な場合だけ表示。「内容」の下位構成は可変。
- 日英profileをプロンプト・JSON Schema・Markdownテンプレートの組に拡張。`generation.profile_dirs` によるユーザー定義profileと、CLI・設定による選択に対応。各段階のリソース変更を再利用判定へ反映。
- シート別の原本識別値とprofile名・版・ハッシュを保存。古い・未検証のcontextをブック要約とFrontmatterで識別。
- 取り込み済みcontextを保持し、記録と一致する旧ブロックだけ見出しを移行。選択外の手編集済み本文は未検証としてそのまま統合。
- 設定スキーマを1.3.0、アプリを1.4.0へ更新。旧1系設定は読み取り時に正規化。

### 1.3.0

- `pull --cover sheet` で、保存済みシートの指定範囲から高解像度 PNG の cover を生成。初回既定は先頭の表示ワークシート、A1:Q50、幅 2400 px。AI 接続は不要。
- `--cover-sheet` / `--cover-range` / `--cover-width` と `cover` 設定を追加。生成条件を同期記録に残し、通常 pull で再利用・更新。`--cover embedded` で従来方式へ切り替え可能。
- 手動 cover と画像化失敗時の既存 cover を保護。dry-run は Excel を起動せず、予定だけを確認。
- 設定スキーマを 1.2.0 に更新。旧1系の設定は読み取り時に正規化し、設定ファイルは自動変更しない。

### 1.2.0

- pull の全 source 処理には --all-sources を必須化。対象未指定は設定を読み込まずヘルプだけを表示。単体ファイル・--source・--all-sources の併用はエラー。

### 1.1.0

- 通常の pull でシート別の保存範囲・値/数式の範囲とセル数・テーブル/図形/画像/グラフ数を取得し、Workbook Map の表と同期状態に記録。Excel 起動・AI 呼び出しは不要。
- 未対応・取得失敗は unknown と理由を表示。既存のシート説明は保持。

### 1.0.3

- 同期に使用していない Frontmatter `files: []` の自動追加を廃止。削除済みの項目は再追加せず、既存の値は保持。

### 1.0.2

- `pull --ai` を `pull --context` に変更。単体・一括の両方で利用でき、生成・再利用・dry-run の挙動は維持。旧オプションの別名は提供しない。

### 1.0.1

- アプリの保存領域を `~/.tkn/excel_note/` に統一。旧名への自動切り替えを廃止し、設定・同期記録・生成履歴をフォルダ全体の名前変更で引き継ぐ手順を記載。

### 1.0.0

- 製品・配布・CLI 名を `tkn-excel-note`、リポジトリ名を `tkn_excel_note` に変更。
- `export` と `context build` を `pull --ai` に統合。Frontmatter、ブック全体の説明、シート説明、根拠画像を同じ代理ノートで管理。
- source 登録不要の `pull <workbook> [--output <note>]` と `push <note>` を追加。単体と一括で同期・競合判定・基準値を共有。
- 変更のないシート説明とブック全体の説明を再利用。AI なしの更新で古くなった説明は `contextStatus: stale` として保持。
- 対象・省略シートと生成日時を Frontmatter に記録。手直ししたシート説明・ブック概要を保護し、明示的な `--ai --force` で置き換え可能。
- 単体から source 管理へ移行してもノート名・ID・未反映のメタデータ・基準値を保持。内容が同じ別の既存ブックを名前変更として取り違えないよう修正。
- `push` 時に抽出本文を短縮しないよう修正。同期コマンドの標準出力を 1 行の結果 JSON に統一。
- 新規保存領域は `~/.tkn/excel_note/`。既存の旧領域がある場合は継続利用し、ノート・固定 ID・管理マーカーも維持。
- README を単体 Markdown 化、更新、Frontmatter 反映、フォルダ一括管理の順へ再構成。

### 0.9.0

- `context import` コマンドを廃止。過去に取り込んだシート説明は保持し、通常の `context build` による上書きも引き続き防止。

- 読み取り専用のシート一覧コマンドを `context sheets` から `workbook list-sheets` に変更。

- 設定名を `sources.<id>.workbooks_dir` と `sources.<id>.notes.dir` に変更。旧名の読み込みを維持し、設定の表示・同梱例は新名に統一。設定形式は `1.1.0`。

- 同期設定不要の `export <workbook> --output <markdown>` を追加。シート画像解析とブック全体の統合で AI context 用 Markdown を生成。
- 日英の生成プロンプトを `context_profiles/default-ja` / `default-en` に外出し。`export` と `context build` の `--profile`、設定の `generation.prompt_profile` に対応。
- 書き込みなしの dry-run、既存出力の明示的な上書き、生成中の変更検知、出典画像・プロンプトハッシュ・使用量記録を追加。


### 追加

- `delete-notes` を追加。元ブックが存在しない代理ノートを名前・パスで個別指定、または `--all-missing` で一括指定して削除できる。`--dry-run` で保存せずに確認可能。
- 削除前にノート全文・同期記録・復旧用の対応表をバックアップし、該当する同期記録だけを解除。曖昧な照合、存在する元ブック、読み取り失敗を保護し、書き込み失敗時は削除済みノートの復元を試行。
- アプリのバージョンを `0.8.0` に更新。

### 変更

- アプリを `0.8.1`、`tkn-genai-bridge` を `0.10.0` の固定リビジョンへ更新。
- シート説明の画像入力で Claude Code、GitHub Copilot、Antigravity も利用可能に。共有プロファイルによる切り替えと Ollama の `local-vision` の利用手順を追加。

- 設定の `schema_version` を `"1.0.0"` に変更し、生成 AI の設定セクションを `generation` に統一。設定例と表示も更新。
- 旧整数版 `1` と `context` は読み込み時に変換。各設定層で版・値を検証し、未知の版や同一層の新旧セクション併記を拒否。

## [0.7.0] - 2026-09-27

### 追加

- `pull` でブック内のサムネイルを PNG 化し、Obsidian Bases のカード表示用に Frontmatter `cover` を設定。
- Windows の EMF / WMF と PNG / JPEG に対応。画像はノートルートの `img/` に保存し、Vault 内リンクで参照。
- 手動の `cover` を保持し、自動生成画像の更新・欠損修復と書き込みなしの dry-run に対応。変換失敗は警告として報告。

### 修正

- 通常の同期と `context import` で `sourceFullPath` をシングルクォート付きで出力し、パス内のアポストロフィを正しく保持。

### ドキュメント

- `CHANGELOG.md` を追加。
- README の PowerShell 用コード例の言語指定を `shell` に統一。

## [0.6.0] - 2026-09-27

### 変更

- シート説明の生成を画像対応の `tkn-genai-bridge` 0.8.0 へ移行。プロセス実行、JSON Schema検証、使用量集計はBridgeに集約。
- シート画像化に必要な Python 依存関係を通常のインストールに含め、`[context]` の追加指定を廃止。
- `context.bridge_profile` と `context.overrides` で共有接続設定を使用。Codex、Ollama、Azure OpenAIの画像入力に対応。
- 明示された旧接続設定は読み込み時に上書き設定へ変換し、ユーザー設定ファイルは保持。未指定項目は共有プロファイルを使用。
- Bridgeの生成条件・版・スキーマを再利用判定に含め、共有モデル変更を反映。
- 成功・失敗時のBridge実行記録と画像ハッシュを使用量記録に追加。不明な総量と既知小計は区別して保持。
- dry-runは共有設定を検証し、描画・認証・生成・保存は実施しない。既存ノートの編集保護と取り込み済み説明の保持を維持。

## [0.5.0] - 2026-09-23

### 変更

- 設定の `sources` を、`id` を持つリストから source ID をキーにしたマッピング形式へ変更。同梱テンプレートと `config show` の出力も新形式に統一。
- 旧リスト形式の読み込みは維持し、利用者の設定ファイルは自動で書き換えない。
- 設定を重ねて読み込む場合、`sources` は従来どおり上位設定の値で全体を置き換える。

### 修正

- YAML 設定の明示的な重複キーをエラーにし、設定値が黙って上書きされることを防止。
- source のキーと明示的な `id` が一致しない設定を検証エラーにする。

### ドキュメント

- README を現行の設定形式と利用手順に合わせて更新。
- `README_ja.md` を削除し、日本語の説明を `README.md` に集約。

## [0.4.2] - 2026-09-23

### 変更

- `context build` と `context import` のシート見出しを `## <シート名> (sheetId: <ID>)` に統一し、シート ID を明示。

## [0.4.1] - 2026-09-22

### 追加

- `context build` を追加。Excel のシート画像とセル・図形の情報を使い、Codex による説明文と根拠画像を代理ノートへ保存。
- `context import` を追加。既存の Markdown 説明文と参照画像を取り込み、シート単位の管理領域へ保存。
- シート描画・画像処理用の任意依存関係 `context` を追加。

### 変更

- 代理ノートの形式を `schemaVersion: "2.1"` に更新し、元ブックの場所を示す `sourceFullPath` を追加。
- 旧 `Overview` と `Workbook Path` の本文セクションを移行する処理を追加。

## [0.2.0] - 2026-08-19

### 変更

- `pull`、`push`、`adopt` は通常実行で書き込む動作へ変更。変更予定だけを確認する場合は `--dry-run` を指定する。
- `--dry-run` ではノート・ブック・状態・実行レポートを保存しない。
- 旧 `--write-notes` / `--write-excel` は非推奨の互換オプションとして維持。
- 旧 `.xls` の変換スクリプトに `-DryRun` を追加。
- `config show` は、有効な設定ファイルのパスに続けて、解決済み設定をインデント付き JSON で表示する形式へ変更（2026-09-21）。

## [0.1.0] - 2026-08-03

### 追加

- `.xlsx` / `.xlsm` のメタデータと Markdown 代理ノートを同期する CLI を追加。
- `pull`、`push`、`adopt` による同期・追跡と、競合の検出、実行結果のレポートを実装。
- ノートのテンプレート読み込みと検証、再帰的なフォルダ探索、除外パターン、`status` による集計表示を追加。
- Excel 書き込み時のバックアップ、一時ファイル、整合性検証を強化。

### 変更

- 代理ノートにスキーマバージョンを導入し、2.0 で OOXML のプロパティに対応したメタデータ項目と状態移行を整備。
- `pull` のメタデータ保持とレポート表示を改善。
- Frontmatter の用語を Obsidian リンクまたは通常文字列で出力する `frontmatter_term_format` を追加。
- CLI 名を `excel-catalog` から `tkn-excel-catalog` へ変更。
