# Changelog

## Unreleased

- `pull --context` / `export --context` で、画像の枚数・セルの数・根拠データの文字数が1シートの上限（`max_images`・`max_cells`・`max_input_chars`）を超えるシートを飛ばさず、長いデータの範囲（Excel のテーブルと、数値や日付を含み、同じ列を埋めた行が30行以上続く範囲）を短くして解析する。文字だけの行（列に書いたメモなど）は短くしない。見出しの行と最初の10行のデータはセルと画像のまま送り、残りの行は、値の多い列のセルを列ごとの要約（根拠データの `dataSummaries`：値と空欄の数、種類、数値の最小・最大、日付の最初・最後、異なる値の数、多い値、最後の値、数式の例）に置き換える。表の横のメモなど値の少ない列のセルは、そのまま送る。画像の枚数のほか、描画後の根拠データの文字数が上限を超えた場合も、短くしてから描画し直す。結果に `summarizedRanges` を追加。短くしても上限を超えるシートと、`max_objects` を超えるシートは、これまでどおり `skipped`。プロンプトは変えないため、生成済みの説明は生成し直さない。
- ファイルサイズが `max_workbook_mb` を超えるブックを、止めずに Excel で開かず画像なしで解析する（結果に `textOnly`）。根拠データにグラフの種類・タイトル・系列の参照範囲を加え、シートの説明に出典画像の欄を作らない。`pull <ブック>` と `export` も、上限を超えるブックで止めない。シート方式の cover は、これまでどおり上限を超えるブックに作らない。
- 同じ固定 ID（`TknExcelCatalogId`）を持つブックがあっても、照合を止めないよう変更。同期記録にあるパスのブック（そのパスがない場合は、同期記録と内容が一致するブック）が ID と代理ノートを引き継ぎ、ほかのブックはコピーとしてブック自身のパスで追跡し、別の代理ノートを作る。結果の `warnings` にコピーとして扱ったことを表示する。`adopt` はコピーに新しい ID を付ける（`push` でコピーへ書き込むときにも付く）。一度コピーとして記録したブックは、元のブックの削除後もコピーとして扱い、`delete-notes` も元のブックの存在の証拠としない。どれが元のブックか決められない場合だけ、これまでどおり `duplicate-id` として止まる。
- 代理ノートの「Extracted Text」を、ワークシートの1行を1項目とする `セル位置: 値` の形式に変更。数式を値に添え、日付・時刻の表示形式のセルを `2026-01-15` の形で、数値を保存時の誤差を除いた値で出力し、ふりがな（ルビ）を除外。同じ値の省略をやめ、図形・図形グループの文字、画像の代替テキスト、グラフのタイトル、VML のテキストボックス、セルのコメントを位置の順に追加。`sync.max_extracted_text_chars` の上限をブック全体から1シートごとに変更し、前のシートの文字数で後ろのシートが省略されないようにした（非互換変更）。次回の `pull` で、すべての代理ノートの抽出欄が新しい形式に更新されます。
- `pull --context` / `export --context` で、1シートを Excel で画像にする処理に時間の上限 `generation.limits.max_render_seconds`（既定600秒）を追加。画像化をシートごとに別のプロセスで行い、上限を超えたら、そのプロセスと、そのシートのためにこのツールが起動した Excel だけを終了して、AI を呼ばずにそのシートを飛ばす（結果は `skipped`）。利用者が開いている Excel は終了しない。1回の実行で2回続けて上限を超えた場合は、Excel 自体が応答していないとみなし、そのブックの生成を止める。これまでは、応答しなくなった Excel で実行が何時間も止まり、外から止めるとロックファイルと一時フォルダが残った。
- `pull --context` / `export --context` で、出力スキーマで検証した AI の出力を画像と同じフォルダの `output.json`（ブック要約は再利用の記録）に保存。テンプレートだけを変更した場合は、AI を呼ばずに表示を作り直す（結果は `rerendered`、`--dry-run` では `would-rerender`）。出力を保存していない既存の説明は、テンプレートを変えても `--force` までそのまま残す（`templateOutdated`）。
- 説明の再利用の判定から、テンプレート、GenAI Bridge の版・タイムアウト・実行ファイル、上限の設定（`max_images`・`max_cells`・`max_objects`・`max_input_chars`・`max_workbook_mb`・`max_reference_chars`）を除外。シートの同一性は、ブック全体の書式定義ではなく、そのシートが使う書式だけで判定し、ほかのシートの編集で書式が増えても生成し直さない。この変更で判定の記録の形式が変わるため、既存の説明は次回の `pull --context` で一度生成し直します。
- AI に渡す根拠データに、日付に直した値（`display`）、セルの書式の一覧（`styles`）、セルのコメント（`comments`）、VML のテキストボックスの文字（`legacyShapes`）を追加。既定の文章 profile（`default-ja`・`default-en`、版 1.1）のプロンプトに、罫線・塗り・字下げ・列のずらしで表した表や階層の読み取りと、追加した根拠データの説明を追加。
- Workbook Map で、VML・OLE オブジェクト・ActiveX コントロールや拡張要素（`extLst`）のあるシートも図形・画像・グラフの数を表示し、数えていない要素を `Notes` に示すよう変更（これまでは `unknown`）。
- シートの説明の再利用の判定から、表示の状態（選択中のシート、セルの選択、スクロール位置、拡大率、表示モード）を除外。ブックを開いてシートを見て回ってから保存しても、生成し直さない。以前の判定で作った記録は、ブックが変わっていなければ、次回の `pull --context` で生成し直さずに新しい判定へ更新する。
- `pull --context` / `export --context` で、1シートの上限（`max_images`・`max_cells`・`max_objects`・`max_input_chars`）を超えるシートと、描画する内容のない空のシートを、ブック全体の停止ではなくそのシートだけの `skipped` として飛ばし、AI を呼ばずに残りのシートとブック要約の生成を続ける。すべてのシートを飛ばした場合はブック要約を生成せず、説明のあるシートがなくなった生成済みのブック要約を外す。`export` の抽出欄では、セル・オブジェクトの上限を超えるシートに理由を書き、ブック全体を止めずに続ける。`max_images` を再利用の判定に含めていた記録は、`max_images` を変えていても、次回の `pull --context` で生成し直さずに新しい判定へ更新する。
- AI の使用量で、キャッシュを除いた入力トークンだけを報告する接続先（Claude Code など）の入力トークン数に、キャッシュの読み込み・書き込みを含める（記録に `totalInputTokens` を追加）。

- 設定形式を `"3.0.0"` に更新。入力ブック側の設定を `sources.<id>.workbooks`（`dir`・`recursive`・`include`・`ignore`）にまとめて `notes` と対にし、`notes.profile` を `notes.template`、`generation.profile_dirs` を `generation.prompt_profile_dirs` に改名。生成の上限（`max_images`・`max_reference_chars` など）は `generation.limits` に、`thumbnails` は `cover.enabled` と `cover.image` に統合。動作に影響しなかった `sync` の4項目（`pull_preserves_user_metadata`・`delete_missing_notes`・`delete_missing_workbooks`・`allow_source_rename`）を廃止（非互換変更）。
- 以前の設定形式の読み替えを廃止。`"3.0.0"` 以外の `schema_version`、旧キー（`workbooks_dir`・`path`・`notes.root`・`notes.profile`・`thumbnails`・`generation.profile_dirs`・`generation` 直下の上限・`bridge_profile`・`language` など）、リスト形式の `sources` と source 内の `id` はエラーになります。既存の設定ファイルは新しい形式に書き換えてください。`config list --json` の `loaded_sources` から `migrated` を削除。`config list` の出力キー・`--source` 実行時の設定表示・設定エラーの文言も新しいキー名に変わります（非互換変更）。
- 代理ノートの `type` を、固有名詞の表記どおり引用符なしの `Excel` に変更し、`schemaVersion` を `"4.0.0"` に更新（非互換変更）。読み取りでは以前の `excel` も代理ノートとして認識し、独自の note profile の `type: excel` もそのまま使用可能。`type: excel` で絞り込む外部スクリプトや Bases は `type: Excel` に更新し、既存ノートは `python -m excel_catalog_pipeline.note_migration` で移行してください。
- `pull --context` / `export --context` の Frontmatter に `contextProfiles` を追加。ノート内のシート説明・ブック要約に使った文章の profile 名を、重複なしの一覧で記録。
- `export` の Frontmatter の引用符を共通のYAML記述規約に統一。日時（`sourceCreated`・`sourceModified`・`generatedAt`、`--context` 時の `updated`・`contextGeneratedAt`）と `schemaVersion` をダブルクォート、空文字列を `""` で出力し、`sourceFullPath` はシングルクォートを維持。項目・値（小数秒・タイムゾーン）と書き出しの `schemaVersion: "1.0"` は変更なし。既存の書き出しを `--force` で再作成すると、引用符だけの差分が出ます。これまで引用符なしだった日時を YAML 1.1 のローダーで日時型として読み込んでいた外部スクリプトは、文字列として受け取るよう更新してください。
- 代理ノートと `export` の Frontmatter で、ハッシュ値（`export` の `sourceSnapshotSha256`、`pull --context` / `export --context` の `contextSourceFingerprint`）を常にダブルクォートで出力。`yes`・`1.0`・`[x]` などのシングルクォートをダブルクォートに変更し、YAML 1.1 または 1.2 で真偽値・数値・日付として読まれるのにこれまで引用符なしだった `y`・`1e3`・`0o17`・日付の形の文字列（シート名・独自項目など）もダブルクォートで囲む。バックスラッシュを含む値と `sourceFullPath` はシングルクォートを維持し、値は変更なし。代理ノートは `schemaVersion: "4.0.0"` のまま、`python -m excel_catalog_pipeline.note_migration` で既存ノートの引用符だけを更新可能。書き出しは `--force` で再作成すると、引用符だけの差分が出ます。

## 5.1.0 - 2026-11-01

- GenAI Bridge を Copilot CLI 1.0.105 対応済みの固定コミット `848769a4842ceacd20875ba0934cdd0fc478dc2e`（0.11.0）へ更新。
- AI 接続設定をアプリ専用 `config.yaml` の `genai_profiles` へ統合し、Bridge の共有設定を読み込まない構成へ変更。既存の共有 profile は利用側設定へ移行が必要。
- 設定形式を 2.2.0 に更新し、generator の接続名を `genai_profile` に統一。旧 `bridge_profile` は互換名として受理。項目単位のマージ、設定元、設定ファイル基準の実行ファイルパス、実行記録の profile 名を保持。

## 5.0.0 - 2026-10-29

- CLI全般の既定出力を1行ごとの `key=value` に変更（非互換変更）。`--json` で機械処理用JSON、`--full-output` で変更なし・skipを含む全対象を表示。両指定は多段サブコマンドの前後でも利用可能。
- 同期結果へ対象別の状態・理由・パスを追加。失敗・要確認対象を残し、省略した詳細には確認方法を表示。保存レポートとPython APIの戻り値は維持。
- 引数エラー、設定エラー、実行時エラー、中断も指定形式で出力。進捗・診断を標準エラーへ分離し、quiet/verboseはログ量だけを変更。既存終了コード0/1/2/3を維持し、中断は130。
- ノート移行CLIとPowerShell変換スクリプトも一覧既定へ変更。PowerShellは `-Json` を使用。

## 4.2.0 - 2026-10-17

- `thumbnails` 設定でcoverの有効化・保存先・最大辺・形式・品質・可逆圧縮を指定可能に。VLM用画像とは独立し、メタデータを除去。
- 同期済みの旧coverキャッシュを変換に再利用。手動cover・本文画像・旧ファイルは保持。

## 4.1.0 - 2026-10-17

- cover 用サムネイルを `thumbnails/` の可逆圧縮 WebP に変更。埋め込み方式とシート方式の両方に適用し、寸法を維持。
- 次回の `pull` で自動管理の cover を更新。旧画像・手動 cover を保持し、VLM 用出典画像は従来の `img/` の PNG を継続。

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
