# 代理ノートの管理

代理ノートを Obsidian で使うための画像の設定、元ブックがなくなったノートの削除、旧形式の `.xls` の変換を説明します。
同期の基本は [README](../../README.md) を参照してください。

## Obsidian のカードにブックの画像を表示する

`pull` は、Frontmatter の `cover` に、ブックの画像へのリンクを設定します。
Obsidian Bases のカードビューで `cover` を画像として指定すると、ブックごとのサムネイルを表示できます。

画像の作り方は2つあります。

| 方式 | 画像の元 | 必要な環境 |
| --- | --- | --- |
| `embedded`（既定） | Excel が保存時にブックに埋め込んだサムネイル | なし（Excel を起動しません） |
| `sheet` | シートの指定範囲を描画した画像 | Windows とデスクトップ版 Excel |

どちらの方式も AI を使わず、画像を外部に送信しません。

### カードビューを設定する

既存の Base の `views` に、次の設定を追加します。
Obsidian の画面からカードビューを作り、画像のプロパティに `cover` を指定しても同じです。

```yaml
  - type: cards
    name: Cards View
    order:
      - file.name
      - title
      - description
    image: note.cover
    imageFit: contain
    cardSize: 200
```

カードビューの設定項目は [Obsidian のヘルプ](https://help.obsidian.md/bases/views/cards)を参照してください。

### 埋め込みサムネイルを使う

既定の `embedded` 方式では、`pull` がブックに埋め込まれたサムネイルを 既定で可逆圧縮の WebP に変換し、`notes.dir/thumbnails/excel-cover-<画像のハッシュ値>.webp` に保存します。
Frontmatter の `cover` には、Vault 内の画像へのリンクが入ります。
次は、代理ノートを Vault の `catalog` フォルダに保存している場合の例です。

```yaml
cover: "[[catalog/thumbnails/excel-cover-<画像のハッシュ値>.webp]]"
```

```shell
tkn-excel-note pull --source workbooks --dry-run
tkn-excel-note pull --source workbooks
```

- 対象は、`.xlsx` / `.xlsm` のファイル内に保存されたサムネイルです。シートに挿入した画像は使いません。
- Windows では、EMF / WMF / PNG / JPEG 形式のサムネイルに対応します。
- 変換した画像は、長辺が最大 1200 px の WebP（可逆圧縮） です。
- サムネイルがないブックでは、`cover` は空になります。
- 変換できないサムネイルは警告を表示し、既存の `cover` を残します。メタデータの同期は続けます。
- `--dry-run` は画像の変換まで検証しますが、画像・ノート・同期記録は保存しません。

### シートの指定範囲を画像にする

サムネイルがないブックや、特定の範囲を大きく表示したい場合は、`sheet` 方式を使います。
保存済みのシートを Excel で描画し、指定した範囲を1枚の WebP（可逆圧縮）にします。

```shell
tkn-excel-note pull --source workbooks --cover sheet --dry-run
tkn-excel-note pull --source workbooks --cover sheet
tkn-excel-note pull --source workbooks --cover sheet --cover-sheet "概要" --cover-range "B2:R51" --cover-width 3000
```

| オプション | 既定値 | 意味 |
| --- | --- | --- |
| `--cover sheet` / `embedded` / `auto` | `auto` | 画像の方式です。`auto` はブックごとに最後に成功した方式を使い、記録がなければ `embedded` を使います。 |
| `--cover-sheet` | 先頭の表示中のワークシート | 画像にするシートの名前です。非表示のシートとチャートシートは指定できません。 |
| `--cover-range` | `A1:Q50` | 画像にする範囲です。1つの長方形の範囲を `A1:Q50` の形式で指定します。シート名付きの参照、複数の範囲、行・列全体は指定できません。 |
| `--cover-width` | `2400` | 画像の幅（ピクセル）です。600～4000 の整数を指定します。高さは範囲の縦横比で決まります。 |

`--cover-sheet`、`--cover-range`、`--cover-width` は `sheet` 方式で使います。
AI の解析対象を選ぶ `--sheet` とは別のオプションです。
`--context` と同時に指定できます。

成功すると、方式・シート・範囲・幅とブックの内容のハッシュ値を、同期記録の `coverGeneration` にブックごとに記録します。
次回からは、オプションを省略した `pull` でも同じ条件で画像を作ります。
ブックと条件が変わらず画像も正常に残っていれば、Excel を起動せずに既存の画像を使います。
文書プロパティの変更や `push` の後も、ブックのファイルが変わるため、画像を作り直す場合があります。

`--dry-run` は、シートと範囲の指定が正しいか、画像を作り直す必要があるかを確認します。
Excel の起動、画像の作成、ノート・同期記録・レポートの保存は行いません。
結果 JSON の `details.cover` と標準エラー出力に、対象と条件、`planned` / `cached` / `rendered` / `warning` のいずれかの結果を表示します。
画像の高さと、実際に描画できるかは、実行するまで確定しません。

#### 描画の仕様

- ブックの一時コピーを Excel で開き、マクロ・イベント・外部リンクの更新・再計算を止めて描画します。元のブックと、利用者が開いている Excel は変更しません。
- 保存済みの印刷範囲・印刷タイトル・ヘッダー・フッターは使わず、指定範囲を PDF に出力してから画像化し、WebP に変換します。
- 範囲内の空白のセルは残し、用紙の余白だけを取り除きます。
- グリッド線と行・列の見出しは表示せず、罫線は表示します。
- 非表示の行・列や図形の印刷設定は Excel の印刷結果に従うため、画面のスクリーンショットとは異なる場合があります。
- 結合セルや図形が範囲の境界をまたぐ場合は、範囲を広げてください。

画像は最大 1600 万画素、描画する範囲は幅・高さとも 12,000 ポイントまでです。
超える場合は、範囲か画像の幅を小さくしてください。
入力ブックのサイズの上限には、`generation.max_workbook_mb`（既定 100 MiB）を使います。

設定ファイルで常に `sheet` 方式を使う方法は、「[cover の設定](../reference/configuration.md#cover-の設定)」を参照してください。

### 手動で設定した画像と失敗時の扱い

- 手動で設定した `cover` は、どちらの方式でも上書きしません。自動設定に戻すには、`cover` を空にします。
- 自動で設定したリンクは、同期記録の `managedCover` で識別します。同期記録を失った場合、空でない `cover` は手動の設定として保護します。
- 自動で作った画像が削除・破損していれば、次回の `pull` で作り直します。
- 画像を作れなかった場合（シートが存在しない、Excel がない、描画に失敗したなど）は、警告を表示し、以前の `cover` と記録を残して同期を続けます。
- 元のブックと、過去に作った画像は削除しません。

## 元ブックがない代理ノートを削除する

ブックを削除しても、代理ノートは残り、`pull` のたびに `missing-source` と報告されます。
`delete-notes` は、元ブックが存在しないことを確認したうえで、代理ノートと、そのノートの同期記録を削除します。
削除する前に、ノートと同期記録をバックアップします。

ノートを指定して削除する場合は、次のように実行します。

```shell
tkn-excel-note delete-notes --source workbooks --note "example.xlsx.md" --dry-run
tkn-excel-note delete-notes --source workbooks --note "example.xlsx.md"
```

元ブックがないノートをまとめて削除する場合は、`--all-missing` を使います。

```shell
tkn-excel-note delete-notes --source workbooks --all-missing --dry-run
tkn-excel-note delete-notes --source workbooks --all-missing
```

`--note` と `--all-missing` のどちらか一方を指定します。
`--note` は繰り返し指定でき、次のいずれかの形式で指定できます。

- ノートのファイル名、絶対パス、ノートの保存先からの相対パス。
- 元ブックの相対パス（`sourceFileName`）。
- `noteId` または `sourceId`。

指定が複数のノートに一致する場合は、削除しません。
相対パスか絶対パスで指定し直してください。

### 削除の条件

- 元ブックが存在するノートは削除しません。`--note` で指定した場合も削除せず、`--all-missing` の場合は対象から外します。
- 元ブックの存在は、記録したパス、ブックの固定 ID、内容の照合で確認します。ブックを移動・名前変更していても見落とさないよう、設定したすべての source の入力フォルダを、サブフォルダや除外したファイルも含めて調べます。
- 対象のノートは、source の `recursive`、`include`、`ignore` の設定に従って選びます。
- `--all-missing` では、Frontmatter の `sourceRoot` が別の source を示すノートを対象外（`skipped-note`）にします。
- 入力フォルダが見つからない、読み取れない、同期記録が欠けている・曖昧、同じブックのノートが重複しているなどの問題がある場合は、何も削除せずに停止します。

### バックアップと復元

削除する前に、対象のノートの全文と同期記録を `~/.tkn/excel_note/state/backups/deleted-notes/<run-id>/` に保存します。
Excel ブック、添付画像、説明の生成記録は削除しません。
ノートを削除した後に、そのノートの同期記録だけを削除するため、次回の `pull` で `missing-source` は表示されなくなります。
後でブックを元に戻した場合は、次回の `pull` で新しい代理ノートを作成します。

実行が終わると、結果にバックアップと実行レポートの保存先を表示します。
`--dry-run` では、バックアップとレポートも作りません。

削除の途中で書き込みに失敗した場合や、中断した場合は、削除したノートを自動で戻すことを試みます。
`delete-notes` の実行中は、ほかの同期コマンドを実行しないでください。

プロセスの強制終了などで自動では戻せなかった場合は、バックアップの `manifest.json` を使って手作業で戻します。

1. `manifest.json` で、元のノートのパス、バックアップのファイル名、同期記録のキーを確認します。
2. バックアップのノートを、元のパスにコピーします。
3. 必要なら、バックアップの `sync-state.json` から、該当するキーの項目だけを現在の `sync-state.json` に戻します。

同期記録のファイル全体をバックアップで上書きすると、削除後に行った同期の記録を失います。
該当する項目だけを戻してください。

## 旧 `.xls` を変換する

このツールは `.xls` を読み取れません。
同梱の PowerShell スクリプト [Convert-XlsToOpenXml.ps1](../../scripts/Convert-XlsToOpenXml.ps1) で、`.xlsx` または `.xlsm` に変換してから使います。
変換には Windows とデスクトップ版 Excel が必要です。

- VBA を含むブックは `.xlsm`、それ以外は `.xlsx` に変換します。
- 元の `.xls` は残します。
- 同じ名前の変換先が既にある場合は、上書きせずにスキップします。

PowerShell でリポジトリのフォルダに移動し、入力フォルダのパスを置き換えて実行します。

```powershell
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath "C:\path\to\legacy-workbooks" -DryRun
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath "C:\path\to\legacy-workbooks"
```

| パラメーター | 意味 |
| --- | --- |
| `-SourcePath` | `.xls` を探すフォルダです。省略すると、現在のフォルダを対象にします。 |
| `-OutputDirectory` | 変換先のフォルダです。省略すると、元の `.xls` と同じフォルダに保存します。`-Recurse` と併用すると、サブフォルダのブックもすべてこのフォルダに保存するため、同じ名前のブックは衝突として報告します。 |
| `-Recurse` | サブフォルダの `.xls` も対象にします。 |
| `-DryRun` | 変換せずに、変換先の形式と衝突を確認します。 |

`-DryRun` でも Excel を起動し、元のファイルを読み取り専用で開いて VBA の有無を確認します。
フォルダやブックは作成しません。
変換中は、マクロ、イベント、外部リンクの更新、警告の表示を止めます。
拡張子が `.xls` のファイルだけを対象にし、`.xlsx` / `.xlsm` は変更しません。
進捗は標準エラー出力に、最後の集計は標準出力に1件の JSON で出力します。

### 旧バージョンの cover 画像

次回の `pull` で、自動管理の cover は `thumbnails/` の WebP へのリンクに更新します。シート方式は同期記録と一致する旧画像を変換に再利用し、正本が変わっている場合は再描画します。`--dry-run` は画像・ノート・同期記録を書き込みません。手動指定の cover と旧 `img/` の画像は保持します。VLM 用の出典画像は引き続き `img/` の PNG を使います。

`thumbnails` 設定で最大辺・品質・形式などを変更できます。たとえば `max_edge: 480`、`format: webp`、`quality: 85`、`lossless: false` で小さなプレビューを生成できます。詳細は[設定リファレンス](../reference/configuration.md#プレビュー画像の設定)を参照してください。
