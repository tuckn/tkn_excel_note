# CLI出力契約

通常の標準出力は `key=value`、`--json` は1行の妥当なJSONです。処理と終了コードは同一で、進捗・警告・診断は標準エラーです。help/versionは通常の案内を維持します。JSON文字列の制御文字はJSON規則でエスケープし、ログやANSI色は混ぜません。

`--json` / `--full-output` はroot・中間・末端コマンドで指定でき、子コマンドで指定が消えません。`--` の後ろは位置引数として扱います。quiet/verboseは結果を抑制せず、パイプ・リダイレクト・TTYで形式を切り替えません。

通常一覧は入れ子を `.`、配列を `[0]`、空の配列・mappingを `[]` / `{}` とします。文字列はJSON引用符で囲まず、パスはそのままコピーできます。改行・タブ・ANSIを含む制御文字をエスケープします。

| 結果 | 主な項目 |
| --- | --- |
| 同期・status | `status`, `command`, `changed`, `conflicts`, `errors`, `writeEnabled`, `statusCounts`, `reportPath`, `differencesPath`, `backupPath` と対象別 `results`。context実行時は`usage`。 |
| 対象別同期結果 | `status`, `sourceRoot`, `sourcePath`, `notePath`, `workbookPath`, `message`, `conflictFields`, `fieldDifferences`, `details`。sourcePathは従来の相対パス、workbookPathは入力ルートを補った絶対パス。 |
| export | `command`, `status`, `dryRun`, `files`, `exported`, `planned`, `errors`, `usage`, `results`。各結果の `inputPath`, `outputPath`, `status`, `message` 。 |
| 設定操作 | initは `status`, `command`, `configPath`。listは既存の `config`, `loaded_sources`, `winning_sources` と選択profile・schema情報。 |
| workbook list-sheets | `command`, `status`, `sheets`。 |
| エラー | `status`, `command`, `message` と分かる範囲の `configPath`, `inputPath`, `workbookPath`, `sourceId`, `targetPath`。パスを文章から抽出する必要はありません。 |

`results` がある結果では `resultCount` は省略前の全件数、`omittedResults` は省略件数です。既定で省略するのは理由・警告・競合を伴わない `unchanged`, `tracked`, `skipped`, `skip` のみです。失敗・要確認・変更対象は省略しません。`--full-output` は両形式で全件を出します。省略時の `detailsHint` は再実行時の指定方法、`detailsPath` は既存レポートの詳細JSONです。dry-runは保存せず、通常実行へ切り替えるよう求めません。

終了コードは0=成功、1=処理失敗・一部失敗、2=同期競合または引数エラー、3=設定・state・contextエラーです。要確認状態でも0となる場合があるためstatusCountsを確認します。中断は130で、完了済み成果物を自動削除しません。ノートYAML正規化用CLIは成功0・処理失敗1・引数エラー2・中断130を使用します。PowerShell変換補助は `-Json` / `-FullOutput` / `-Quiet` を使用し、失敗は1です。PowerShell自体によるパラメーター構文エラーはホストの診断となる場合があります。

JSONを解析する呼び出しでは、たとえば `tkn-excel-note --json pull --source example --dry-run` の標準出力をJSONとして読み、終了コードも確認します。PowerShell変換補助のJSON指定は `-Json` です。表示形式の指定は、処理内容、副作用、dry-runの境界、Python APIの戻り値、保存するJSON・manifest・ノート・レポートに影響しません。
