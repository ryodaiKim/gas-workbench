# Issue: OCR が Google ドキュメント MIME で失敗する

## 概要
PDF取り込みパイプライン実行時に、以下エラーで処理が失敗する。

`drive.files.insert の呼び出しに失敗しました: OCR is not supported for files of type application/vnd.google-apps.document`

同エラーが複数回再現しており、OCR対象判定または対象ファイル解決（ショートカット/実体）に不整合がある可能性が高い。

## 影響
- `受付情報一覧` への登録が停止
- `ログ` に FAIL が継続記録される
- 定期実行でも同様に失敗し、運用上の欠損が発生

## 発生ログ（例）
- 日時: `2026/03/26 19:04:24`
- fileId: `1JgesWLOfmCK_BQmnOTyvRYwNtmMAIC2i`
- fileName: `KUH-0053.pdf`
- result: `FAIL`
- message: `次のエラーが発生し、drive.files.insert の呼び出しに失敗しました: OCR is not supported for files of type application/vnd.google-apps.document`

## 想定される原因
1. OCR呼び出し前の MIME 判定が想定外ケースを取りこぼしている
   - ファイル名が `.pdf` でも実MIMEが Google Docs のケース
   - ショートカット解決後の実体MIMEが Google Docs のケース
2. 実行中コードが最新デプロイでない
   - ローカル修正済みでも、Apps Script 側に反映されていない
3. `DriveApp.getFileById(...).getBlob()` の contentType が `application/vnd.google-apps.document` になる経路がある
   - この場合 `Drive.Files.insert(..., {ocr:true})` は仕様上失敗する

## 再現手順
1. `設定` シート `読み取りPDF格納先` に対象フォルダURLを設定
2. メニュー `受付自動化 > PDF手動読み取り＆表更新` を実行
3. `ログ` シートで FAIL を確認
4. 同一メッセージ（Google Docs MIME の OCR 非対応）が出ることを確認

## 期待動作
- OCRは **PDF実体** にのみ適用される
- Google Docs 実体は OCR を使わず `DocumentApp.openById(...).getBody().getText()` で直接抽出される
- 非対応形式は OCR 実行前に明示エラーとしてログ記録される

## 恒久対応タスク（次スプリント）
1. **前段診断ログを追加**
   - source fileId / sourceName / sourceMime
   - resolved fileId / resolvedName / resolvedMime（shortcut解決後）
   - blob contentType（OCR直前）
2. **OCR呼び出しガードを厳格化**
   - `mime === application/pdf` かつ `blob contentType === application/pdf` のときのみ OCR 実行
   - 条件不一致なら OCR 呼び出しを禁止し、FAIL理由を明示
3. **処理対象フィルタの定義を固定**
   - フォルダ走査時の対象選定条件を仕様化（PDF / Google Docs / shortcut）
4. **回帰テスト追加（疑似データ）**
   - 実体PDF
   - 実体Google Docs（`.pdf`風ファイル名含む）
   - shortcut -> PDF
   - shortcut -> Google Docs
   - 非対応MIME

## 受け入れ条件 (Acceptance Criteria)
- [ ] 同一 fileId で再実行しても `application/vnd.google-apps.document` に対する OCR 呼び出しが発生しない
- [ ] FAIL 時のログに `sourceMime / resolvedMime / blobMime` が記録される
- [ ] PDF 実体は従来通り OCR でテキスト抽出できる
- [ ] Google Docs 実体は OCR を使わず抽出できる

## 暫定回避策（運用）
1. 対象フォルダ内で失敗ファイルを確認
2. Google Docs 実体の場合:
   - そのまま直接読取り経路に通ることを確認（必要なら一時的に対象から除外）
3. ショートカットの場合:
   - 実体が PDF か Google Docs かを Drive UI で確認
4. スクリプト反映確認:
   - `npm --workspace projects/sample-registration-pipeline run push` 実行後にシート再読込

## 参考（関連実装箇所）
- `projects/sample-registration-pipeline/src/main.ts`
  - `extractIdFromDriveUrl`
  - `listTargetPdfs`
  - `resolveShortcutTarget`
  - `extractTextForSupportedFile`
  - `extractTextFromPdfViaOcr`
  - `processUnreadPdfs`
