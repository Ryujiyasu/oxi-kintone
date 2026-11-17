# サンプル

- `見積書.docx`・`納品書.docx`（Word）、`請求書.xlsx`（Excel）: お試し用アプリのフィールドに合わせたひな形。プラグインの設定画面の「サンプルのひな形を追加」で入るものと同じです。
- `oxi-kintone-sample-app.zip`（[Releases](../../../releases) に添付）: お試し用アプリの kintone アプリテンプレート（フィールドと、Oxi 帳票プラグインの追加まで）。プラグインの設定とレコードはテンプレートに含まれません。
- `app-fields.json`: お試し用アプリのフィールド定義（REST API `/k/v1/preview/app/form/fields.json` の形式）。
- `make_samples.py`: ひな形を作り直すスクリプト（python-docx・openpyxl）。`請求書.xlsx` は、数式の計算結果を保存するため、作ったあとに一度 LibreOffice で保存し直しています（`soffice --headless --convert-to xlsx`）。
- `sample-record.mjs`・`try.mjs`: kintone の外でひな形を試すためのスクリプト（`node try.mjs <@oxidocs/wasm のフォルダ> <出力先>`）。
