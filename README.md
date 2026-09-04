# SharePoint 利用者アクセス集計

IE11 / Edge 95で動作する、外部ライブラリなしのSTI利用記録復号・集計画面です。

## 現行仕様

対象は `stiuseraccess` の **SchemaVersion=2** です。

- 1ユーザー = SharePoint 1項目
- 復号後 `eventType` = `question_start_summary`
- `questionStartCount` = 問題を開始した回数
- `lastQuestionStartUtc` = 最終問題開始日時（UTC）
- 旧 `SchemaVersion=1 / eventType=session_start` は現在集計に混ぜません

対象となるSTIモードは `fourCorrect`、`fourWrong`、`trueFalse` です。`browse` は個人開始回数に加算されません。

## 接続先

既定値:

- SharePoint Web Root: `/na/NA/NAFin/fin_csm`
- 個人別利用記録: `stiuseraccess`
- 全体ACCESSカウンター: `stiaccesscounter`（この解析画面の個人集計対象外）

## 暗号方式

暗号方式は AES-256-CBC です。

```text
AES_KEY = SHA256(UTF8(sharedSecret))
EncryptedPayload =
  "STI-AES-CBC-SHA256-v1:"
  + Base64(IV[16 bytes] || CipherText)
```

`SchemaVersion=2` はデータ構造のバージョンで、`STI-AES-CBC-SHA256-v1` は暗号方式のバージョンです。

現在の復号設定は `web/index.html` の `window.STI_DECRYPTION_CONFIG` にあります。STI側と同じ共有秘密を設定してください。

## 復号後JSON

```json
{
  "eventType": "question_start_summary",
  "userId": "123",
  "questionStartCount": 7,
  "lastQuestionStartUtc": "2026-09-04T03:50:00.000Z"
}
```

解析画面は次を検証します。

- `SchemaVersion == 2`
- `eventType == "question_start_summary"`
- `userId` が存在
- `questionStartCount` が0以上の整数
- `lastQuestionStartUtc` が有効な日時

氏名は復号後の `userId` と SharePoint `siteusers` を照合して表示します。画面上の最終問題開始日時はJSTに変換して表示し、内部データとJSON出力ではUTC値も保持します。

## SharePoint Title

STI側ではユーザーごとの固定キーを次の方式で作成します。

```text
source = "sti-user:" + sharedSecret + ":" + SharePointUserId
Title = "user-" + HEX(SHA256(UTF8(source)))
```

これはHMACではありません。

## 使い方

1. `web/index.html`、`web/app.js`、`web/styles.css` をSharePoint上の同一フォルダーへ配置します。
2. `web/index.html` の `STI_DECRYPTION_CONFIG.secretKey` がSTI側の共有秘密と一致していることを確認します。
3. SharePointへログイン済みのIE11 / Edge 95で `index.html` を開きます。
4. 「取得して集計」を押します。

SharePoint REST APIの次ページURLを追跡して全ページを取得します。復号失敗や不正データは項目単位でスキップし、処理全体は継続します。

## 出力

画面表示:

- 利用者ID
- 氏名
- 問題開始回数
- 最終問題開始（JST）

CSV / JSON出力では `lastQuestionStartUtc` とJST表示値の両方を出力します。

## 注意

共有秘密はブラウザー側コードに設定されるため、本格的な秘匿用途ではありません。SharePointリストを直接見ただけでは利用者ID等を読み取れないようにする用途を想定しています。
