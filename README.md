# SharePoint 利用者アクセス集計

IE11 / Edge 95で動作する、外部ライブラリなしのSharePoint集計画面です。

## 使い方

1. `web/index.html` をSharePointサイト（例: Site Assets）へ配置します。
2. `web/index.html` 内の `window.STI_PRIVATE_KEY_JWK = null;` に、RSA秘密鍵をJWK形式で設定します。
3. SharePointへログイン済みのIE11 / Edge 95で画面を開き、「取得して集計」を押します。

既定の接続先は `/na/NA/NAFin/fin_csm`、対象リストはその配下の `stiuseraccess` です。ホスト名を含むURLへ変更することもできます。

JWKの必須項目は次のとおりです。

```javascript
window.STI_PRIVATE_KEY_JWK = {
  kty: "RSA",
  n: "RSA modulus (Base64URL)",
  e: "AQAB",
  d: "private exponent (Base64URL)",
  p: "prime1 (Base64URL)",
  q: "prime2 (Base64URL)",
  dp: "exponent1 (Base64URL)",
  dq: "exponent2 (Base64URL)",
  qi: "coefficient (Base64URL)"
};
```

## 処理内容

SharePoint REST APIで `stiuseraccess` を全ページ取得し、項目IDの重複を除外します。`SchemaVersion=1`、復号後の `eventType=session_start` のみを `userId` ごとに集計し、`lastAccessUtc` の最大値を表示します。SharePointサイトユーザーのIDと照合して氏名を表示し、照合できない場合は `未照合` とします。

復号失敗、JSON不正、日時不正、必須値欠落はスキップし、項目IDと理由を解析ログへ記録します。CSVとJSONのダウンロードにも対応しています。

## 注意

- 秘密鍵はHTMLに含まれるため、そのHTMLを読める利用者には秘密鍵も見えます。これは今回の簡易構成の前提です。
- 暗号方式は `STI-RSA-OAEP-SHA1-v1`、RSA-OAEP、SHA-1です。
- SharePointのログインセッションを使うため、画面はSharePointと同一オリジンで配信してください。`file://` 直開きは避けてください。
