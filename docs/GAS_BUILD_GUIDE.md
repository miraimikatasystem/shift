# GAS Build Guide

このテンプレートを使って新しい GAS アプリを組み込むときの基本手順です。

## 目的

このテンプレートは、Next.js 側で Google ログインを受け、GAS 側で短期セッションを発行し、そのセッションを iframe に `POST` して起動する構成です。

重要なのは次の 2 点です。

1. GAS の初回表示 HTML は、認証済みセッションがあるときだけ返す
2. 初回表示後の公開関数も、毎回 `sessionToken` を検証する

## 最小構成

GAS 側には最低限、次の要素を用意します。

- `doGet(e)`
- `doPost(e)`
- `buildAppHtml_(sessionToken)`
- `createUnauthorizedHtml_()`
- `createSession_(email)`
- `getSessionFromToken_(token, touch)`
- `revokeSession_(token)`
- `requireSession_(sessionToken)`
- `verifyGoogleIdToken_(idToken)`
- `isAllowedEmail_(email)`

参考実装:
- [gas-sample/WebApp.gs](/C:/Users/Ayuki/Documents/template/gas-sample/WebApp.gs)

## 実装手順

1. まず `gas-sample/WebApp.gs` を新しい GAS プロジェクトへコピーする

2. `Index.html` を用意し、テンプレート変数 `sessionToken` を受け取れるようにする

3. `doPost(e)` に `bootstrapSessionToken` 分岐を入れる

```javascript
const bootstrapSessionToken = e && e.parameter ? String(e.parameter.bootstrapSessionToken || '') : '';
if (bootstrapSessionToken) {
  const session = getSessionFromToken_(bootstrapSessionToken, true);
  if (!session) return createUnauthorizedHtml_();
  return buildAppHtml_(bootstrapSessionToken);
}
```

4. API 用 `doPost(e)` では `idToken`, `action`, `payload` を受け、`createSession` と `revokeSession` を処理する

5. 公開関数はすべて `sessionToken` を最後の引数に追加する

例:

```javascript
function getInitialData(sessionToken) {
  requireSession_(sessionToken);
  return { ok: true };
}

function saveItem(data, sessionToken) {
  requireSession_(sessionToken);
  // 保存処理
}
```

6. 内部専用処理は `Internal_` 付きなどで分離する

例:

```javascript
function getInitialData(sessionToken) {
  requireSession_(sessionToken);
  return getInitialDataInternal_();
}

function getInitialDataInternal_() {
  // 認証前提の実処理
}
```

## 既存 GAS ファイルを載せ替えるときの考え方

既存の `.gs` / `.html` をこのテンプレートに載せる場合は、次の順で直すと安全です。

1. まず入口を作る
`WebApp.gs` でセッション発行、セッション検証、HTML返却を完成させる

2. 次に公開関数を洗い出す
`google.script.run` から呼ばれている関数を全部確認する

3. その関数に `sessionToken` を追加する
引数の最後に追加し、先頭で `requireSession_(sessionToken)` を呼ぶ

4. 重い処理は内部関数へ分ける
認証済み前提の処理を `...Internal_()` に逃がすと整理しやすい

5. HTML 側の呼び出しを合わせる
`google.script.run.someAction(data, window.__sessionToken)` の形に揃える

## Script Properties

最低限これを設定します。

- `GOOGLE_OAUTH_CLIENT_ID`
- `ALLOWED_EMAILS`

必要に応じて追加:

- `GEMINI_API_KEY`
- アプリ固有の設定値

## 避けるべき構成

- `sessionToken` を URL クエリに載せる
- `id_token` を localStorage に保存する
- `doGet()` だけ認証して、公開関数は素通しにする
- 更新系関数でサーバー側認可を省く

## 作業チェックリスト

- `createSession` が動く
- `revokeSession` が動く
- `bootstrapSessionToken` で iframe 起動できる
- 公開関数すべてが `requireSession_()` を通る
- セッション切れ時に再ログインへ戻せる
