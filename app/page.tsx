"use client";

import Script from "next/script";
import { useCallback, useEffect, useRef, useState } from "react";
import { GasApiError, callGasApi } from "@/lib/gasApi";

type AuthState =
  | "checking"
  | "signed_out"
  | "authorizing"
  | "signed_in"
  | "denied"
  | "error";

type JwtPayload = {
  email?: string;
  name?: string;
  picture?: string;
};

type UserInfo = {
  email: string;
  name: string;
  picture?: string;
};

type GoogleCredentialResponse = {
  credential?: string;
};

type GoogleIdentity = {
  accounts: {
    id: {
      initialize: (options: {
        client_id: string;
        callback: (response: GoogleCredentialResponse) => void;
      }) => void;
      renderButton: (
        parent: HTMLElement,
        options: Record<string, string | number | boolean>,
      ) => void;
      prompt: () => void;
      cancel: () => void;
      revoke: (hint: string, callback: () => void) => void;
      disableAutoSelect: () => void;
    };
  };
};

const gasUrl = process.env.NEXT_PUBLIC_GAS_WEBAPP_URL;
const googleClientId = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID;
const googleAuthEnabled = process.env.NEXT_PUBLIC_ENABLE_GOOGLE_AUTH !== "false";
const directTestBypassEnabled =
  process.env.NEXT_PUBLIC_ENABLE_DIRECT_TEST_BYPASS === "true";
const googleAuthUiEnabled = googleAuthEnabled && !directTestBypassEnabled;
const directTestBypassToken =
  process.env.NEXT_PUBLIC_DIRECT_TEST_BYPASS_TOKEN ||
  "cafe-shift-direct-8f3c27d4a91b";
const directTestUser: UserInfo = {
  email: "direct-test@local",
  name: "Direct Test",
};

function buildGasIframeUrl(
  baseUrl: string,
  view: "staff" | "admin",
  directTestToken: string,
  nonce: number,
): string {
  const url = new URL(baseUrl);
  url.searchParams.set("directTest", directTestToken);
  url.searchParams.set("embedded", "1");
  url.searchParams.set("_ts", String(nonce));
  if (view === "admin") url.searchParams.set("view", "admin");
  else url.searchParams.delete("view");
  return url.toString();
}

function base64UrlDecode(value: string): string {
  const padded = value + "=".repeat((4 - (value.length % 4)) % 4);
  const base64 = padded.replace(/-/g, "+").replace(/_/g, "/");
  return atob(base64);
}

function parseJwt(credential: string): JwtPayload | null {
  const sections = credential.split(".");
  if (sections.length < 2) return null;

  try {
    return JSON.parse(base64UrlDecode(sections[1])) as JwtPayload;
  } catch {
    return null;
  }
}

function toUserMessage(error: unknown): string {
  if (!(error instanceof GasApiError)) {
    return "認証処理に失敗しました。時間をおいて再実行してください。";
  }

  if (error.status === 401) {
    return error.message || "認証に失敗しました。Googleで再ログインしてください。";
  }
  if (error.status === 403) {
    return error.message || "このGoogleアカウントにはアクセス権限がありません。";
  }
  if (error.status === 0) {
    return "ネットワークエラーが発生しました。接続を確認してください。";
  }

  return error.message || "サーバーとの通信に失敗しました。";
}

export default function Page() {
  const [authState, setAuthState] = useState<AuthState>(
    directTestBypassEnabled ? "signed_in" : googleAuthUiEnabled ? "checking" : "signed_out",
  );
  const [user, setUser] = useState<UserInfo | null>(
    directTestBypassEnabled ? directTestUser : null,
  );
  const [scriptReady, setScriptReady] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [sessionToken, setSessionToken] = useState("");
  const [idToken, setIdToken] = useState("");
  const [iframeView, setIframeView] = useState<"staff" | "admin">("staff");
  const [directTestNonce, setDirectTestNonce] = useState(0);
  const buttonRef = useRef<HTMLDivElement | null>(null);
  const iframeRef = useRef<HTMLIFrameElement | null>(null);
  const bootstrapFormRef = useRef<HTMLFormElement | null>(null);
  const submittedSessionRef = useRef("");
  const directTestIframeUrl =
    gasUrl && directTestBypassEnabled && !sessionToken
      ? buildGasIframeUrl(gasUrl, iframeView, directTestBypassToken, directTestNonce)
      : "about:blank";

  const handleCredential = useCallback(async (response: GoogleCredentialResponse) => {
    if (!response.credential) {
      setAuthState("error");
      setErrorMessage("Googleからトークンを受け取れませんでした。");
      return;
    }

    const payload = parseJwt(response.credential);
    if (!payload?.email) {
      setAuthState("error");
      setErrorMessage("Googleアカウント情報の読み取りに失敗しました。");
      return;
    }

    setAuthState("authorizing");
    setErrorMessage("");
    setIdToken(response.credential);

    try {
      const gasResponse = await callGasApi<{
        sessionToken: string;
      }>(response.credential, "createSession", {});
      setUser({
        email: payload.email,
        name: payload.name ?? payload.email,
        picture: payload.picture,
      });
      setSessionToken(gasResponse.data.sessionToken);
      setAuthState("signed_in");
    } catch (error) {
      setUser(null);
      setIdToken("");
      setSessionToken("");
      setAuthState(error instanceof GasApiError && error.status === 403 ? "denied" : "error");
      setErrorMessage(toUserMessage(error));
    }
  }, []);

  useEffect(() => {
    const onMessage = (event: MessageEvent) => {
      if (event.data?.type === "gas-switch-admin" && event.data?.sessionToken) {
        submittedSessionRef.current = "";
        setIframeView("admin");
        setSessionToken(String(event.data.sessionToken));
        return;
      }
      if (event.data?.type !== "gas-session-expired") return;

      submittedSessionRef.current = "";
      if (directTestBypassEnabled) {
        setUser(directTestUser);
        setIdToken("");
        setSessionToken("");
        setIframeView("staff");
        setAuthState("signed_in");
        setErrorMessage("テスト用セッションを再発行しました。");
        setDirectTestNonce((prev) => prev + 1);
        return;
      }
      setUser(null);
      setIdToken("");
      setSessionToken("");
      setIframeView("staff");
      setAuthState("signed_out");
      setErrorMessage("セッションの有効期限が切れました。Googleで再ログインしてください。");
    };

    window.addEventListener("message", onMessage);
    return () => window.removeEventListener("message", onMessage);
  }, []);

  useEffect(() => {
    if (authState !== "signed_in" || !sessionToken || !gasUrl || !bootstrapFormRef.current) return;
    const nextSubmitKey = `${iframeView}:${sessionToken}`;
    if (submittedSessionRef.current === nextSubmitKey) return;

    submittedSessionRef.current = nextSubmitKey;
    bootstrapFormRef.current.submit();
  }, [authState, sessionToken, iframeView]);

  useEffect(() => {
    if (!googleAuthUiEnabled) return;
    if (!googleClientId) {
      setAuthState("error");
      setErrorMessage("NEXT_PUBLIC_GOOGLE_CLIENT_ID が未設定です。");
      return;
    }

    if (!scriptReady || authState === "signed_in" || authState === "authorizing" || !buttonRef.current) {
      return;
    }

    const google = (window as Window & { google?: GoogleIdentity }).google;
    if (!google) return;

    buttonRef.current.innerHTML = "";
    google.accounts.id.initialize({
      client_id: googleClientId,
      callback: (credentialResponse) => {
        void handleCredential(credentialResponse);
      },
    });
    google.accounts.id.renderButton(buttonRef.current, {
      theme: "outline",
      size: "large",
      shape: "pill",
      width: 280,
      text: "signin_with",
    });
    google.accounts.id.prompt();
    if (authState === "checking") setAuthState("signed_out");

    return () => {
      google.accounts.id.cancel();
    };
  }, [authState, handleCredential, scriptReady]);

  const handleSignOut = () => {
    const google = (window as Window & { google?: GoogleIdentity }).google;
    if (googleAuthUiEnabled && google && user?.email) {
      google.accounts.id.revoke(user.email, () => {});
    }
    if (googleAuthUiEnabled && google) google.accounts.id.disableAutoSelect();

    if (sessionToken) {
      void callGasApi(idToken, "revokeSession", { sessionToken }).catch(() => undefined);
    }

    submittedSessionRef.current = "";
    if (directTestBypassEnabled) {
      setUser(directTestUser);
      setIdToken("");
      setSessionToken("");
      setIframeView("staff");
      setErrorMessage("");
      setAuthState("signed_in");
      setDirectTestNonce((prev) => prev + 1);
      return;
    }
    setUser(null);
    setIdToken("");
    setSessionToken("");
    setIframeView("staff");
    setErrorMessage("");
    setAuthState("signed_out");
  };

  return (
    <main className="page">
      {googleAuthUiEnabled ? (
        <Script
          src="https://accounts.google.com/gsi/client"
          strategy="afterInteractive"
          onLoad={() => setScriptReady(true)}
        />
      ) : null}

      {!gasUrl ? (
        <section className="notice">
          <h2>GAS URL が未設定です</h2>
          <p>
            <code>NEXT_PUBLIC_GAS_WEBAPP_URL</code> にデプロイ済み Web アプリ URL を設定してください。
          </p>
        </section>
      ) : googleAuthUiEnabled && !googleClientId ? (
        <section className="notice">
          <h2>Google Client ID が未設定です</h2>
          <p>
            <code>NEXT_PUBLIC_GOOGLE_CLIENT_ID</code> を設定してください。
          </p>
        </section>
      ) : !googleAuthEnabled && !directTestBypassEnabled ? (
        <section className="notice">
          <h2>ログイン画面を停止中です</h2>
          <p>この環境では Google 認証 UI を表示していません。</p>
        </section>
      ) : authState !== "signed_in" ? (
        <section className="notice">
          <h2>Google ログインが必要です</h2>
          <div ref={buttonRef} className="googleButton" />
          {authState === "authorizing" ? <p className="noticeText">認証中...</p> : null}
          {authState === "denied" || authState === "error" ? (
            <p className="noticeError">{errorMessage}</p>
          ) : null}
        </section>
      ) : (
        <section className="appShell">
          {directTestBypassEnabled ? (
            <section className="notice" style={{ marginBottom: 12 }}>
              <h2>テストモード</h2>
              <p>Google認証をバイパスして GAS を直接表示しています。</p>
            </section>
          ) : null}
          <form ref={bootstrapFormRef} action={gasUrl} method="post" target="gas-app-frame" hidden>
            <input type="hidden" name="bootstrapSessionToken" value={sessionToken} />
            <input type="hidden" name="view" value={iframeView} />
          </form>
          <button type="button" onClick={handleSignOut} className="floatingSignOut">
            {directTestBypassEnabled ? "Reload Test" : "Sign out"}
          </button>
          <iframe
            ref={iframeRef}
            name="gas-app-frame"
            src={directTestIframeUrl}
            title="GAS Web App"
            className="frame frameFullscreen"
            loading="lazy"
            allow="clipboard-read; clipboard-write"
            referrerPolicy="no-referrer"
          />
        </section>
      )}
    </main>
  );
}
