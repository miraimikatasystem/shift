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
    return "認証に失敗しました。Googleで再ログインしてください。";
  }
  if (error.status === 403) {
    return "このGoogleアカウントにはアクセス権限がありません。";
  }
  if (error.status === 0) {
    return "ネットワークエラーが発生しました。接続を確認してください。";
  }

  return error.message || "サーバーとの通信に失敗しました。";
}

export default function Page() {
  const [authState, setAuthState] = useState<AuthState>("checking");
  const [user, setUser] = useState<UserInfo | null>(null);
  const [scriptReady, setScriptReady] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [sessionToken, setSessionToken] = useState("");
  const [idToken, setIdToken] = useState("");
  const buttonRef = useRef<HTMLDivElement | null>(null);
  const iframeRef = useRef<HTMLIFrameElement | null>(null);
  const bootstrapFormRef = useRef<HTMLFormElement | null>(null);
  const submittedSessionRef = useRef("");

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
      if (event.source !== iframeRef.current?.contentWindow) return;
      if (event.data?.type !== "gas-session-expired") return;

      submittedSessionRef.current = "";
      setUser(null);
      setIdToken("");
      setSessionToken("");
      setAuthState("signed_out");
      setErrorMessage("セッションの有効期限が切れました。Googleで再ログインしてください。");
    };

    window.addEventListener("message", onMessage);
    return () => window.removeEventListener("message", onMessage);
  }, []);

  useEffect(() => {
    if (authState !== "signed_in" || !sessionToken || !gasUrl || !bootstrapFormRef.current) return;
    if (submittedSessionRef.current === sessionToken) return;

    submittedSessionRef.current = sessionToken;
    bootstrapFormRef.current.submit();
  }, [authState, sessionToken]);

  useEffect(() => {
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
    if (google && user?.email) google.accounts.id.revoke(user.email, () => {});
    if (google) google.accounts.id.disableAutoSelect();

    if (idToken && sessionToken) {
      void callGasApi(idToken, "revokeSession", { sessionToken }).catch(() => undefined);
    }

    submittedSessionRef.current = "";
    setUser(null);
    setIdToken("");
    setSessionToken("");
    setErrorMessage("");
    setAuthState("signed_out");
  };

  return (
    <main className="page">
      <Script
        src="https://accounts.google.com/gsi/client"
        strategy="afterInteractive"
        onLoad={() => setScriptReady(true)}
      />

      {!gasUrl ? (
        <section className="notice">
          <h2>GAS URL が未設定です</h2>
          <p>
            <code>NEXT_PUBLIC_GAS_WEBAPP_URL</code> にデプロイ済み Web アプリ URL を設定してください。
          </p>
        </section>
      ) : !googleClientId ? (
        <section className="notice">
          <h2>Google Client ID が未設定です</h2>
          <p>
            <code>NEXT_PUBLIC_GOOGLE_CLIENT_ID</code> を設定してください。
          </p>
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
          <form ref={bootstrapFormRef} action={gasUrl} method="post" target="gas-app-frame" hidden>
            <input type="hidden" name="bootstrapSessionToken" value={sessionToken} />
          </form>
          <button type="button" onClick={handleSignOut} className="floatingSignOut">
            Sign out
          </button>
          <iframe
            ref={iframeRef}
            name="gas-app-frame"
            src="about:blank"
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
