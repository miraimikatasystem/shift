export type GasApiSuccess<T = unknown> = {
  status: number;
  ok: true;
  data: T;
  email?: string;
};

type GasApiEnvelope<T = unknown> = {
  status?: number;
  ok?: boolean;
  data?: T;
  message?: string;
  code?: string;
  email?: string;
};

export class GasApiError extends Error {
  status: number;
  code?: string;

  constructor(message: string, status: number, code?: string) {
    super(message);
    this.name = "GasApiError";
    this.status = status;
    this.code = code;
  }
}

function messageFromStatus(status: number): string {
  if (status === 401) return "認証に失敗しました。Googleで再ログインしてください。";
  if (status === 403) return "このアカウントにはアクセス権限がありません。";
  if (status === 0) return "ネットワークエラーが発生しました。接続を確認してください。";
  return "サーバーとの通信に失敗しました。";
}

export async function callGasApi<T = unknown>(
  idToken: string,
  action: string,
  payload?: Record<string, unknown>,
): Promise<GasApiSuccess<T>> {
  let response: Response;
  try {
    response = await fetch("/api/gas", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ idToken, action, payload: payload ?? {} }),
    });
  } catch {
    throw new GasApiError(messageFromStatus(0), 0, "network_error");
  }

  let body: GasApiEnvelope<T> = {};
  try {
    body = (await response.json()) as GasApiEnvelope<T>;
  } catch {
    body = {};
  }

  if (!response.ok || body.ok !== true) {
    const message = body.message || messageFromStatus(response.status);
    throw new GasApiError(message, response.status, body.code);
  }

  return {
    status: response.status,
    ok: true,
    data: (body.data as T) ?? ({} as T),
    email: body.email,
  };
}
