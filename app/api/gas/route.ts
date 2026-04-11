import { NextResponse } from "next/server";

type GasRequestBody = {
  idToken?: string;
  action?: string;
  payload?: Record<string, unknown>;
};

type GasResponseBody = {
  status?: number;
  ok?: boolean;
  message?: string;
  code?: string;
  data?: unknown;
};

const gasWebAppUrl = process.env.NEXT_PUBLIC_GAS_WEBAPP_URL;

function asHttpStatus(raw: unknown, fallback: number): number {
  const n = Number(raw);
  if (Number.isInteger(n) && n >= 100 && n <= 599) return n;
  return fallback;
}

function summarizeUpstreamText(raw: string): string {
  return raw.replace(/\s+/g, " ").trim().slice(0, 240);
}

export async function POST(req: Request) {
  if (!gasWebAppUrl) {
    return NextResponse.json(
      { status: 500, ok: false, code: "missing_gas_url", message: "GAS URL is not configured" },
      { status: 500 },
    );
  }

  let body: GasRequestBody;
  try {
    body = (await req.json()) as GasRequestBody;
  } catch {
    return NextResponse.json(
      { status: 400, ok: false, code: "invalid_json", message: "invalid request body" },
      { status: 400 },
    );
  }

  let gasRes: Response;
  try {
    gasRes = await fetch(gasWebAppUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        idToken: body.idToken ?? "",
        action: body.action ?? "",
        payload: body.payload ?? {},
      }),
      cache: "no-store",
    });
  } catch {
    return NextResponse.json(
      { status: 502, ok: false, code: "upstream_unreachable", message: "failed to reach gas" },
      { status: 502 },
    );
  }

  let gasBody: GasResponseBody = {};
  let gasText = "";
  try {
    gasText = await gasRes.text();
    gasBody = gasText ? (JSON.parse(gasText) as GasResponseBody) : {};
  } catch {
    gasBody = {
      status: gasRes.status,
      ok: false,
      code: "invalid_upstream_response",
      message: gasText
        ? `invalid response from gas: ${summarizeUpstreamText(gasText)}`
        : "invalid response from gas",
    };
  }

  const httpStatus = asHttpStatus(gasBody.status, asHttpStatus(gasRes.status, 502));
  return NextResponse.json(gasBody, { status: httpStatus });
}
