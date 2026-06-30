const APP_TOKEN_SECRET = Deno.env.get("APP_PASSWORD") ?? "";
const encoder = new TextEncoder();
const decoder = new TextDecoder();

function base64UrlEncode(bytes: Uint8Array): string {
  const b64 = btoa(String.fromCharCode(...bytes));
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function base64UrlDecode(value: string): Uint8Array {
  const pad = "=".repeat((4 - (value.length % 4)) % 4);
  const b64 = value.replace(/-/g, "+").replace(/_/g, "/") + pad;
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function constantTimeEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i += 1) {
    result |= a[i] ^ b[i];
  }
  return result === 0;
}

async function hmac(message: string): Promise<Uint8Array> {
  if (!APP_TOKEN_SECRET) {
    throw new Error("Token secret is not configured");
  }

  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(APP_TOKEN_SECRET),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );

  const signature = await crypto.subtle.sign("HMAC", key, encoder.encode(message));
  return new Uint8Array(signature);
}

export async function signToken(payload: Record<string, unknown>): Promise<string> {
  const payloadJson = JSON.stringify(payload);
  const signature = await hmac(payloadJson);
  return `${base64UrlEncode(encoder.encode(payloadJson))}.${base64UrlEncode(signature)}`;
}

export async function verifyToken(token: string): Promise<Record<string, unknown> | null> {
  if (!token) return null;
  const [payloadPart, signaturePart] = token.split(".");
  if (!payloadPart || !signaturePart) return null;

  const payloadBytes = base64UrlDecode(payloadPart);
  const payloadJson = decoder.decode(payloadBytes);
  let payload: Record<string, unknown>;

  try {
    payload = JSON.parse(payloadJson);
  } catch {
    return null;
  }

  const expectedSignature = await hmac(payloadJson);
  const actualSignature = base64UrlDecode(signaturePart);
  if (!constantTimeEqual(expectedSignature, actualSignature)) return null;

  if (typeof payload.exp !== "number" || payload.exp < Date.now()) {
    return null;
  }

  return payload;
}
