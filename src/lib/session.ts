const SESSION_KEY = "rfp-session-id";

export function getSessionId(): string {
  if (typeof window === "undefined") return "";

  let sessionId = localStorage.getItem(SESSION_KEY);
  if (sessionId) return sessionId;

  sessionId = crypto.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  localStorage.setItem(SESSION_KEY, sessionId);
  return sessionId;
}
