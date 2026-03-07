export type Suggestions = {
  refHospitals: string[];
  doctors: string[];
  refHospitalEmails?: Record<string, string>;
};

const BASE_URL =
  (import.meta as any).env?.VITE_SERVER_BASE_URL?.trim() || "http://localhost:8787";

async function requestJson<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE_URL}${path}`, {
    headers: { "Content-Type": "application/json", ...(init?.headers || {}) },
    ...init,
  });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  return (await res.json()) as T;
}

export async function fetchSuggestions(): Promise<Suggestions> {
  return requestJson("/api/suggestions");
}

export async function addRefHospital(name: string): Promise<void> {
  if (!name.trim()) return;
  await requestJson("/api/suggestions/ref-hospital", {
    method: "POST",
    body: JSON.stringify({ name }),
  });
}