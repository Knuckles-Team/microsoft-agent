import { describe, expect, it, vi } from "vitest";

import { BackendClient, loadRuntimeConfig, type FetchImplementation } from "../src/backend-client";

const allowedOrigins = ["https://api.example.test"];

describe("BackendClient", () => {
  it("sends a bearer token only to the exact approved origin", async () => {
    const fetchImplementation = vi.fn<FetchImplementation>(async () =>
      new Response(JSON.stringify({ status: "ok" }), {
        status: 200,
        headers: { "content-type": "application/json" },
      }),
    );
    const client = new BackendClient("https://api.example.test", allowedOrigins, 5_000, fetchImplementation);

    await expect(client.requestJson("/health", { token: "session-token" })).resolves.toEqual({ status: "ok" });
    expect(fetchImplementation).toHaveBeenCalledOnce();
    const [url, init] = fetchImplementation.mock.calls[0] ?? [];
    expect(url).toBe("https://api.example.test/health");
    expect(init?.credentials).toBe("omit");
    expect(init?.redirect).toBe("error");
    expect((init?.headers as Record<string, string>).Authorization).toBe("Bearer session-token");
  });

  it("rejects non-JSON backend responses", async () => {
    const fetchImplementation: FetchImplementation = async () =>
      new Response("healthy", {
        status: 200,
        headers: { "content-type": "text/plain" },
      });
    const client = new BackendClient("https://api.example.test", allowedOrigins, 5_000, fetchImplementation);

    await expect(client.requestJson("/health")).rejects.toMatchObject({
      code: "response",
    });
  });

  it("does not expose an error response body", async () => {
    const fetchImplementation: FetchImplementation = async () =>
      new Response(JSON.stringify({ secret: "do not echo" }), {
        status: 401,
        headers: { "content-type": "application/json" },
      });
    const client = new BackendClient("https://api.example.test", allowedOrigins, 5_000, fetchImplementation);

    await expect(client.requestJson("/health")).rejects.toThrow("Backend returned HTTP 401.");
    await expect(client.requestJson("/health")).rejects.not.toThrow("do not echo");
  });

  it("blocks endpoint traversal before fetch", async () => {
    const fetchImplementation = vi.fn<FetchImplementation>();
    const client = new BackendClient("https://api.example.test", allowedOrigins, 5_000, fetchImplementation);

    await expect(client.requestJson("//attacker.test/collect", { token: "valuable-token" })).rejects.toThrow();
    expect(fetchImplementation).not.toHaveBeenCalled();
  });

  it("aborts a backend request when its timeout expires", async () => {
    vi.useFakeTimers();
    try {
      const fetchImplementation: FetchImplementation = async (_url, init) =>
        new Promise((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () => reject(new DOMException("Aborted", "AbortError")), {
            once: true,
          });
        });
      const client = new BackendClient("https://api.example.test", allowedOrigins, 1_000, fetchImplementation);
      const expectation = expect(client.requestJson("/health")).rejects.toMatchObject({ code: "timeout" });

      await vi.advanceTimersByTimeAsync(1_000);
      await expectation;
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("loadRuntimeConfig", () => {
  it("loads and validates same-origin JSON configuration", async () => {
    const fetchImplementation = vi.fn<FetchImplementation>(async () =>
      new Response(
        JSON.stringify({
          allowedBackendOrigins: ["https://api.example.test"],
          defaultBackendUrl: "https://api.example.test",
          healthPath: "/health",
          requestTimeoutMs: 10_000,
        }),
        { status: 200, headers: { "content-type": "application/json" } },
      ),
    );

    await expect(loadRuntimeConfig(fetchImplementation, "/config.json")).resolves.toMatchObject({
      defaultBackendUrl: "https://api.example.test",
    });
    const [, init] = fetchImplementation.mock.calls[0] ?? [];
    expect(init?.credentials).toBe("same-origin");
    expect(init?.redirect).toBe("error");
  });
});
