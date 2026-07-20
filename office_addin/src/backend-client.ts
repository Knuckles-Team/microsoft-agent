import {
  MAX_BACKEND_REQUEST_BYTES,
  MAX_BACKEND_RESPONSE_BYTES,
  parseRuntimeConfig,
  requireAllowedBackendUrl,
  requireEndpointPath,
  type RuntimeConfig,
  validateBearerToken,
} from "./validation";

export type FetchImplementation = (input: string, init?: RequestInit) => Promise<Response>;

export class BackendError extends Error {
  readonly code: "configuration" | "network" | "timeout" | "response";
  readonly status: number | undefined;

  constructor(
    code: BackendError["code"],
    message: string,
    status?: number,
  ) {
    super(message);
    this.name = "BackendError";
    this.code = code;
    this.status = status;
  }
}

interface RequestOptions {
  readonly method?: "GET" | "POST";
  readonly body?: unknown;
  readonly token?: string;
  readonly signal?: AbortSignal;
}

export class BackendClient {
  readonly baseUrl: string;
  readonly timeoutMs: number;
  private readonly fetchImplementation: FetchImplementation;

  constructor(
    baseUrl: string,
    allowedOrigins: readonly string[],
    timeoutMs: number,
    fetchImplementation: FetchImplementation = globalThis.fetch.bind(globalThis),
  ) {
    this.baseUrl = requireAllowedBackendUrl(baseUrl, allowedOrigins);
    if (!Number.isSafeInteger(timeoutMs) || timeoutMs < 1_000 || timeoutMs > 60_000) {
      throw new BackendError("configuration", "Backend timeout must be between 1 and 60 seconds.");
    }
    this.timeoutMs = timeoutMs;
    this.fetchImplementation = fetchImplementation;
  }

  async requestJson(endpoint: string, options: RequestOptions = {}): Promise<unknown> {
    const path = requireEndpointPath(endpoint);
    const token = validateBearerToken(options.token);
    const method = options.method ?? "GET";
    const headers: Record<string, string> = {
      Accept: "application/json",
    };
    if (token !== undefined) {
      headers.Authorization = `Bearer ${token}`;
    }

    let body: string | undefined;
    if (options.body !== undefined) {
      try {
        body = JSON.stringify(options.body);
      } catch {
        throw new BackendError("configuration", "Backend request body must be JSON-serializable.");
      }
      if (body === undefined || new TextEncoder().encode(body).byteLength > MAX_BACKEND_REQUEST_BYTES) {
        throw new BackendError("configuration", "Backend request body is too large.");
      }
      headers["Content-Type"] = "application/json";
    }

    const controller = new AbortController();
    const abortFromCaller = (): void => controller.abort(options.signal?.reason);
    if (options.signal?.aborted) {
      abortFromCaller();
    } else {
      options.signal?.addEventListener("abort", abortFromCaller, { once: true });
    }
    const timeoutHandle = globalThis.setTimeout(() => controller.abort("timeout"), this.timeoutMs);

    const init: RequestInit = {
      method,
      headers,
      signal: controller.signal,
      credentials: "omit",
      redirect: "error",
      referrerPolicy: "no-referrer",
      cache: "no-store",
    };
    if (body !== undefined) {
      init.body = body;
    }

    try {
      const response = await this.fetchImplementation(`${this.baseUrl}${path}`, init);
      if (!response.ok) {
        throw new BackendError("response", `Backend returned HTTP ${response.status}.`, response.status);
      }
      if (response.status === 204) {
        return null;
      }

      const declaredLength = Number(response.headers.get("content-length"));
      if (Number.isFinite(declaredLength) && declaredLength > MAX_BACKEND_RESPONSE_BYTES) {
        throw new BackendError("response", "Backend response is too large.", response.status);
      }
      const contentType = response.headers.get("content-type")?.toLowerCase() ?? "";
      if (!contentType.includes("application/json") && !contentType.includes("+json")) {
        throw new BackendError("response", "Backend response must use a JSON content type.", response.status);
      }
      const responseText = await response.text();
      if (new TextEncoder().encode(responseText).byteLength > MAX_BACKEND_RESPONSE_BYTES) {
        throw new BackendError("response", "Backend response is too large.", response.status);
      }
      try {
        return JSON.parse(responseText) as unknown;
      } catch {
        throw new BackendError("response", "Backend returned invalid JSON.", response.status);
      }
    } catch (error: unknown) {
      if (error instanceof BackendError) {
        throw error;
      }
      if (controller.signal.aborted) {
        throw new BackendError(
          "timeout",
          options.signal?.aborted ? "Backend request was cancelled." : "Backend request timed out.",
        );
      }
      throw new BackendError("network", "Backend could not be reached over the approved HTTPS origin.");
    } finally {
      globalThis.clearTimeout(timeoutHandle);
      options.signal?.removeEventListener("abort", abortFromCaller);
    }
  }
}

export async function loadRuntimeConfig(
  fetchImplementation: FetchImplementation = globalThis.fetch.bind(globalThis),
  configUrl = "/config.json",
): Promise<RuntimeConfig> {
  const controller = new AbortController();
  const timeoutHandle = globalThis.setTimeout(() => controller.abort("timeout"), 5_000);
  try {
    const response = await fetchImplementation(configUrl, {
      method: "GET",
      headers: { Accept: "application/json" },
      signal: controller.signal,
      credentials: "same-origin",
      redirect: "error",
      cache: "no-store",
    });
    if (!response.ok) {
      throw new BackendError("configuration", `Configuration request returned HTTP ${response.status}.`);
    }
    const text = await response.text();
    if (new TextEncoder().encode(text).byteLength > 65_536) {
      throw new BackendError("configuration", "Runtime configuration is too large.");
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(text) as unknown;
    } catch {
      throw new BackendError("configuration", "Runtime configuration is not valid JSON.");
    }
    return parseRuntimeConfig(parsed);
  } catch (error: unknown) {
    if (error instanceof BackendError) {
      throw error;
    }
    if (controller.signal.aborted) {
      throw new BackendError("timeout", "Loading runtime configuration timed out.");
    }
    throw new BackendError("configuration", "Runtime configuration could not be loaded.");
  } finally {
    globalThis.clearTimeout(timeoutHandle);
  }
}
