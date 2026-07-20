export const MAX_DOCUMENT_TEXT_LENGTH = 100_000;
export const MAX_PLACEHOLDER_LENGTH = 500;
export const MAX_TOKEN_LENGTH = 8_192;
export const MAX_BACKEND_RESPONSE_BYTES = 1_048_576;
export const MAX_BACKEND_REQUEST_BYTES = 262_144;

export class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ValidationError";
  }
}

export interface RuntimeConfig {
  readonly allowedBackendOrigins: readonly string[];
  readonly defaultBackendUrl: string;
  readonly healthPath: string;
  readonly requestTimeoutMs: number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function boundedText(
  value: unknown,
  fieldName: string,
  maximumLength = MAX_DOCUMENT_TEXT_LENGTH,
): string {
  if (typeof value !== "string") {
    throw new ValidationError(`${fieldName} must be text.`);
  }
  if (value.length > maximumLength) {
    throw new ValidationError(`${fieldName} cannot exceed ${maximumLength.toLocaleString()} characters.`);
  }
  return value;
}

export function requireText(
  value: unknown,
  fieldName: string,
  maximumLength = MAX_DOCUMENT_TEXT_LENGTH,
): string {
  const text = boundedText(value, fieldName, maximumLength);
  if (text.length === 0) {
    throw new ValidationError(`${fieldName} cannot be empty.`);
  }
  return text;
}

export function requirePlaceholder(value: unknown): string {
  const placeholder = requireText(value, "Placeholder", MAX_PLACEHOLDER_LENGTH);
  if (/\r|\n/.test(placeholder)) {
    throw new ValidationError("Placeholder must be a single line.");
  }
  return placeholder;
}

export function requireSlideIndex(value: unknown, slideCount?: number): number {
  const numberValue = typeof value === "number" ? value : Number(value);
  if (!Number.isSafeInteger(numberValue) || numberValue < 1) {
    throw new ValidationError("Slide number must be a positive whole number.");
  }
  if (slideCount !== undefined && numberValue > slideCount) {
    throw new ValidationError(`Slide number must be between 1 and ${slideCount}.`);
  }
  return numberValue - 1;
}

export function requireDimension(
  value: unknown,
  fieldName: string,
  options: { readonly allowZero: boolean },
): number {
  const numberValue = typeof value === "number" ? value : Number(value);
  const minimum = options.allowZero ? 0 : Number.EPSILON;
  if (!Number.isFinite(numberValue) || numberValue < minimum || numberValue > 10_000) {
    const range = options.allowZero ? "0 through 10,000" : "greater than 0 and at most 10,000";
    throw new ValidationError(`${fieldName} must be ${range}.`);
  }
  return numberValue;
}

function parseOrigin(value: unknown, fieldName: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new ValidationError(`${fieldName} must be a non-empty HTTPS origin.`);
  }

  let url: URL;
  try {
    url = new URL(value.trim());
  } catch {
    throw new ValidationError(`${fieldName} is not a valid URL.`);
  }

  if (url.protocol !== "https:") {
    throw new ValidationError(`${fieldName} must use HTTPS.`);
  }
  if (
    url.username !== "" ||
    url.password !== "" ||
    url.pathname !== "/" ||
    url.search !== "" ||
    url.hash !== ""
  ) {
    throw new ValidationError(`${fieldName} must contain only an origin, without credentials, a path, query, or fragment.`);
  }
  return url.origin;
}

export function normalizeAllowedOrigins(value: unknown): readonly string[] {
  if (!Array.isArray(value) || value.length === 0 || value.length > 32) {
    throw new ValidationError("allowedBackendOrigins must contain between 1 and 32 HTTPS origins.");
  }
  const origins = value.map((entry, index) => parseOrigin(entry, `Allowed backend origin ${index + 1}`));
  return Object.freeze([...new Set(origins)]);
}

export function requireAllowedBackendUrl(value: unknown, allowedOrigins: readonly string[]): string {
  const origin = parseOrigin(value, "Backend URL");
  if (!allowedOrigins.includes(origin)) {
    throw new ValidationError("Backend URL is not in this deployment's approved origin list.");
  }
  return origin;
}

export function requireEndpointPath(value: unknown, fieldName = "Endpoint path"): string {
  if (typeof value !== "string" || value.length === 0 || value.length > 256) {
    throw new ValidationError(`${fieldName} must be between 1 and 256 characters.`);
  }
  if (
    !value.startsWith("/") ||
    value.startsWith("//") ||
    value.includes("\\") ||
    /[?#\u0000-\u001f\u007f]/.test(value) ||
    value.split("/").some((part) => part === "." || part === "..")
  ) {
    throw new ValidationError(`${fieldName} must be an absolute path without traversal, a query, or a fragment.`);
  }
  return value;
}

export function validateBearerToken(value: unknown): string | undefined {
  if (value === undefined || value === null || value === "") {
    return undefined;
  }
  if (typeof value !== "string") {
    throw new ValidationError("Access token must be text.");
  }
  const token = value.trim();
  if (token.length === 0 || token.length > MAX_TOKEN_LENGTH || /[\r\n]/.test(token)) {
    throw new ValidationError(`Access token must be a single line of at most ${MAX_TOKEN_LENGTH.toLocaleString()} characters.`);
  }
  return token;
}

export function parseRuntimeConfig(value: unknown): RuntimeConfig {
  if (!isRecord(value)) {
    throw new ValidationError("Runtime configuration must be a JSON object.");
  }

  const allowedBackendOrigins = normalizeAllowedOrigins(value.allowedBackendOrigins);
  const defaultBackendUrl = requireAllowedBackendUrl(value.defaultBackendUrl, allowedBackendOrigins);
  const healthPath = requireEndpointPath(value.healthPath, "Health endpoint path");
  const requestTimeoutMs = Number(value.requestTimeoutMs);
  if (!Number.isSafeInteger(requestTimeoutMs) || requestTimeoutMs < 1_000 || requestTimeoutMs > 60_000) {
    throw new ValidationError("requestTimeoutMs must be a whole number from 1,000 through 60,000.");
  }

  return Object.freeze({
    allowedBackendOrigins,
    defaultBackendUrl,
    healthPath,
    requestTimeoutMs,
  });
}
