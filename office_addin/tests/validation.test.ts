import { describe, expect, it } from "vitest";

import {
  ValidationError,
  normalizeAllowedOrigins,
  parseRuntimeConfig,
  requireAllowedBackendUrl,
  requireEndpointPath,
  requireSlideIndex,
  validateBearerToken,
} from "../src/validation";

describe("backend URL validation", () => {
  it("normalizes and deduplicates approved HTTPS origins", () => {
    expect(
      normalizeAllowedOrigins([
        "https://api.example.test",
        "https://api.example.test/",
        "https://localhost:8000",
      ]),
    ).toEqual(["https://api.example.test", "https://localhost:8000"]);
  });

  it.each([
    "http://api.example.test",
    "https://user:password@api.example.test",
    "https://api.example.test/v1",
    "https://api.example.test?target=elsewhere",
    "not a url",
  ])("rejects unsafe origin %s", (origin) => {
    expect(() => normalizeAllowedOrigins([origin])).toThrow(ValidationError);
  });

  it("rejects a valid URL when its origin is not approved", () => {
    expect(() =>
      requireAllowedBackendUrl("https://unapproved.example.test", ["https://api.example.test"]),
    ).toThrow(/approved origin/i);
  });

  it.each(["https://api.example.test/v1", "https://api.example.test.evil.test", "http://api.example.test"])(
    "does not accept near-match backend URL %s",
    (url) => {
      expect(() => requireAllowedBackendUrl(url, ["https://api.example.test"])).toThrow();
    },
  );
});

describe("endpoint validation", () => {
  it("accepts a simple absolute path", () => {
    expect(requireEndpointPath("/health/live")).toBe("/health/live");
  });

  it.each(["https://evil.test/", "//evil.test/", "/../secret", "/health?redirect=x", "/health#x", "/a\\b"])(
    "rejects endpoint %s",
    (path) => {
      expect(() => requireEndpointPath(path)).toThrow(ValidationError);
    },
  );
});

describe("runtime configuration", () => {
  it("parses a complete valid configuration", () => {
    const config = parseRuntimeConfig({
      allowedBackendOrigins: ["https://api.example.test"],
      defaultBackendUrl: "https://api.example.test/",
      healthPath: "/health",
      requestTimeoutMs: 10_000,
    });
    expect(config.defaultBackendUrl).toBe("https://api.example.test");
    expect(config.healthPath).toBe("/health");
    expect(config.requestTimeoutMs).toBe(10_000);
  });

  it("rejects an out-of-range timeout", () => {
    expect(() =>
      parseRuntimeConfig({
        allowedBackendOrigins: ["https://api.example.test"],
        defaultBackendUrl: "https://api.example.test",
        healthPath: "/health",
        requestTimeoutMs: 100,
      }),
    ).toThrow(/requestTimeoutMs/);
  });
});

describe("document input validation", () => {
  it("converts a one-based slide number to a zero-based index", () => {
    expect(requireSlideIndex("3", 4)).toBe(2);
  });

  it.each(["0", "1.2", "5", "not-a-number"])("rejects invalid slide number %s", (number) => {
    expect(() => requireSlideIndex(number, 4)).toThrow(ValidationError);
  });

  it("rejects header injection in bearer tokens", () => {
    expect(() => validateBearerToken("safe-looking\r\nX-Evil: yes")).toThrow(ValidationError);
  });
});
