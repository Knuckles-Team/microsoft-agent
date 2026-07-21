import { beforeEach, describe, expect, it, vi } from "vitest";

const operations = vi.hoisted(() => ({
  readWordSelection: vi.fn(async () => "selected text"),
  writeWordSelection: vi.fn(async () => undefined),
  replaceWordPlaceholders: vi.fn(async () => 3),
  listPowerPointSlides: vi.fn(async () => [{ number: 1, id: "slide-1" }]),
  addPowerPointSlide: vi.fn(async () => ({ number: 2, id: "slide-2" })),
  deletePowerPointSlide: vi.fn(async () => undefined),
  insertPowerPointTextBox: vi.fn(async () => "shape-1"),
}));

vi.mock("../src/office-operations", () => ({
  ...operations,
  OfficeCapabilityError: class OfficeCapabilityError extends Error {},
}));

import { BackendClient, type FetchImplementation } from "../src/backend-client";
import {
  OfficeBridgeClient,
  executeOfficeCommand,
  parseOfficeCommandEnvelope,
  type OfficeCommandEnvelope,
} from "../src/office-bridge";
import type { CapabilityReport } from "../src/office-operations";

const sessionId = "11111111-1111-4111-8111-111111111111";
const commandId = "22222222-2222-4222-8222-222222222222";
const future = "2099-01-01T00:00:00Z";

const capabilityReport: CapabilityReport = {
  host: "Word",
  platform: "PC",
  officeVersion: "16.0",
  requirements: [
    { requirementSet: "WordApi", version: "1.1", supported: true },
  ],
};

function command(payload: OfficeCommandEnvelope["payload"]): OfficeCommandEnvelope {
  return {
    command_id: commandId,
    session_id: sessionId,
    created_at: "2026-07-17T12:00:00Z",
    expires_at: future,
    payload,
  };
}

describe("Office bridge protocol", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("rejects commands for another paired session or host", () => {
    const raw = command({ kind: "word.read_selection" });
    expect(() =>
      parseOfficeCommandEnvelope(raw, {
        session_id: "33333333-3333-4333-8333-333333333333",
        host: "Word",
      }),
    ).toThrow(/another session/i);
    expect(() =>
      parseOfficeCommandEnvelope(raw, { session_id: sessionId, host: "PowerPoint" }),
    ).toThrow(/does not match/i);
  });

  it("dispatches only a modeled Word operation and returns its typed result", async () => {
    const envelope = command({
      kind: "word.write_selection",
      content: "Approved content",
      mode: "Replace",
    });

    await expect(executeOfficeCommand(envelope)).resolves.toEqual({
      command_id: commandId,
      status: "succeeded",
      kind: "word.write_selection",
      result: { kind: "word.write_selection", applied: true },
    });
    expect(operations.writeWordSelection).toHaveBeenCalledWith("Approved content", "Replace");
  });

  it("uses the one-time pairing body and bearer token only after pairing", async () => {
    const fetchImplementation = vi.fn<FetchImplementation>(async (url) => {
      if (url.endsWith("/office-bridge/session")) {
        return new Response(
          JSON.stringify({
            session_id: sessionId,
            session_token: "a".repeat(43),
            host: "Word",
            label: "Budget draft",
            expires_at: future,
          }),
          { status: 201, headers: { "content-type": "application/json" } },
        );
      }
      return new Response(JSON.stringify({ command: null }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });
    const backend = new BackendClient(
      "https://api.example.test",
      ["https://api.example.test"],
      10_000,
      fetchImplementation,
    );
    const client = new OfficeBridgeClient(backend);

    const session = await client.pair("p".repeat(43), capabilityReport);
    await expect(client.poll(session, 0)).resolves.toBeNull();

    const [, pairInit] = fetchImplementation.mock.calls[0] ?? [];
    const [, pollInit] = fetchImplementation.mock.calls[1] ?? [];
    expect((pairInit?.headers as Record<string, string>).Authorization).toBeUndefined();
    expect(JSON.parse(String(pairInit?.body))).toMatchObject({
      pairing_token: "p".repeat(43),
      capabilities: { host: "Word", office_version: "16.0" },
    });
    expect((pollInit?.headers as Record<string, string>).Authorization).toBe(
      `Bearer ${"a".repeat(43)}`,
    );
  });

  it("rejects arbitrary or expired commands before Office APIs run", () => {
    expect(() =>
      parseOfficeCommandEnvelope(
        { ...command({ kind: "word.read_selection" }), payload: { kind: "word.run_javascript", code: "evil()" } },
        { session_id: sessionId, host: "Word" },
      ),
    ).toThrow(/not supported/i);
    expect(() =>
      parseOfficeCommandEnvelope(
        { ...command({ kind: "word.read_selection" }), expires_at: "2020-01-01T00:00:00Z" },
        { session_id: sessionId, host: "Word" },
      ),
    ).toThrow(/expired/i);
    expect(operations.readWordSelection).not.toHaveBeenCalled();
  });

  it("requires visible user confirmation before a remote slide deletion", async () => {
    const envelope = command({
      kind: "powerpoint.delete_slide",
      slide_number: 2,
    });

    await expect(executeOfficeCommand(envelope, () => false)).rejects.toThrow(/declined/i);
    expect(operations.deletePowerPointSlide).not.toHaveBeenCalled();
  });
});
