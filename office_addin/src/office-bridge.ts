import { BackendClient } from "./backend-client";
import {
  OfficeCapabilityError,
  addPowerPointSlide,
  deletePowerPointSlide,
  insertPowerPointTextBox,
  listPowerPointSlides,
  readWordSelection,
  replaceWordPlaceholders,
  writeWordSelection,
  type CapabilityReport,
  type SupportedOfficeHost,
  type WordInsertMode,
} from "./office-operations";
import { ValidationError, boundedText, validateBearerToken } from "./validation";

const PAIR_PATH = "/office-bridge/session";
const POLL_PATH = "/office-bridge/commands/poll";
const RESULT_PATH = "/office-bridge/commands/result";
const DISCONNECT_PATH = "/office-bridge/session/disconnect";
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const MAX_ERROR_MESSAGE_LENGTH = 1_000;

export type OfficeCommandKind =
  | "word.read_selection"
  | "word.write_selection"
  | "word.replace_placeholders"
  | "powerpoint.list_slides"
  | "powerpoint.add_slide"
  | "powerpoint.delete_slide"
  | "powerpoint.insert_text_box";

interface WordReadSelectionPayload {
  readonly kind: "word.read_selection";
}

interface WordWriteSelectionPayload {
  readonly kind: "word.write_selection";
  readonly content: unknown;
  readonly mode: unknown;
}

interface WordReplacePlaceholdersPayload {
  readonly kind: "word.replace_placeholders";
  readonly placeholder: unknown;
  readonly replacement: unknown;
}

interface PowerPointListSlidesPayload {
  readonly kind: "powerpoint.list_slides";
}

interface PowerPointAddSlidePayload {
  readonly kind: "powerpoint.add_slide";
}

interface PowerPointDeleteSlidePayload {
  readonly kind: "powerpoint.delete_slide";
  readonly slide_number: unknown;
}

interface PowerPointInsertTextBoxPayload {
  readonly kind: "powerpoint.insert_text_box";
  readonly slide_number: unknown;
  readonly text: unknown;
  readonly left: unknown;
  readonly top: unknown;
  readonly width: unknown;
  readonly height: unknown;
}

export type OfficeCommandPayload =
  | WordReadSelectionPayload
  | WordWriteSelectionPayload
  | WordReplacePlaceholdersPayload
  | PowerPointListSlidesPayload
  | PowerPointAddSlidePayload
  | PowerPointDeleteSlidePayload
  | PowerPointInsertTextBoxPayload;

export interface OfficeCommandEnvelope {
  readonly command_id: string;
  readonly session_id: string;
  readonly created_at: string;
  readonly expires_at: string;
  readonly payload: OfficeCommandPayload;
}

export interface OfficeSessionGrant {
  readonly session_id: string;
  readonly session_token: string;
  readonly host: SupportedOfficeHost;
  readonly label: string;
  readonly expires_at: string;
}

interface OfficeCommandError {
  readonly code: "capability" | "validation" | "office" | "unexpected";
  readonly message: string;
}

type OfficeCommandResult =
  | { readonly kind: "word.read_selection"; readonly text: string }
  | { readonly kind: "word.write_selection"; readonly applied: true }
  | { readonly kind: "word.replace_placeholders"; readonly replacements: number }
  | {
      readonly kind: "powerpoint.list_slides";
      readonly slides: readonly { readonly number: number; readonly id: string }[];
    }
  | {
      readonly kind: "powerpoint.add_slide";
      readonly slide: { readonly number: number; readonly id: string };
    }
  | { readonly kind: "powerpoint.delete_slide"; readonly deleted: true }
  | { readonly kind: "powerpoint.insert_text_box"; readonly shape_id: string };

export type OfficeCommandOutcome =
  | {
      readonly command_id: string;
      readonly status: "succeeded";
      readonly kind: OfficeCommandKind;
      readonly result: OfficeCommandResult;
    }
  | {
      readonly command_id: string;
      readonly status: "failed";
      readonly kind: OfficeCommandKind;
      readonly error: OfficeCommandError;
    };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function requireRecord(value: unknown, name: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new ValidationError(`${name} must be a JSON object.`);
  }
  return value;
}

function requireString(value: unknown, name: string, maximumLength: number): string {
  if (typeof value !== "string" || value.length === 0 || value.length > maximumLength) {
    throw new ValidationError(`${name} is invalid.`);
  }
  return value;
}

function requireUuid(value: unknown, name: string): string {
  const candidate = requireString(value, name, 36);
  if (!UUID_PATTERN.test(candidate)) {
    throw new ValidationError(`${name} is invalid.`);
  }
  return candidate;
}

function requireTimestamp(value: unknown, name: string): string {
  const candidate = requireString(value, name, 64);
  if (!Number.isFinite(Date.parse(candidate))) {
    throw new ValidationError(`${name} is invalid.`);
  }
  return candidate;
}

function requireCommandKind(value: unknown): OfficeCommandKind {
  switch (value) {
    case "word.read_selection":
    case "word.write_selection":
    case "word.replace_placeholders":
    case "powerpoint.list_slides":
    case "powerpoint.add_slide":
    case "powerpoint.delete_slide":
    case "powerpoint.insert_text_box":
      return value;
    default:
      throw new ValidationError("Office command kind is not supported.");
  }
}

function parsePayload(value: unknown): OfficeCommandPayload {
  const payload = requireRecord(value, "Office command payload");
  const kind = requireCommandKind(payload.kind);
  switch (kind) {
    case "word.read_selection":
    case "powerpoint.list_slides":
    case "powerpoint.add_slide":
      return { kind };
    case "word.write_selection":
      return {
        kind,
        content: payload.content,
        mode: payload.mode,
      };
    case "word.replace_placeholders":
      return {
        kind,
        placeholder: payload.placeholder,
        replacement: payload.replacement,
      };
    case "powerpoint.delete_slide":
      return { kind, slide_number: payload.slide_number };
    case "powerpoint.insert_text_box":
      return {
        kind,
        slide_number: payload.slide_number,
        text: payload.text,
        left: payload.left,
        top: payload.top,
        width: payload.width,
        height: payload.height,
      };
  }
}

export function parseOfficeCommandEnvelope(
  value: unknown,
  session: Pick<OfficeSessionGrant, "session_id" | "host">,
): OfficeCommandEnvelope {
  const envelope = requireRecord(value, "Office command");
  const commandId = requireUuid(envelope.command_id, "Office command ID");
  const sessionId = requireUuid(envelope.session_id, "Office session ID");
  if (sessionId !== session.session_id) {
    throw new ValidationError("Office command belongs to another session.");
  }
  const payload = parsePayload(envelope.payload);
  const expectedHost: SupportedOfficeHost = payload.kind.startsWith("word.") ? "Word" : "PowerPoint";
  if (expectedHost !== session.host) {
    throw new ValidationError("Office command does not match this Office host.");
  }
  const createdAt = requireTimestamp(envelope.created_at, "Office command creation time");
  const expiresAt = requireTimestamp(envelope.expires_at, "Office command expiry");
  if (Date.parse(expiresAt) <= Date.now()) {
    throw new ValidationError("Office command has expired.");
  }
  return {
    command_id: commandId,
    session_id: sessionId,
    created_at: createdAt,
    expires_at: expiresAt,
    payload,
  };
}

function requirePairingToken(value: unknown): string {
  if (typeof value !== "string") {
    throw new ValidationError("Pairing secret must be text.");
  }
  const token = value.trim();
  if (token.length < 32 || token.length > 256 || /\s/.test(token)) {
    throw new ValidationError("Pairing secret is invalid.");
  }
  return token;
}

function capabilityRequest(report: CapabilityReport): Record<string, unknown> {
  if (report.host !== "Word" && report.host !== "PowerPoint") {
    throw new ValidationError("Pairing requires a supported Word or PowerPoint host.");
  }
  return {
    host: report.host,
    platform: report.platform,
    office_version: report.officeVersion,
    requirements: report.requirements.map((requirement) => ({
      requirement_set: requirement.requirementSet,
      version: requirement.version,
      supported: requirement.supported,
    })),
  };
}

function parseSessionGrant(value: unknown): OfficeSessionGrant {
  const grant = requireRecord(value, "Office session grant");
  const host = grant.host;
  if (host !== "Word" && host !== "PowerPoint") {
    throw new ValidationError("Office session host is invalid.");
  }
  const token = validateBearerToken(grant.session_token);
  if (token === undefined || token.length < 32) {
    throw new ValidationError("Office session token is invalid.");
  }
  return {
    session_id: requireUuid(grant.session_id, "Office session ID"),
    session_token: token,
    host,
    label: requireString(grant.label, "Office session label", 128),
    expires_at: requireTimestamp(grant.expires_at, "Office session expiry"),
  };
}

function parsePolledCommand(value: unknown, session: OfficeSessionGrant): OfficeCommandEnvelope | null {
  const response = requireRecord(value, "Office poll response");
  if (response.command === null) {
    return null;
  }
  return parseOfficeCommandEnvelope(response.command, session);
}

export class OfficeBridgeClient {
  private readonly backend: BackendClient;

  constructor(backend: BackendClient) {
    this.backend = backend;
  }

  async pair(pairingToken: unknown, capabilities: CapabilityReport): Promise<OfficeSessionGrant> {
    const response = await this.backend.requestJson(PAIR_PATH, {
      method: "POST",
      body: {
        pairing_token: requirePairingToken(pairingToken),
        capabilities: capabilityRequest(capabilities),
      },
    });
    const grant = parseSessionGrant(response);
    if (grant.host !== capabilities.host) {
      throw new ValidationError("Backend paired this task pane to another Office host.");
    }
    return grant;
  }

  async poll(
    session: OfficeSessionGrant,
    waitSeconds: number,
    signal?: AbortSignal,
  ): Promise<OfficeCommandEnvelope | null> {
    if (!Number.isFinite(waitSeconds) || waitSeconds < 0 || waitSeconds > 5) {
      throw new ValidationError("Office poll wait must be between 0 and 5 seconds.");
    }
    const options = signal === undefined
      ? {
          method: "POST" as const,
          token: session.session_token,
          body: { wait_seconds: waitSeconds },
        }
      : {
          method: "POST" as const,
          token: session.session_token,
          body: { wait_seconds: waitSeconds },
          signal,
        };
    return parsePolledCommand(await this.backend.requestJson(POLL_PATH, options), session);
  }

  async submit(
    session: OfficeSessionGrant,
    outcome: OfficeCommandOutcome,
    signal?: AbortSignal,
  ): Promise<void> {
    const options = signal === undefined
      ? { method: "POST" as const, token: session.session_token, body: outcome }
      : { method: "POST" as const, token: session.session_token, body: outcome, signal };
    await this.backend.requestJson(RESULT_PATH, options);
  }

  async disconnect(session: OfficeSessionGrant): Promise<void> {
    await this.backend.requestJson(DISCONNECT_PATH, {
      method: "POST",
      token: session.session_token,
      body: {},
    });
  }
}

function requireInsertMode(value: unknown): WordInsertMode {
  if (value === "Replace" || value === "Start" || value === "End") {
    return value;
  }
  throw new ValidationError("Word insertion mode is invalid.");
}

function defaultDestructiveConfirmation(message: string): boolean {
  return typeof globalThis.confirm === "function" && globalThis.confirm(message);
}

export async function executeOfficeCommand(
  command: OfficeCommandEnvelope,
  confirmDestructive: (message: string) => boolean = defaultDestructiveConfirmation,
): Promise<OfficeCommandOutcome> {
  const payload = command.payload;
  let result: OfficeCommandResult;
  switch (payload.kind) {
    case "word.read_selection":
      result = { kind: payload.kind, text: await readWordSelection() };
      break;
    case "word.write_selection":
      await writeWordSelection(payload.content, requireInsertMode(payload.mode));
      result = { kind: payload.kind, applied: true };
      break;
    case "word.replace_placeholders":
      result = {
        kind: payload.kind,
        replacements: await replaceWordPlaceholders(payload.placeholder, payload.replacement),
      };
      break;
    case "powerpoint.list_slides":
      result = { kind: payload.kind, slides: await listPowerPointSlides() };
      break;
    case "powerpoint.add_slide":
      result = { kind: payload.kind, slide: await addPowerPointSlide() };
      break;
    case "powerpoint.delete_slide":
      if (!confirmDestructive(`Allow the agent to permanently delete slide ${String(payload.slide_number)}?`)) {
        throw new ValidationError("The user declined the PowerPoint slide deletion.");
      }
      await deletePowerPointSlide(payload.slide_number);
      result = { kind: payload.kind, deleted: true };
      break;
    case "powerpoint.insert_text_box":
      result = {
        kind: payload.kind,
        shape_id: await insertPowerPointTextBox({
          slideNumber: payload.slide_number,
          text: payload.text,
          left: payload.left,
          top: payload.top,
          width: payload.width,
          height: payload.height,
        }),
      };
      break;
  }
  return {
    command_id: command.command_id,
    status: "succeeded",
    kind: payload.kind,
    result,
  };
}

function failureOutcome(command: OfficeCommandEnvelope, error: unknown): OfficeCommandOutcome {
  let code: OfficeCommandError["code"] = "unexpected";
  if (error instanceof OfficeCapabilityError) {
    code = "capability";
  } else if (error instanceof ValidationError) {
    code = "validation";
  } else if (error instanceof Error) {
    code = "office";
  }
  const rawMessage = error instanceof Error && error.message.trim() !== ""
    ? error.message
    : "The Office operation could not be completed.";
  return {
    command_id: command.command_id,
    status: "failed",
    kind: command.payload.kind,
    error: {
      code,
      message: boundedText(rawMessage, "Office error", MAX_ERROR_MESSAGE_LENGTH),
    },
  };
}

export class OfficeBridgeRunner {
  private readonly client: OfficeBridgeClient;
  private readonly session: OfficeSessionGrant;
  private readonly onStatus: (message: string, isError: boolean) => void;

  constructor(
    client: OfficeBridgeClient,
    session: OfficeSessionGrant,
    onStatus: (message: string, isError: boolean) => void = () => undefined,
  ) {
    this.client = client;
    this.session = session;
    this.onStatus = onStatus;
  }

  async pollOnce(signal?: AbortSignal): Promise<boolean> {
    const command = await this.client.poll(this.session, 5, signal);
    if (command === null) {
      return false;
    }
    let outcome: OfficeCommandOutcome;
    try {
      outcome = await executeOfficeCommand(command);
    } catch (error: unknown) {
      outcome = failureOutcome(command, error);
    }
    await this.client.submit(this.session, outcome, signal);
    this.onStatus(
      outcome.status === "succeeded"
        ? `Completed ${outcome.kind}.`
        : `${outcome.kind} failed: ${outcome.error.message}`,
      outcome.status === "failed",
    );
    return true;
  }

  async run(signal: AbortSignal): Promise<void> {
    while (!signal.aborted) {
      try {
        await this.pollOnce(signal);
      } catch (error: unknown) {
        if (signal.aborted) {
          return;
        }
        const message = error instanceof Error ? error.message : "Office bridge connection failed.";
        this.onStatus(message, true);
        await new Promise<void>((resolve) => globalThis.setTimeout(resolve, 1_000));
      }
    }
  }
}
