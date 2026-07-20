import { BackendClient, loadRuntimeConfig } from "./backend-client";
import {
  OfficeBridgeClient,
  OfficeBridgeRunner,
  type OfficeSessionGrant,
} from "./office-bridge";
import {
  addPowerPointSlide,
  deletePowerPointSlide,
  detectHost,
  getCapabilityReport,
  insertPowerPointTextBox,
  isRequirementSupported,
  listPowerPointSlides,
  readWordSelection,
  replaceWordPlaceholders,
  writeWordSelection,
  type SlideSummary,
  type WordInsertMode,
} from "./office-operations";
import {
  requireAllowedBackendUrl,
  type RuntimeConfig,
  validateBearerToken,
} from "./validation";

const BACKEND_URL_STORAGE_KEY = "microsoft-agent-office-addin.backend-origin";
const MAX_DISPLAY_LENGTH = 20_000;

let runtimeConfig: RuntimeConfig | undefined;
let officeBridgeClient: OfficeBridgeClient | undefined;
let officeBridgeSession: OfficeSessionGrant | undefined;
let officeBridgeAbortController: AbortController | undefined;

function element<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (found === null) {
    throw new Error(`Required page element '${id}' is missing.`);
  }
  return found as T;
}

function inputValue(id: string): string {
  return element<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(id).value;
}

function setStatus(message: string, kind: "info" | "success" | "error" = "info"): void {
  const status = element<HTMLDivElement>("status");
  status.textContent = message;
  status.dataset.kind = kind;
}

function userMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim() !== "") {
    return error.message;
  }
  return "The operation could not be completed.";
}

function displayJson(value: unknown): string {
  let rendered: string;
  try {
    rendered = JSON.stringify(value, null, 2);
  } catch {
    rendered = "The result could not be displayed as JSON.";
  }
  if (rendered.length <= MAX_DISPLAY_LENGTH) {
    return rendered;
  }
  return `${rendered.slice(0, MAX_DISPLAY_LENGTH)}\n… display truncated …`;
}

async function runButtonAction(
  buttonId: string,
  workingMessage: string,
  action: () => Promise<void>,
): Promise<void> {
  const button = element<HTMLButtonElement>(buttonId);
  button.disabled = true;
  setStatus(workingMessage);
  try {
    await action();
  } catch (error: unknown) {
    setStatus(userMessage(error), "error");
  } finally {
    button.disabled = false;
  }
}

function refreshCapabilityReport(): void {
  const report = getCapabilityReport();
  element<HTMLPreElement>("capability-report").textContent = displayJson(report);
  element<HTMLParagraphElement>("host-summary").textContent =
    report.host === "Unsupported" ? "Unsupported or unavailable Office host" : `${report.host} on ${report.platform}`;
  element<HTMLElement>("word-tools").hidden = report.host !== "Word";
  element<HTMLElement>("powerpoint-tools").hidden = report.host !== "PowerPoint";

  const textButton = element<HTMLButtonElement>("insert-ppt-text");
  const supportsTextBoxes = report.host === "PowerPoint" && isRequirementSupported("PowerPointApi", "1.4");
  textButton.disabled = !supportsTextBoxes;
  textButton.title = supportsTextBoxes
    ? "Insert a text box"
    : "This client does not support PowerPointApi 1.4 text boxes.";
}

function renderSlides(slides: readonly SlideSummary[]): void {
  const list = element<HTMLOListElement>("slide-list");
  list.replaceChildren();
  for (const slide of slides) {
    const item = document.createElement("li");
    item.textContent = `Slide ${slide.number} — ${slide.id}`;
    list.append(item);
  }
}

function safelyReadSavedBackendUrl(): string | undefined {
  try {
    return globalThis.localStorage.getItem(BACKEND_URL_STORAGE_KEY) ?? undefined;
  } catch {
    return undefined;
  }
}

function safelySaveBackendUrl(url: string): void {
  try {
    globalThis.localStorage.setItem(BACKEND_URL_STORAGE_KEY, url);
  } catch {
    throw new Error("The backend URL was validated but this Office client blocked local settings storage.");
  }
}

function requireRuntimeConfig(): RuntimeConfig {
  if (runtimeConfig === undefined) {
    throw new Error("Backend configuration has not loaded. Reload the task pane and verify config.json.");
  }
  return runtimeConfig;
}

async function initializeBackend(): Promise<void> {
  try {
    const loadedConfig = await loadRuntimeConfig();
    runtimeConfig = loadedConfig;
    const savedUrl = safelyReadSavedBackendUrl();
    let selectedUrl = loadedConfig.defaultBackendUrl;
    if (savedUrl !== undefined) {
      try {
        selectedUrl = requireAllowedBackendUrl(savedUrl, loadedConfig.allowedBackendOrigins);
      } catch {
        // Ignore stale settings when a deployment removes an origin from its allowlist.
      }
    }
    element<HTMLInputElement>("backend-url").value = selectedUrl;
    element<HTMLPreElement>("backend-result").textContent = `Approved origins:\n${loadedConfig.allowedBackendOrigins.join("\n")}`;
  } catch (error: unknown) {
    element<HTMLPreElement>("backend-result").textContent = userMessage(error);
    setStatus(userMessage(error), "error");
  }
}

function registerWordHandlers(): void {
  element<HTMLButtonElement>("read-word-selection").addEventListener("click", () => {
    void runButtonAction("read-word-selection", "Reading Word selection…", async () => {
      const selection = await readWordSelection();
      const output = element<HTMLTextAreaElement>("word-selection");
      output.value = selection.slice(0, MAX_DISPLAY_LENGTH);
      setStatus(
        selection.length > MAX_DISPLAY_LENGTH
          ? `Selection read; showing the first ${MAX_DISPLAY_LENGTH.toLocaleString()} characters.`
          : "Selection read.",
        "success",
      );
    });
  });

  element<HTMLButtonElement>("write-word-content").addEventListener("click", () => {
    void runButtonAction("write-word-content", "Updating Word selection…", async () => {
      const mode = inputValue("word-insert-mode") as WordInsertMode;
      await writeWordSelection(inputValue("word-content"), mode);
      setStatus("Word content updated.", "success");
    });
  });

  element<HTMLButtonElement>("replace-word-placeholders").addEventListener("click", () => {
    void runButtonAction("replace-word-placeholders", "Replacing Word placeholders…", async () => {
      const count = await replaceWordPlaceholders(
        inputValue("word-placeholder"),
        inputValue("word-replacement"),
      );
      setStatus(`Replaced ${count.toLocaleString()} occurrence${count === 1 ? "" : "s"}.`, "success");
    });
  });
}

function registerPowerPointHandlers(): void {
  element<HTMLButtonElement>("list-slides").addEventListener("click", () => {
    void runButtonAction("list-slides", "Reading PowerPoint slides…", async () => {
      const slides = await listPowerPointSlides();
      renderSlides(slides);
      setStatus(`Found ${slides.length.toLocaleString()} slide${slides.length === 1 ? "" : "s"}.`, "success");
    });
  });

  element<HTMLButtonElement>("add-slide").addEventListener("click", () => {
    void runButtonAction("add-slide", "Adding a PowerPoint slide…", async () => {
      const slide = await addPowerPointSlide();
      renderSlides(await listPowerPointSlides());
      setStatus(`Added slide ${slide.number}.`, "success");
    });
  });

  element<HTMLButtonElement>("delete-slide").addEventListener("click", () => {
    const slideNumber = inputValue("delete-slide-index");
    if (!globalThis.confirm(`Permanently delete slide ${slideNumber || "?"}?`)) {
      setStatus("Slide deletion cancelled.");
      return;
    }
    void runButtonAction("delete-slide", "Deleting the PowerPoint slide…", async () => {
      await deletePowerPointSlide(slideNumber);
      renderSlides(await listPowerPointSlides());
      setStatus(`Deleted slide ${slideNumber}.`, "success");
    });
  });

  element<HTMLButtonElement>("insert-ppt-text").addEventListener("click", () => {
    void runButtonAction("insert-ppt-text", "Inserting a PowerPoint text box…", async () => {
      const shapeId = await insertPowerPointTextBox({
        slideNumber: inputValue("ppt-text-slide-index"),
        text: inputValue("ppt-text"),
        left: inputValue("ppt-left"),
        top: inputValue("ppt-top"),
        width: inputValue("ppt-width"),
        height: inputValue("ppt-height"),
      });
      setStatus(`Inserted text box ${shapeId}.`, "success");
    });
  });
}

function registerBackendHandlers(): void {
  element<HTMLButtonElement>("save-backend").addEventListener("click", () => {
    try {
      const config = requireRuntimeConfig();
      const approvedUrl = requireAllowedBackendUrl(inputValue("backend-url"), config.allowedBackendOrigins);
      safelySaveBackendUrl(approvedUrl);
      element<HTMLInputElement>("backend-url").value = approvedUrl;
      setStatus("Approved backend URL saved. Access tokens are never saved.", "success");
    } catch (error: unknown) {
      setStatus(userMessage(error), "error");
    }
  });

  element<HTMLButtonElement>("clear-backend-token").addEventListener("click", () => {
    element<HTMLInputElement>("backend-token").value = "";
    setStatus("In-memory access token cleared.", "success");
  });

  element<HTMLButtonElement>("test-backend").addEventListener("click", () => {
    void runButtonAction("test-backend", "Testing the approved backend…", async () => {
      const config = requireRuntimeConfig();
      const backendUrl = requireAllowedBackendUrl(inputValue("backend-url"), config.allowedBackendOrigins);
      const token = validateBearerToken(inputValue("backend-token"));
      const client = new BackendClient(
        backendUrl,
        config.allowedBackendOrigins,
        config.requestTimeoutMs,
      );
      const options = token === undefined ? {} : { token };
      const result = await client.requestJson(config.healthPath, options);
      element<HTMLPreElement>("backend-result").textContent = displayJson(result);
      setStatus("Backend connection succeeded.", "success");
    });
  });

  element<HTMLButtonElement>("pair-office").addEventListener("click", () => {
    void runButtonAction("pair-office", "Pairing this Office task pane…", async () => {
      const config = requireRuntimeConfig();
      const backendUrl = requireAllowedBackendUrl(inputValue("backend-url"), config.allowedBackendOrigins);
      const backend = new BackendClient(
        backendUrl,
        config.allowedBackendOrigins,
        config.requestTimeoutMs,
      );
      const client = new OfficeBridgeClient(backend);
      const session = await client.pair(inputValue("office-pairing-token"), getCapabilityReport());

      officeBridgeAbortController?.abort();
      officeBridgeClient = client;
      officeBridgeSession = session;
      officeBridgeAbortController = new AbortController();
      element<HTMLInputElement>("office-pairing-token").value = "";
      element<HTMLButtonElement>("disconnect-office").disabled = false;
      element<HTMLPreElement>("office-bridge-result").textContent = displayJson({
        session_id: session.session_id,
        host: session.host,
        label: session.label,
        expires_at: session.expires_at,
      });
      const runner = new OfficeBridgeRunner(client, session, (message, isError) => {
        setStatus(message, isError ? "error" : "success");
      });
      void runner.run(officeBridgeAbortController.signal);
      setStatus(`Paired as “${session.label}”. Agent commands are now enabled.`, "success");
    });
  });

  element<HTMLButtonElement>("disconnect-office").addEventListener("click", () => {
    const client = officeBridgeClient;
    const session = officeBridgeSession;
    officeBridgeAbortController?.abort();
    officeBridgeAbortController = undefined;
    officeBridgeClient = undefined;
    officeBridgeSession = undefined;
    element<HTMLButtonElement>("disconnect-office").disabled = true;
    element<HTMLPreElement>("office-bridge-result").textContent = "Office bridge is disconnected.";
    setStatus("Office bridge disconnected locally.", "success");
    if (client !== undefined && session !== undefined) {
      void client.disconnect(session).catch(() => {
        // The in-memory bearer token has already been discarded. Server-side
        // idle and absolute expiry still revoke a failed disconnect request.
      });
    }
  });
}

async function initializeOffice(): Promise<void> {
  if (typeof Office === "undefined") {
    setStatus("Office.js did not load. Open this page from a sideloaded Word or PowerPoint add-in.", "error");
    return;
  }
  await Office.onReady();
  refreshCapabilityReport();
  const host = detectHost();
  setStatus(
    host === "Unsupported"
      ? "Open this task pane in a supported Word or PowerPoint client."
      : `${host} is ready.`,
    host === "Unsupported" ? "error" : "success",
  );
}

function start(): void {
  registerWordHandlers();
  registerPowerPointHandlers();
  registerBackendHandlers();
  element<HTMLButtonElement>("refresh-capabilities").addEventListener("click", refreshCapabilityReport);
  void initializeBackend();
  void initializeOffice().catch((error: unknown) => setStatus(userMessage(error), "error"));
}

start();
