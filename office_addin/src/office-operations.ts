import {
  boundedText,
  requireDimension,
  requirePlaceholder,
  requireSlideIndex,
  requireText,
} from "./validation";

export type SupportedOfficeHost = "Word" | "PowerPoint";
export type DetectedOfficeHost = SupportedOfficeHost | "Unsupported";
export type WordInsertMode = "Replace" | "Start" | "End";

export interface RequirementSupport {
  readonly requirementSet: "WordApi" | "PowerPointApi";
  readonly version: string;
  readonly supported: boolean;
}

export interface CapabilityReport {
  readonly host: DetectedOfficeHost;
  readonly platform: string;
  readonly officeVersion: string;
  readonly requirements: readonly RequirementSupport[];
}

export interface SlideSummary {
  readonly number: number;
  readonly id: string;
}

export interface TextBoxOptions {
  readonly slideNumber: unknown;
  readonly text: unknown;
  readonly left: unknown;
  readonly top: unknown;
  readonly width: unknown;
  readonly height: unknown;
}

export class OfficeCapabilityError extends Error {
  readonly requirementSet: string;
  readonly minimumVersion: string;

  constructor(message: string, requirementSet: string, minimumVersion: string) {
    super(message);
    this.name = "OfficeCapabilityError";
    this.requirementSet = requirementSet;
    this.minimumVersion = minimumVersion;
  }
}

function officeIsAvailable(): boolean {
  return typeof Office !== "undefined" && Office.context !== undefined;
}

export function detectHost(): DetectedOfficeHost {
  if (!officeIsAvailable()) {
    return "Unsupported";
  }
  if (Office.context.host === Office.HostType.Word) {
    return "Word";
  }
  if (Office.context.host === Office.HostType.PowerPoint) {
    return "PowerPoint";
  }
  return "Unsupported";
}

export function isRequirementSupported(requirementSet: string, minimumVersion: string): boolean {
  if (!officeIsAvailable()) {
    return false;
  }
  try {
    return Office.context.requirements.isSetSupported(requirementSet, minimumVersion);
  } catch {
    return false;
  }
}

function requireHost(host: SupportedOfficeHost): void {
  const detected = detectHost();
  if (detected !== host) {
    throw new OfficeCapabilityError(
      `This operation is available only in ${host}; the current host is ${detected}.`,
      host === "Word" ? "WordApi" : "PowerPointApi",
      "1.1",
    );
  }
}

function requireApi(requirementSet: "WordApi" | "PowerPointApi", minimumVersion: string): void {
  if (!isRequirementSupported(requirementSet, minimumVersion)) {
    throw new OfficeCapabilityError(
      `${requirementSet} ${minimumVersion} is required. Update Microsoft 365 or use a supported Office client.`,
      requirementSet,
      minimumVersion,
    );
  }
}

export function getCapabilityReport(): CapabilityReport {
  const requirements: RequirementSupport[] = [
    {
      requirementSet: "WordApi",
      version: "1.1",
      supported: isRequirementSupported("WordApi", "1.1"),
    },
    {
      requirementSet: "WordApi",
      version: "1.3",
      supported: isRequirementSupported("WordApi", "1.3"),
    },
    {
      requirementSet: "PowerPointApi",
      version: "1.2",
      supported: isRequirementSupported("PowerPointApi", "1.2"),
    },
    {
      requirementSet: "PowerPointApi",
      version: "1.3",
      supported: isRequirementSupported("PowerPointApi", "1.3"),
    },
    {
      requirementSet: "PowerPointApi",
      version: "1.4",
      supported: isRequirementSupported("PowerPointApi", "1.4"),
    },
  ];

  return Object.freeze({
    host: detectHost(),
    platform: officeIsAvailable() ? String(Office.context.platform ?? "Unknown") : "Unavailable",
    officeVersion: officeIsAvailable() ? String(Office.context.diagnostics?.version ?? "Unknown") : "Unavailable",
    requirements: Object.freeze(requirements),
  });
}

export async function readWordSelection(): Promise<string> {
  requireHost("Word");
  requireApi("WordApi", "1.1");
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text;
  });
}

export async function writeWordSelection(contentInput: unknown, mode: WordInsertMode): Promise<void> {
  requireHost("Word");
  requireApi("WordApi", "1.1");
  const content = requireText(contentInput, "Word content");
  const insertionLocations: Record<WordInsertMode, Word.InsertLocation> = {
    Replace: Word.InsertLocation.replace,
    Start: Word.InsertLocation.start,
    End: Word.InsertLocation.end,
  };
  const insertLocation = insertionLocations[mode];
  if (insertLocation === undefined) {
    throw new Error("Unsupported Word insertion mode.");
  }

  await Word.run(async (context) => {
    context.document.getSelection().insertText(content, insertLocation);
    await context.sync();
  });
}

export async function replaceWordPlaceholders(
  placeholderInput: unknown,
  replacementInput: unknown,
): Promise<number> {
  requireHost("Word");
  requireApi("WordApi", "1.1");
  const placeholder = requirePlaceholder(placeholderInput);
  const replacement = boundedText(replacementInput, "Replacement");

  return Word.run(async (context) => {
    const matches = context.document.body.search(placeholder, {
      matchCase: false,
      matchWholeWord: false,
    });
    matches.load("items");
    await context.sync();

    for (const match of matches.items) {
      match.insertText(replacement, Word.InsertLocation.replace);
    }
    await context.sync();
    return matches.items.length;
  });
}

export async function listPowerPointSlides(): Promise<readonly SlideSummary[]> {
  requireHost("PowerPoint");
  requireApi("PowerPointApi", "1.2");
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    return slides.items.map((slide, index) => ({ number: index + 1, id: slide.id }));
  });
}

export async function addPowerPointSlide(): Promise<SlideSummary> {
  requireHost("PowerPoint");
  requireApi("PowerPointApi", "1.2");
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.add();
    await context.sync();
    slides.load("items/id");
    await context.sync();
    const lastIndex = slides.items.length - 1;
    const newSlide = slides.items[lastIndex];
    if (newSlide === undefined) {
      throw new Error("PowerPoint did not return the newly added slide.");
    }
    return { number: lastIndex + 1, id: newSlide.id };
  });
}

export async function deletePowerPointSlide(slideNumberInput: unknown): Promise<void> {
  requireHost("PowerPoint");
  requireApi("PowerPointApi", "1.2");
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    if (slides.items.length <= 1) {
      throw new Error("PowerPoint must retain at least one slide.");
    }
    const zeroBasedIndex = requireSlideIndex(slideNumberInput, slides.items.length);
    slides.getItemAt(zeroBasedIndex).delete();
    await context.sync();
  });
}

export async function insertPowerPointTextBox(options: TextBoxOptions): Promise<string> {
  requireHost("PowerPoint");
  requireApi("PowerPointApi", "1.4");
  const text = requireText(options.text, "PowerPoint text");
  const left = requireDimension(options.left, "Left", { allowZero: true });
  const top = requireDimension(options.top, "Top", { allowZero: true });
  const width = requireDimension(options.width, "Width", { allowZero: false });
  const height = requireDimension(options.height, "Height", { allowZero: false });

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const zeroBasedIndex = requireSlideIndex(options.slideNumber, slides.items.length);
    const shape = slides.getItemAt(zeroBasedIndex).shapes.addTextBox(text, {
      left,
      top,
      width,
      height,
    });
    shape.load("id");
    await context.sync();
    return shape.id;
  });
}
