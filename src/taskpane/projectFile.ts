import type { ImageItem, SortMode } from "./types";

export const PROJECT_FILE_NAME = "werkhaus_bilddaten.json";
export const LEGACY_PROJECT_FILE_NAMES = ["bilddaten.json"] as const;
export const PROJECT_SCHEMA = "werkhaus.word-bild.image-metadata";
export const PROJECT_FORMAT_VERSION = 1;
const LEGACY_PROJECT_TYPE = "WERKHAUS-Bilddaten";
const LEGACY_PROJECT_SCHEMA_VERSION = 1;
const DEFAULT_DOCUMENT_BASIS = "Blanco-Dokument";
const DEFAULT_OUTPUT_BASENAME = "werkhaus_bildprojekt";
const DEFAULT_PREVIEW_SIZE_PX = 120;
const DEFAULT_TOOL_VERSION = "1.0.0.10";

const DEFAULT_ANALYSIS_PROFILES = {
  allgemein: {
    key: "allgemein",
    label: "Allgemeine Analyse",
    description: "Beschreibung des Bildes in Bezug auf Gebäudemerkmale",
    version: 1,
    system_prompt:
      "Du analysierst Bilder fuer eine WERKHAUS Bild- und Analyseplattform. Antworte nur als JSON.",
    user_prompt:
      "Analysiere dieses Bild allgemein und sachlich. Antworte ausschliesslich als JSON.",
  },
  risse: {
    key: "risse",
    label: "Risse",
    description: "Analyse auf Rissbildung und Oberflaechenrisse.",
    version: 1,
    system_prompt:
      "Du suchst gezielt nach Rissbildung, Materialtrennung und Oberflaechenrissen. Antworte nur als JSON.",
    user_prompt: "Analysiere dieses Bild mit Fokus auf Risse. Antworte ausschliesslich als JSON.",
  },
  feuchtigkeit: {
    key: "feuchtigkeit",
    label: "Feuchtigkeit",
    description: "Analyse auf Feuchte-, Wasser- oder Schimmelsignale.",
    version: 1,
    system_prompt:
      "Du suchst gezielt nach Feuchte-, Wasser- oder Schimmelsignalen. Antworte nur als JSON.",
    user_prompt:
      "Analysiere dieses Bild mit Fokus auf Feuchtigkeit. Antworte ausschliesslich als JSON.",
  },
  schimmel: {
    key: "schimmel",
    label: "Schimmel",
    description: "Analyse auf Schimmelspuren und Folgeschaeden.",
    version: 1,
    system_prompt:
      "Du suchst gezielt nach Schimmelspuren und Folgeschaeden. Antworte nur als JSON.",
    user_prompt:
      "Analysiere dieses Bild mit Fokus auf Schimmel. Antworte ausschliesslich als JSON.",
  },
  fassade: {
    key: "fassade",
    label: "Fassade",
    description: "Analyse von Fassaden, Anschluessen und Aussenbauteilen.",
    version: 1,
    system_prompt: "Du analysierst Fassaden und Aussenbauteile. Antworte nur als JSON.",
    user_prompt: "Analysiere dieses Bild mit Fokus auf Fassade. Antworte ausschliesslich als JSON.",
  },
  dach: {
    key: "dach",
    label: "Dach",
    description: "Analyse von Dachbereichen, Eindeckung und Anschluessen.",
    version: 1,
    system_prompt:
      "Du analysierst Dachbereiche, Eindeckung und Anschluesse. Antworte nur als JSON.",
    user_prompt:
      "Analysiere dieses Bild mit Fokus auf das Dach. Antworte ausschliesslich als JSON.",
  },
} as const;

export interface ProjectDocumentSettings {
  basis?: string;
  template_path?: string;
  start_page?: number;
  [key: string]: unknown;
}

export interface ProjectOutputSettings {
  images_per_page?: number;
  layout_images_per_page?: number;
  sort_columns_per_row?: number;
  sort_card_size?: string;
  image_management_filter?: string;
  caption_font_size?: number;
  compression?: string;
  output_basename?: string;
  sort_mode?: SortMode;
  [key: string]: unknown;
}

export interface ProjectPreviewSettings {
  load_scope?: string;
  quality?: string;
  mode?: string;
  zoom_percent?: number;
  preview_size_px?: number;
  [key: string]: unknown;
}

export interface ProjectImportSettings {
  use_werkhaus_json?: boolean;
  [key: string]: unknown;
}

export interface ProjectCaptionsSettings {
  show_image_number?: boolean;
  show_filename?: boolean;
  show_date?: boolean;
  show_time?: boolean;
  show_caption?: boolean;
  [key: string]: unknown;
}

export interface ProjectUiSettings {
  right_box_caption_layout?: boolean;
  right_box_output?: boolean;
  right_box_word_document?: boolean;
  right_box_ai_api?: boolean;
  left_box_selection_open?: boolean;
  right_box_creation?: boolean;
  right_box_output_caption_layout?: boolean;
  right_box_folder?: boolean;
  left_box_list_open?: boolean;
  right_box_analysis_profiles?: boolean;
  right_box_project_filename?: boolean;
  sort_mode?: SortMode;
  preview_size_px?: number;
  insert_size_cm?: number;
  show_info?: boolean;
  show_captions?: boolean;
  [key: string]: unknown;
}

export interface ProjectAnalysisResult {
  image_key?: string;
  analysis_mode?: string;
  analysis_status?: string;
  openai_description?: string;
  suggested_caption?: string;
  user_decision?: string;
  timestamp?: string;
  model_info?: {
    provider?: string;
    vision_model?: string;
    [key: string]: unknown;
  };
  error_message?: string;
  analysis_context?: string;
  [key: string]: unknown;
}

export interface ProjectAnalysisEntry {
  analysis_id?: string;
  image_key?: string;
  profile_key?: string;
  profile_version?: number;
  status?: string;
  created_at?: string;
  model?: string;
  image_hash?: string;
  image_type?: string;
  semantic_tags?: unknown[];
  region_hint?: string;
  relations?: unknown[];
  result?: ProjectAnalysisResult;
  image_id?: string;
  profile?: string;
  hash?: string;
  [key: string]: unknown;
}

export interface ProjectAnalysisImageGroup {
  image_key?: string;
  image_id?: string;
  analyses?: ProjectAnalysisEntry[];
  [key: string]: unknown;
}

export interface ProjectAnalysisProfile {
  key?: string;
  label?: string;
  description?: string;
  version?: number;
  system_prompt?: string;
  user_prompt?: string;
  [key: string]: unknown;
}

export interface BilddatenProjectImage {
  id?: string;
  key?: string;
  image_key?: string;
  image_id?: string;
  image_hash?: string;
  relative_path?: string;
  fileName?: string;
  filename?: string;
  full_path?: string;
  size?: number;
  lastModified?: number;
  active?: boolean;
  selected?: boolean;
  visible?: boolean;
  position?: number;
  location?: string;
  caption?: string;
  image_number?: string;
  includeCaptionInWord?: boolean;
  include_caption_in_word?: boolean;
  analyses?: ProjectAnalysisEntry[];
  analysis?: ProjectAnalysisEntry[] | Record<string, unknown>;
  [key: string]: unknown;
}

export interface BilddatenProjectFile {
  schema: string;
  version: number;
  project_format_version: number;
  tool_version?: string;
  createdAt: string;
  updatedAt: string;
  saved_at: string;
  image_folder?: string;
  include_subfolders: boolean;
  document: ProjectDocumentSettings;
  output: ProjectOutputSettings;
  preview: ProjectPreviewSettings;
  import: ProjectImportSettings;
  captions: ProjectCaptionsSettings;
  ui: ProjectUiSettings;
  images: BilddatenProjectImage[];
  analysis: {
    schema_version: number;
    images: Record<string, ProjectAnalysisImageGroup>;
  };
  analysis_profiles: Record<string, ProjectAnalysisProfile>;
  analysis_ui: Record<string, unknown>;
  sortMode?: SortMode;
  [key: string]: unknown;
}

export interface ProjectParseResult {
  projectFile: BilddatenProjectFile;
  format: "current" | "legacy";
  warnings: string[];
}

export interface ProjectMatch {
  projectIndex: number;
  image: BilddatenProjectImage;
  itemIndex: number;
}

export interface ProjectMatchResult {
  matches: ProjectMatch[];
  unmatchedProjectImages: BilddatenProjectImage[];
  unmatchedItemIndexes: number[];
}

export interface ProjectSaveSnapshot {
  sortMode: SortMode;
  previewSizePx?: number;
  insertSizeCm?: number;
  showInfo?: boolean;
  showCaptions?: boolean;
  collapsedSections?: string[];
}

export interface ProjectValidationResult {
  valid: boolean;
  format: "current" | "legacy" | "unknown";
  projectFile?: BilddatenProjectFile;
  warnings: string[];
  errors: string[];
}

export interface ProjectStateUpdate {
  toolVersion?: string;
  imageFolder?: string;
  includeSubfolders?: boolean;
  createdAt?: string;
  updatedAt?: string;
  savedAt?: string;
  document?: Partial<ProjectDocumentSettings>;
  output?: Partial<ProjectOutputSettings>;
  preview?: Partial<ProjectPreviewSettings>;
  import?: Partial<ProjectImportSettings>;
  captions?: Partial<ProjectCaptionsSettings>;
  ui?: Partial<ProjectUiSettings>;
  images?: BilddatenProjectImage[];
  analysis?: BilddatenProjectFile["analysis"];
  analysisProfiles?: Record<string, ProjectAnalysisProfile>;
  analysisUi?: Record<string, unknown>;
  sortMode?: SortMode;
}

export function buildBilddatenProjectFile(
  items: ImageItem[],
  sortMode: SortMode,
  createdAt?: string,
  previousProjectFile?: BilddatenProjectFile,
  snapshot?: ProjectSaveSnapshot
): BilddatenProjectFile {
  const now = new Date();
  const nowIso = now.toISOString();
  const savedAt = formatLocalTimestamp(now);
  const baseProject = previousProjectFile
    ? cloneProjectFile(previousProjectFile)
    : createDefaultProjectFile();
  const matchResult = previousProjectFile
    ? matchProjectImagesToItems(previousProjectFile, items)
    : undefined;
  const matchedProjectIndexes = new Set<number>();
  const images: BilddatenProjectImage[] = [];

  for (let itemIndex = 0; itemIndex < items.length; itemIndex += 1) {
    const item = items[itemIndex];
    const matched = matchResult?.matches.find((entry) => entry.itemIndex === itemIndex);

    if (matched && previousProjectFile) {
      matchedProjectIndexes.add(matched.projectIndex);
      images.push(
        buildProjectImageFromItem(
          item,
          itemIndex + 1,
          previousProjectFile.images[matched.projectIndex]
        )
      );
      continue;
    }

    images.push(buildProjectImageFromItem(item, itemIndex + 1));
  }

  if (previousProjectFile) {
    previousProjectFile.images.forEach((projectImage, projectIndex) => {
      if (!matchedProjectIndexes.has(projectIndex)) {
        images.push(cloneProjectImage(projectImage));
      }
    });
  }

  const mergedUi = {
    ...cloneRecord(baseProject.ui),
    ...(snapshot?.collapsedSections
      ? buildCollapsedSectionSnapshot(snapshot.collapsedSections)
      : {}),
    sort_mode: sortMode,
    preview_size_px: Number.isFinite(snapshot?.previewSizePx ?? NaN)
      ? snapshot?.previewSizePx
      : baseProject.ui.preview_size_px,
    insert_size_cm: Number.isFinite(snapshot?.insertSizeCm ?? NaN)
      ? snapshot?.insertSizeCm
      : baseProject.ui.insert_size_cm,
    show_info: snapshot?.showInfo ?? baseProject.ui.show_info ?? false,
    show_captions: snapshot?.showCaptions ?? baseProject.ui.show_captions ?? false,
  } as ProjectUiSettings;

  const projectFile: BilddatenProjectFile = {
    ...baseProject,
    schema: PROJECT_SCHEMA,
    version: normalizeProjectVersion(baseProject.version, PROJECT_FORMAT_VERSION),
    project_format_version: normalizeProjectVersion(
      baseProject.project_format_version,
      PROJECT_FORMAT_VERSION
    ),
    tool_version: baseProject.tool_version,
    createdAt: normalizeString(createdAt) || baseProject.createdAt || nowIso,
    updatedAt: nowIso,
    saved_at: savedAt,
    image_folder: normalizeString(baseProject.image_folder) || undefined,
    include_subfolders: Boolean(baseProject.include_subfolders),
    document: cloneDocumentSettings(baseProject.document),
    output: cloneOutputSettings(baseProject.output, sortMode),
    preview: clonePreviewSettings(baseProject.preview, snapshot?.previewSizePx),
    import: cloneImportSettings(baseProject.import),
    captions: cloneCaptionSettings(baseProject.captions),
    ui: mergedUi,
    images,
    analysis: cloneAnalysisBlock(baseProject.analysis),
    analysis_profiles: cloneAnalysisProfiles(baseProject.analysis_profiles),
    analysis_ui: cloneRecord(baseProject.analysis_ui),
    sortMode,
  };

  if (!projectFile.output.output_basename) {
    projectFile.output.output_basename = DEFAULT_OUTPUT_BASENAME;
  }

  return projectFile;
}

export function createEmptyProjectForFolder(
  imageFolder?: string,
  state?: ProjectStateUpdate
): BilddatenProjectFile {
  const projectFile = createDefaultProjectFile();
  projectFile.image_folder = normalizeOptionalText(imageFolder) || undefined;
  return updateProjectFromState(projectFile, state || {});
}

export function validateProjectJson(
  input: string | Record<string, unknown>
): ProjectValidationResult {
  const warnings: string[] = [];
  const errors: string[] = [];

  if (typeof input === "string") {
    try {
      const parsed = JSON.parse(input) as Record<string, unknown>;
      return validateProjectJson(parsed);
    } catch (error) {
      return {
        valid: false,
        format: "unknown",
        warnings,
        errors: [
          "Projektdatei konnte nicht als JSON gelesen werden.",
          error instanceof Error ? error.message : String(error),
        ],
      };
    }
  }

  if (!input || typeof input !== "object" || Array.isArray(input)) {
    return {
      valid: false,
      format: "unknown",
      warnings,
      errors: ["Projektdatei hat kein lesbares Objektformat."],
    };
  }

  if (isCurrentProjectFormat(input)) {
    const projectFile = normalizeCurrentProjectFile(input);
    warnings.push(...buildCurrentFormatWarnings(input));
    return {
      valid: true,
      format: "current",
      projectFile,
      warnings,
      errors,
    };
  }

  if (isLegacyProjectFormat(input)) {
    return {
      valid: true,
      format: "legacy",
      projectFile: convertLegacyProjectFile(input),
      warnings: ["Legacy-Format erkannt. Die Datei wurde in das neue Projektformat ueberfuehrt."],
      errors,
    };
  }

  return {
    valid: false,
    format: "unknown",
    warnings,
    errors: ["Unbekanntes Projektformat."],
  };
}

export function normalizeProjectJson(
  input: string | Record<string, unknown>
): BilddatenProjectFile | undefined {
  return validateProjectJson(input).projectFile;
}

export function updateProjectFromState(
  projectFile: BilddatenProjectFile,
  state: ProjectStateUpdate
): BilddatenProjectFile {
  const nextProject: Record<string, unknown> = {
    ...cloneProjectFile(projectFile),
    tool_version: state.toolVersion ?? projectFile.tool_version,
    createdAt: state.createdAt ?? projectFile.createdAt,
    updatedAt: state.updatedAt ?? projectFile.updatedAt,
    saved_at: state.savedAt ?? projectFile.saved_at,
    image_folder: state.imageFolder ?? projectFile.image_folder,
    include_subfolders: state.includeSubfolders ?? projectFile.include_subfolders,
    document: state.document
      ? { ...projectFile.document, ...state.document }
      : projectFile.document,
    output: state.output ? { ...projectFile.output, ...state.output } : projectFile.output,
    preview: state.preview ? { ...projectFile.preview, ...state.preview } : projectFile.preview,
    import: state.import ? { ...projectFile.import, ...state.import } : projectFile.import,
    captions: state.captions
      ? { ...projectFile.captions, ...state.captions }
      : projectFile.captions,
    ui: state.ui ? { ...projectFile.ui, ...state.ui } : projectFile.ui,
    images: state.images
      ? state.images.map((image) => cloneProjectImage(image))
      : projectFile.images,
    analysis: state.analysis ? cloneAnalysisBlock(state.analysis) : projectFile.analysis,
    analysis_profiles: state.analysisProfiles
      ? cloneAnalysisProfiles(state.analysisProfiles)
      : projectFile.analysis_profiles,
    analysis_ui: state.analysisUi ? cloneRecord(state.analysisUi) : projectFile.analysis_ui,
    sortMode: state.sortMode ?? projectFile.sortMode,
  };

  return normalizeCurrentProjectFile(nextProject);
}

export function mergeFolderImagesWithProject(
  items: ImageItem[],
  projectFile: BilddatenProjectFile
): {
  projectFile: BilddatenProjectFile;
  matchResult: ProjectMatchResult;
} {
  const matchResult = matchProjectImagesToItems(projectFile, items);
  const mergedProjectFile = buildBilddatenProjectFile(
    items,
    projectFile.sortMode || projectFile.ui?.sort_mode || "custom",
    projectFile.createdAt,
    projectFile
  );

  return {
    projectFile: mergedProjectFile,
    matchResult,
  };
}

export function exportProjectJson(projectFile: BilddatenProjectFile): string {
  return serializeBilddatenProjectFile(projectFile);
}

export function serializeBilddatenProjectFile(projectFile: BilddatenProjectFile): string {
  return JSON.stringify(projectFile, null, 2);
}

export function parseBilddatenProjectFile(rawText: string): ProjectParseResult | undefined {
  const validation = validateProjectJson(rawText);

  if (!validation.valid || !validation.projectFile) {
    return undefined;
  }

  return {
    projectFile: validation.projectFile,
    format: validation.format === "unknown" ? "current" : validation.format,
    warnings: validation.warnings,
  };
}

export function matchProjectImagesToItems(
  projectFile: BilddatenProjectFile,
  items: ImageItem[]
): ProjectMatchResult {
  const usedItemIndexes = new Set<number>();
  const matches: ProjectMatch[] = [];
  const unmatchedProjectImages: BilddatenProjectImage[] = [];

  projectFile.images.forEach((projectImage, projectIndex) => {
    const itemIndex = findBestMatchingItemIndex(projectImage, items, usedItemIndexes);

    if (itemIndex === -1) {
      unmatchedProjectImages.push(projectImage);
      return;
    }

    usedItemIndexes.add(itemIndex);
    matches.push({
      projectIndex,
      image: projectImage,
      itemIndex,
    });
  });

  const unmatchedItemIndexes: number[] = [];
  for (let index = 0; index < items.length; index += 1) {
    if (!usedItemIndexes.has(index)) {
      unmatchedItemIndexes.push(index);
    }
  }

  return {
    matches,
    unmatchedProjectImages,
    unmatchedItemIndexes,
  };
}

export function getSuggestedProjectFileName(
  projectFile: Pick<BilddatenProjectFile, "output">
): string {
  const baseName = normalizeString(projectFile.output?.output_basename) || DEFAULT_OUTPUT_BASENAME;
  const safeBaseName = sanitizeFileName(baseName) || DEFAULT_OUTPUT_BASENAME;
  return `${safeBaseName}.json`;
}

function buildProjectImageFromItem(
  item: ImageItem,
  position: number,
  baseImage?: BilddatenProjectImage
): BilddatenProjectImage {
  const fileNumber = String(position).padStart(3, "0");
  const existing = baseImage ? cloneProjectImage(baseImage) : {};

  return {
    ...existing,
    id: item.id,
    key: item.key,
    image_key: item.key,
    image_id: item.id,
    image_hash: item.hash,
    relative_path: normalizeOptionalText(item.relativePath) || item.name,
    fileName: item.name,
    filename: item.name,
    full_path:
      normalizeOptionalText(item.fullPath) || normalizeOptionalText(item.relativePath) || item.name,
    size: Number.isFinite(item.size) ? item.size : existing.size,
    lastModified: Number.isFinite(item.lastModified ?? NaN)
      ? item.lastModified
      : existing.lastModified,
    active: item.selected !== false,
    selected: item.selected !== false,
    visible: item.visible ?? item.selected !== false,
    position,
    location: fileNumber,
    caption: normalizeOptionalText(item.caption) ?? existing.caption ?? "",
    image_number: `Bild ${fileNumber}`,
    includeCaptionInWord: item.includeCaptionInWord !== false,
    include_caption_in_word: item.includeCaptionInWord !== false,
  };
}

function normalizeCurrentProjectFile(raw: Record<string, unknown>): BilddatenProjectFile {
  const base = createDefaultProjectFile();
  const ui = normalizeRecord(raw.ui);
  const output = normalizeRecord(raw.output);
  const preview = normalizeRecord(raw.preview);

  const sortMode = normalizeSortMode(
    normalizeString(ui.sort_mode) ||
      normalizeString(output.sort_mode) ||
      normalizeString(raw.sortMode)
  );

  return {
    ...base,
    ...raw,
    schema: normalizeString(raw.schema) || PROJECT_SCHEMA,
    version: normalizeNumber(raw.version, PROJECT_FORMAT_VERSION),
    project_format_version: normalizeNumber(raw.project_format_version, PROJECT_FORMAT_VERSION),
    tool_version: normalizeOptionalText(raw.tool_version) || base.tool_version,
    createdAt: normalizeOptionalText(raw.createdAt) || base.createdAt,
    updatedAt: normalizeOptionalText(raw.updatedAt) || base.updatedAt,
    saved_at: normalizeOptionalText(raw.saved_at) || formatLocalTimestamp(new Date()),
    image_folder: normalizeOptionalText(raw.image_folder) || base.image_folder,
    include_subfolders: normalizeBoolean(raw.include_subfolders, false),
    document: normalizeDocumentSettings(raw.document),
    output: normalizeOutputSettings(output, sortMode),
    preview: normalizePreviewSettings(preview),
    import: normalizeImportSettings(raw.import),
    captions: normalizeCaptionSettings(raw.captions),
    ui: normalizeUiSettings(ui, sortMode),
    images: normalizeProjectImages(raw.images),
    analysis: normalizeAnalysisBlock(raw.analysis),
    analysis_profiles: normalizeAnalysisProfiles(raw.analysis_profiles),
    analysis_ui: normalizeRecord(raw.analysis_ui),
    sortMode,
  };
}

function convertLegacyProjectFile(raw: Record<string, unknown>): BilddatenProjectFile {
  const base = createDefaultProjectFile();
  const legacyImages = Array.isArray(raw.images) ? raw.images : [];
  const sortMode = normalizeSortMode(normalizeString(raw.sortMode));

  return {
    ...base,
    ...raw,
    schema: PROJECT_SCHEMA,
    version: PROJECT_FORMAT_VERSION,
    project_format_version: PROJECT_FORMAT_VERSION,
    tool_version: normalizeOptionalText(raw.tool_version) || base.tool_version,
    createdAt: normalizeOptionalText(raw.createdAt) || base.createdAt,
    updatedAt: normalizeOptionalText(raw.updatedAt) || base.updatedAt,
    saved_at: normalizeOptionalText(raw.updatedAt) || formatLocalTimestamp(new Date()),
    image_folder: normalizeOptionalText(raw.image_folder) || base.image_folder,
    include_subfolders: normalizeBoolean(raw.include_subfolders, false),
    document: normalizeDocumentSettings(raw.document),
    output: normalizeOutputSettings(normalizeRecord(raw.output), sortMode),
    preview: normalizePreviewSettings(normalizeRecord(raw.preview)),
    import: normalizeImportSettings(raw.import),
    captions: normalizeCaptionSettings(raw.captions),
    ui: normalizeUiSettings(normalizeRecord(raw.ui), sortMode),
    images: legacyImages
      .map((image) => normalizeLegacyProjectImage(image))
      .filter(Boolean) as BilddatenProjectImage[],
    analysis: normalizeAnalysisBlock(raw.analysis),
    analysis_profiles: normalizeAnalysisProfiles(raw.analysis_profiles),
    analysis_ui: normalizeRecord(raw.analysis_ui),
    sortMode,
    projectType: normalizeOptionalText(raw.projectType) || LEGACY_PROJECT_TYPE,
    schemaVersion: normalizeNumber(raw.schemaVersion, LEGACY_PROJECT_SCHEMA_VERSION),
  };
}

function createDefaultProjectFile(): BilddatenProjectFile {
  const nowIso = new Date().toISOString();
  return {
    schema: PROJECT_SCHEMA,
    version: PROJECT_FORMAT_VERSION,
    project_format_version: PROJECT_FORMAT_VERSION,
    tool_version: DEFAULT_TOOL_VERSION,
    createdAt: nowIso,
    updatedAt: nowIso,
    saved_at: formatLocalTimestamp(new Date()),
    include_subfolders: false,
    document: createDefaultDocumentSettings(),
    output: createDefaultOutputSettings(),
    preview: createDefaultPreviewSettings(),
    import: createDefaultImportSettings(),
    captions: createDefaultCaptionSettings(),
    ui: createDefaultUiSettings(),
    images: [],
    analysis: {
      schema_version: 1,
      images: {},
    },
    analysis_profiles: createDefaultAnalysisProfiles(),
    analysis_ui: createDefaultAnalysisUi(),
    sortMode: "custom",
  };
}

function cloneProjectFile(projectFile: BilddatenProjectFile): BilddatenProjectFile {
  return {
    ...projectFile,
    document: cloneDocumentSettings(projectFile.document),
    output: cloneOutputSettings(projectFile.output, projectFile.sortMode || "custom"),
    preview: clonePreviewSettings(projectFile.preview),
    import: cloneImportSettings(projectFile.import),
    captions: cloneCaptionSettings(projectFile.captions),
    ui: normalizeUiSettings(normalizeRecord(projectFile.ui), projectFile.sortMode || "custom"),
    images: projectFile.images.map((image) => cloneProjectImage(image)),
    analysis: cloneAnalysisBlock(projectFile.analysis),
    analysis_profiles: cloneAnalysisProfiles(projectFile.analysis_profiles),
    analysis_ui: cloneRecord(projectFile.analysis_ui),
  };
}

function cloneProjectImage(image: BilddatenProjectImage): BilddatenProjectImage {
  return {
    ...image,
    analyses: Array.isArray(image.analyses)
      ? image.analyses.map((analysis) => cloneProjectAnalysisEntry(analysis))
      : image.analyses,
    analysis: Array.isArray(image.analysis)
      ? image.analysis.map((analysis) =>
          cloneProjectAnalysisEntry(analysis as ProjectAnalysisEntry)
        )
      : normalizeRecord(image.analysis),
  };
}

function cloneProjectAnalysisEntry(entry: ProjectAnalysisEntry): ProjectAnalysisEntry {
  return {
    ...entry,
    result: entry.result
      ? {
          ...entry.result,
          model_info: entry.result.model_info
            ? { ...entry.result.model_info }
            : entry.result.model_info,
        }
      : entry.result,
    semantic_tags: Array.isArray(entry.semantic_tags)
      ? [...entry.semantic_tags]
      : entry.semantic_tags,
    relations: Array.isArray(entry.relations) ? [...entry.relations] : entry.relations,
  };
}

function cloneDocumentSettings(settings: ProjectDocumentSettings): ProjectDocumentSettings {
  return {
    ...createDefaultDocumentSettings(),
    ...normalizeRecord(settings),
  };
}

function cloneOutputSettings(
  settings: ProjectOutputSettings,
  sortMode: SortMode
): ProjectOutputSettings {
  return {
    ...createDefaultOutputSettings(),
    ...normalizeRecord(settings),
    sort_mode: normalizeSortMode(normalizeString(settings.sort_mode) || sortMode),
  };
}

function clonePreviewSettings(
  settings: ProjectPreviewSettings,
  previewSizePx?: number
): ProjectPreviewSettings {
  return {
    ...createDefaultPreviewSettings(),
    ...normalizeRecord(settings),
    preview_size_px: Number.isFinite(previewSizePx ?? NaN)
      ? previewSizePx
      : normalizeNumber(settings.preview_size_px, DEFAULT_PREVIEW_SIZE_PX),
  };
}

function cloneImportSettings(settings: ProjectImportSettings): ProjectImportSettings {
  return {
    ...createDefaultImportSettings(),
    ...normalizeRecord(settings),
  };
}

function cloneCaptionSettings(settings: ProjectCaptionsSettings): ProjectCaptionsSettings {
  return {
    ...createDefaultCaptionSettings(),
    ...normalizeRecord(settings),
  };
}

function cloneAnalysisBlock(
  analysis: BilddatenProjectFile["analysis"]
): BilddatenProjectFile["analysis"] {
  const images: Record<string, ProjectAnalysisImageGroup> = {};

  for (const [key, value] of Object.entries(analysis.images || {})) {
    images[key] = {
      ...normalizeRecord(value),
      image_key: normalizeOptionalText(value.image_key) || key,
      image_id:
        normalizeOptionalText(value.image_id) || normalizeOptionalText(value.image_key) || key,
      analyses: Array.isArray(value.analyses)
        ? value.analyses.map((entry) => cloneProjectAnalysisEntry(entry))
        : [],
    };
  }

  return {
    schema_version: normalizeNumber(analysis.schema_version, 1),
    images,
  };
}

function cloneAnalysisProfiles(
  profiles: Record<string, ProjectAnalysisProfile>
): Record<string, ProjectAnalysisProfile> {
  const cloned: Record<string, ProjectAnalysisProfile> = {};

  for (const [key, value] of Object.entries(profiles || {})) {
    cloned[key] = {
      ...normalizeRecord(value),
      key: normalizeOptionalText(value.key) || key,
      label: normalizeOptionalText(value.label) || key,
      description: normalizeOptionalText(value.description),
      version: normalizeNumber(value.version, 1),
      system_prompt: normalizeOptionalText(value.system_prompt),
      user_prompt: normalizeOptionalText(value.user_prompt),
    };
  }

  return cloned;
}

function normalizeProjectImages(rawImages: unknown): BilddatenProjectImage[] {
  if (!Array.isArray(rawImages)) {
    return [];
  }

  return rawImages
    .map((image) => normalizeProjectImage(image))
    .filter(Boolean) as BilddatenProjectImage[];
}

function normalizeProjectImage(rawImage: unknown): BilddatenProjectImage | undefined {
  if (!rawImage || typeof rawImage !== "object") {
    return undefined;
  }

  const image = rawImage as Record<string, unknown>;
  const filename =
    normalizeOptionalText(image.fileName) ||
    normalizeOptionalText(image.filename) ||
    normalizeOptionalText(image.name) ||
    normalizeOptionalText(image.relative_path) ||
    normalizeOptionalText(image.relativePath);

  if (!filename) {
    return undefined;
  }

  const position = normalizeNumber(image.position, 0);
  const fileNumber = String(Math.max(1, Math.floor(position || 1))).padStart(3, "0");
  const active = normalizeBoolean(image.active, normalizeBoolean(image.selected, true));
  const selected = normalizeBoolean(image.selected, active);
  const visible = normalizeBoolean(image.visible, selected);
  const key =
    normalizeOptionalText(image.key) ||
    normalizeOptionalText(image.image_key) ||
    normalizeOptionalText(image.full_path) ||
    normalizeOptionalText(image.relative_path) ||
    normalizeOptionalText(image.relativePath) ||
    filename;

  return {
    ...normalizeRecord(image),
    id: normalizeOptionalText(image.id) || filename,
    key,
    image_key: normalizeOptionalText(image.image_key) || key,
    image_id: normalizeOptionalText(image.image_id) || normalizeOptionalText(image.id) || filename,
    image_hash: normalizeOptionalText(image.image_hash) || normalizeOptionalText(image.hash),
    relative_path:
      normalizeOptionalText(image.relative_path) ||
      normalizeOptionalText(image.relativePath) ||
      filename,
    fileName: normalizeOptionalText(image.fileName) || filename,
    filename: normalizeOptionalText(image.filename) || filename,
    full_path:
      normalizeOptionalText(image.full_path) ||
      normalizeOptionalText(image.fullPath) ||
      normalizeOptionalText(image.relative_path) ||
      normalizeOptionalText(image.relativePath) ||
      filename,
    size: normalizeNumber(image.size, 0),
    lastModified: normalizeNumber(image.lastModified),
    active,
    selected,
    visible,
    position: Number.isFinite(position) && position > 0 ? Math.floor(position) : 0,
    location: normalizeOptionalText(image.location) || fileNumber,
    caption: normalizeOptionalText(image.caption) || "",
    image_number: normalizeOptionalText(image.image_number) || `Bild ${fileNumber}`,
    includeCaptionInWord: normalizeBoolean(
      image.includeCaptionInWord,
      normalizeBoolean(image.include_caption_in_word, true)
    ),
    include_caption_in_word: normalizeBoolean(
      image.include_caption_in_word,
      normalizeBoolean(image.includeCaptionInWord, true)
    ),
    analyses: Array.isArray(image.analyses)
      ? image.analyses.map((analysis) =>
          cloneProjectAnalysisEntry(analysis as ProjectAnalysisEntry)
        )
      : undefined,
    analysis: Array.isArray(image.analysis)
      ? image.analysis.map((analysis) =>
          cloneProjectAnalysisEntry(analysis as ProjectAnalysisEntry)
        )
      : normalizeRecord(image.analysis),
  };
}

function normalizeLegacyProjectImage(rawImage: unknown): BilddatenProjectImage | undefined {
  if (!rawImage || typeof rawImage !== "object") {
    return undefined;
  }

  const image = rawImage as Record<string, unknown>;
  const filename =
    normalizeOptionalText(image.filename) ||
    normalizeOptionalText(image.fileName) ||
    normalizeOptionalText(image.name);

  if (!filename) {
    return undefined;
  }

  const position = normalizeNumber(image.position, 0);
  const fileNumber = String(Math.max(1, Math.floor(position || 1))).padStart(3, "0");
  const active = normalizeBoolean(image.active, true);
  const selected = normalizeBoolean(image.includeCaptionInWord, true);

  return {
    id: filename,
    key: normalizeOptionalText(image.relativePath) || filename,
    image_key: normalizeOptionalText(image.relativePath) || filename,
    image_id: filename,
    image_hash: normalizeOptionalText(image.image_hash) || normalizeOptionalText(image.hash),
    relative_path: normalizeOptionalText(image.relativePath) || filename,
    fileName: filename,
    filename,
    full_path: normalizeOptionalText(image.relativePath) || filename,
    size: normalizeNumber(image.fileSize, 0),
    lastModified: normalizeNumber(image.lastModified),
    active,
    selected: normalizeBoolean(image.active, true),
    visible: active,
    position: Number.isFinite(position) && position > 0 ? Math.floor(position) : 0,
    location: fileNumber,
    caption: normalizeOptionalText(image.caption) || "",
    image_number: normalizeOptionalText(image.imageNumber) || `Bild ${fileNumber}`,
    includeCaptionInWord: selected,
    include_caption_in_word: selected,
  };
}

function normalizeDocumentSettings(raw: unknown): ProjectDocumentSettings {
  const record = normalizeRecord(raw);
  return {
    ...createDefaultDocumentSettings(),
    ...record,
  };
}

function normalizeOutputSettings(
  raw: Record<string, unknown>,
  sortMode: SortMode
): ProjectOutputSettings {
  return {
    ...createDefaultOutputSettings(),
    ...raw,
    sort_mode: normalizeSortMode(normalizeString(raw.sort_mode) || sortMode),
  };
}

function normalizePreviewSettings(raw: Record<string, unknown>): ProjectPreviewSettings {
  return {
    ...createDefaultPreviewSettings(),
    ...raw,
  };
}

function normalizeImportSettings(raw: unknown): ProjectImportSettings {
  const record = normalizeRecord(raw);
  return {
    ...createDefaultImportSettings(),
    ...record,
  };
}

function normalizeCaptionSettings(raw: unknown): ProjectCaptionsSettings {
  const record = normalizeRecord(raw);
  return {
    ...createDefaultCaptionSettings(),
    ...record,
  };
}

function normalizeUiSettings(raw: Record<string, unknown>, sortMode: SortMode): ProjectUiSettings {
  return {
    ...createDefaultUiSettings(),
    ...raw,
    sort_mode: normalizeSortMode(normalizeString(raw.sort_mode) || sortMode),
  };
}

function normalizeAnalysisBlock(raw: unknown): BilddatenProjectFile["analysis"] {
  const record = normalizeRecord(raw);
  const imagesRecord: Record<string, ProjectAnalysisImageGroup> = {};
  const rawImages = record.images;

  if (rawImages && typeof rawImages === "object" && !Array.isArray(rawImages)) {
    for (const [key, value] of Object.entries(rawImages as Record<string, unknown>)) {
      const valueRecord = normalizeRecord(value);
      const analyses = Array.isArray(valueRecord.analyses)
        ? valueRecord.analyses.map((entry) =>
            cloneProjectAnalysisEntry(entry as ProjectAnalysisEntry)
          )
        : [];

      imagesRecord[key] = {
        ...valueRecord,
        image_key: normalizeOptionalText(valueRecord.image_key) || key,
        image_id:
          normalizeOptionalText(valueRecord.image_id) ||
          normalizeOptionalText(valueRecord.image_key) ||
          key,
        analyses,
      };
    }
  } else if (Array.isArray(rawImages)) {
    rawImages.forEach((value, index) => {
      const valueRecord = normalizeRecord(value);
      const entryKey =
        normalizeOptionalText(valueRecord.image_key) ||
        normalizeOptionalText(valueRecord.image_id) ||
        normalizeOptionalText(valueRecord.key) ||
        String(index);
      const analyses = Array.isArray(valueRecord.analyses)
        ? valueRecord.analyses.map((entry) =>
            cloneProjectAnalysisEntry(entry as ProjectAnalysisEntry)
          )
        : [];

      imagesRecord[entryKey] = {
        ...valueRecord,
        image_key: normalizeOptionalText(valueRecord.image_key) || entryKey,
        image_id: normalizeOptionalText(valueRecord.image_id) || entryKey,
        analyses,
      };
    });
  }

  return {
    schema_version: normalizeNumber(record.schema_version, 1),
    images: imagesRecord,
  };
}

function normalizeAnalysisProfiles(raw: unknown): Record<string, ProjectAnalysisProfile> {
  const profiles: Record<string, ProjectAnalysisProfile> = createDefaultAnalysisProfiles();

  if (!raw) {
    return profiles;
  }

  if (Array.isArray(raw)) {
    raw.forEach((entry, index) => {
      const record = normalizeRecord(entry);
      const key =
        normalizeOptionalText(record.key) ||
        normalizeOptionalText(record.label) ||
        `profile-${index}`;

      profiles[key] = {
        ...record,
        key,
        label: normalizeOptionalText(record.label) || key,
        description: normalizeOptionalText(record.description),
        version: normalizeNumber(record.version, 1),
        system_prompt: normalizeOptionalText(record.system_prompt),
        user_prompt: normalizeOptionalText(record.user_prompt),
      };
    });

    return profiles;
  }

  const record = normalizeRecord(raw);
  for (const [key, value] of Object.entries(record)) {
    const valueRecord = normalizeRecord(value);
    const normalizedKey = normalizeOptionalText(valueRecord.key) || key;

    profiles[normalizedKey] = {
      ...valueRecord,
      key: normalizedKey,
      label: normalizeOptionalText(valueRecord.label) || normalizedKey,
      description: normalizeOptionalText(valueRecord.description),
      version: normalizeNumber(valueRecord.version, 1),
      system_prompt: normalizeOptionalText(valueRecord.system_prompt),
      user_prompt: normalizeOptionalText(valueRecord.user_prompt),
    };
  }

  return profiles;
}

function findBestMatchingItemIndex(
  projectImage: BilddatenProjectImage,
  items: ImageItem[],
  usedItemIndexes: Set<number>
): number {
  const candidates = buildProjectImageCandidates(projectImage);

  const matchLevels: Array<(item: ImageItem) => boolean> = [
    (item) => matchesAnyCandidate(candidates.identifiers, buildItemIdentifiers(item)),
    (item) => matchesRelativePath(projectImage, item),
    (item) => matchesFileNameAndMeta(projectImage, item),
    (item) => matchesFileNameAndSize(projectImage, item),
    (item) => matchesFileName(projectImage, item),
  ];

  for (const matcher of matchLevels) {
    for (let index = 0; index < items.length; index += 1) {
      if (usedItemIndexes.has(index)) {
        continue;
      }

      if (matcher(items[index])) {
        return index;
      }
    }
  }

  return -1;
}

function buildProjectImageCandidates(projectImage: BilddatenProjectImage): {
  identifiers: string[];
} {
  const identifiers = [
    projectImage.image_key,
    projectImage.image_id,
    projectImage.image_hash,
    projectImage.key,
    projectImage.id,
    projectImage.full_path,
    projectImage.relative_path,
    projectImage.filename,
    projectImage.fileName,
    projectImage.image_number,
    projectImage.location,
  ]
    .map((value) => normalizeString(value))
    .filter((value) => value.length > 0);

  return { identifiers };
}

function buildItemIdentifiers(item: ImageItem): string[] {
  return [item.key, item.hash, item.id, item.fullPath, item.relativePath, item.name]
    .map((value) => normalizeString(value))
    .filter((value) => value.length > 0);
}

function matchesAnyCandidate(candidates: string[], itemIdentifiers: string[]): boolean {
  if (candidates.length === 0 || itemIdentifiers.length === 0) {
    return false;
  }

  const normalizedCandidates = candidates.map((value) => normalizeIdentifier(value));
  const normalizedItemIdentifiers = itemIdentifiers.map((value) => normalizeIdentifier(value));

  return normalizedCandidates.some((candidate) => normalizedItemIdentifiers.includes(candidate));
}

function matchesRelativePath(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  const projectRelativePath = normalizePath(
    projectImage.relative_path || projectImage.full_path || projectImage.key || projectImage.id
  );
  const itemRelativePath = normalizePath(
    item.fullPath || item.relativePath || item.key || item.name
  );

  return Boolean(
    projectRelativePath && itemRelativePath && projectRelativePath === itemRelativePath
  );
}

function matchesFileNameAndMeta(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  return (
    matchesFileName(projectImage, item) &&
    normalizeNumber(projectImage.size) === normalizeNumber(item.size) &&
    normalizeNumber(projectImage.lastModified) === normalizeNumber(item.lastModified)
  );
}

function matchesFileNameAndSize(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  return (
    matchesFileName(projectImage, item) &&
    normalizeNumber(projectImage.size) === normalizeNumber(item.size)
  );
}

function matchesFileName(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  const projectName = normalizeName(
    projectImage.filename ||
      projectImage.fileName ||
      projectImage.relative_path ||
      projectImage.full_path ||
      projectImage.key ||
      projectImage.id
  );
  const itemName = normalizeName(item.name || item.relativePath || item.fullPath || item.key);

  return Boolean(projectName && itemName && projectName === itemName);
}

function createDefaultDocumentSettings(): ProjectDocumentSettings {
  return {
    basis: DEFAULT_DOCUMENT_BASIS,
    template_path: "",
    start_page: 2,
  };
}

function createDefaultOutputSettings(): ProjectOutputSettings {
  return {
    images_per_page: 6,
    layout_images_per_page: 6,
    sort_columns_per_row: 5,
    sort_card_size: "Mittel",
    image_management_filter: "Alle",
    caption_font_size: 8,
    compression: "Standard",
    output_basename: DEFAULT_OUTPUT_BASENAME,
    sort_mode: "custom",
  };
}

function createDefaultPreviewSettings(): ProjectPreviewSettings {
  return {
    load_scope: "25",
    quality: "Standard",
    mode: "Einzelseite",
    zoom_percent: 100,
    preview_size_px: DEFAULT_PREVIEW_SIZE_PX,
  };
}

function createDefaultImportSettings(): ProjectImportSettings {
  return {
    use_werkhaus_json: false,
  };
}

function createDefaultCaptionSettings(): ProjectCaptionsSettings {
  return {
    show_image_number: true,
    show_filename: true,
    show_date: true,
    show_time: false,
    show_caption: true,
  };
}

function createDefaultUiSettings(): ProjectUiSettings {
  return {
    right_box_caption_layout: true,
    right_box_output: true,
    right_box_word_document: true,
    right_box_ai_api: true,
    left_box_selection_open: false,
    right_box_creation: true,
    right_box_output_caption_layout: true,
    right_box_folder: false,
    left_box_list_open: false,
    right_box_analysis_profiles: true,
    right_box_project_filename: true,
    sort_mode: "custom",
    preview_size_px: 120,
    insert_size_cm: 10,
    show_info: false,
    show_captions: false,
  };
}

function createDefaultAnalysisProfiles(): Record<string, ProjectAnalysisProfile> {
  return cloneAnalysisProfiles(DEFAULT_ANALYSIS_PROFILES as Record<string, ProjectAnalysisProfile>);
}

function createDefaultAnalysisUi(): Record<string, unknown> {
  return {
    ai_api_provider: "OpenAI",
    ai_vision_model: "GPT-4o-mini",
    analysis_profile_label: "Allgemeine Analyse",
  };
}

function buildCollapsedSectionSnapshot(collapsedSections: string[]): Partial<ProjectUiSettings> {
  const collapsed = new Set(
    collapsedSections.filter((key) => typeof key === "string" && key.length > 0)
  );
  return {
    left_box_list_open: !collapsed.has("list"),
    left_box_selection_open: !collapsed.has("view"),
    right_box_folder: !collapsed.has("import"),
    right_box_output: !collapsed.has("data"),
  };
}

function buildCurrentFormatWarnings(raw: Record<string, unknown>): string[] {
  const warnings: string[] = [];

  const version = normalizeNumber(raw.version, PROJECT_FORMAT_VERSION);
  if (version !== PROJECT_FORMAT_VERSION) {
    warnings.push(`Projektversion ${version} erkannt.`);
  }

  const formatVersion = normalizeNumber(raw.project_format_version, PROJECT_FORMAT_VERSION);
  if (formatVersion !== PROJECT_FORMAT_VERSION) {
    warnings.push(`Projektformat ${formatVersion} erkannt.`);
  }

  return warnings;
}

function isCurrentProjectFormat(raw: Record<string, unknown>): boolean {
  return normalizeString(raw.schema) === PROJECT_SCHEMA;
}

function isLegacyProjectFormat(raw: Record<string, unknown>): boolean {
  return (
    normalizeString(raw.projectType) === LEGACY_PROJECT_TYPE &&
    Number(raw.schemaVersion || LEGACY_PROJECT_SCHEMA_VERSION) >= LEGACY_PROJECT_SCHEMA_VERSION
  );
}

function normalizeProjectVersion(value: unknown, fallback: number): number {
  return normalizeNumber(value, fallback);
}

function normalizeRecord(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }

  return { ...(value as Record<string, unknown>) };
}

function cloneRecord(value: Record<string, unknown>): Record<string, unknown> {
  return { ...(value || {}) };
}

function normalizeBoolean(value: unknown, fallback: boolean): boolean {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "true") return true;
    if (normalized === "false") return false;
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  return fallback;
}

function normalizeNumber(value: unknown, fallback = 0): number {
  const numeric = typeof value === "number" ? value : Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function normalizeString(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function normalizeOptionalText(value: unknown): string | undefined {
  const text = normalizeString(value);
  return text.length > 0 ? text : undefined;
}

function normalizeName(value: unknown): string {
  return normalizeString(value).toLowerCase();
}

function normalizePath(value: unknown): string {
  return normalizeString(value).replace(/\\/g, "/").toLowerCase();
}

function normalizeIdentifier(value: string): string {
  return value.replace(/\\/g, "/").trim().toLowerCase();
}

function normalizeSortMode(value: string): SortMode {
  if (value === "exifDate" || value === "name" || value === "custom") {
    return value;
  }

  return "custom";
}

function sanitizeFileName(value: string): string {
  return value
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, " ")
    .trim();
}

function formatLocalTimestamp(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  const second = String(date.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day}T${hour}:${minute}:${second}`;
}
