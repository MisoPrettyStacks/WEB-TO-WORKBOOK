import { useEffect, useMemo, useState } from "react";
import type { ExtractionCell, ExtractionResult, ExtractionSheet } from "@workspace/api-client-react";
import { useExtractWebpage, useHealthCheck } from "@workspace/api-client-react";
import {
  AlertCircle,
  Check,
  CheckCircle2,
  ChevronDown,
  CircleHelp,
  Clipboard,
  Code2,
  Database,
  Download,
  ExternalLink,
  FileText,
  FilePlus2,
  Link2,
  Loader2,
  MoreHorizontal,
  PanelRight,
  Plus,
  RotateCcw,
  Save,
  Sparkles,
  Table2,
  Trash2,
  WandSparkles,
} from "lucide-react";

type Cell = ExtractionCell | string;
type Notice = { kind: "success" | "error" | "info"; message: string } | null;

const STORAGE_KEY = "weave-local-workbook";

const starterUrls = [
  { label: "A product comparison", url: "https://www.nytimes.com/wirecutter/reviews/best-laptops/" },
  { label: "A public directory", url: "https://www.faa.gov/airports/airport_safety/airportdata_5010" },
  { label: "An article with lists", url: "https://en.wikipedia.org/wiki/List_of_largest_companies_by_revenue" },
];

const blankWorkbook: ExtractionResult = {
  sourceUrl: "",
  title: "",
  description: "",
  fetchedAt: "",
  pageType: "",
  sheets: [],
};

function readSavedWorkbook(): ExtractionResult | null {
  try {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    return saved ? (JSON.parse(saved) as ExtractionResult) : null;
  } catch {
    return null;
  }
}

function formatDate(value: string) {
  if (!value) return "Not yet fetched";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("en", { month: "short", day: "numeric", year: "numeric" }).format(date);
}

function confidenceLabel(confidence: number) {
  if (confidence >= 0.86) return "High confidence";
  if (confidence >= 0.64) return "Good confidence";
  return "Review suggested";
}

function cellText(cell: Cell) {
  if (cell === null || cell === undefined) return "";
  return String(cell);
}

function isPastedSource(source: string) {
  return source.startsWith("pasted-content://");
}

function xmlEscape(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function columnName(index: number) {
  let name = "";
  let current = index + 1;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }
  return name;
}

function worksheetXml(sheet: ExtractionSheet) {
  const rows = [sheet.headers, ...sheet.rows];
  const rowXml = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((cell, cellIndex) => {
          const value = cellText(cell);
          const reference = `${columnName(cellIndex)}${rowIndex + 1}`;
          if (typeof cell === "number") {
            return `<c r="${reference}"><v>${cell}</v></c>`;
          }
          if (typeof cell === "boolean") {
            return `<c r="${reference}" t="b"><v>${cell ? 1 : 0}</v></c>`;
          }
          return `<c r="${reference}" t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowXml}</sheetData></worksheet>`;
}

function crc32(bytes: Uint8Array) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function createXlsx(workbook: ExtractionResult) {
  const encoder = new TextEncoder();
  const files = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${workbook.sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${workbook.sheets.map((sheet, index) => `<sheet name="${xmlEscape(sheet.name.slice(0, 31) || `Sheet ${index + 1}`)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${workbook.sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}</Relationships>`],
    ["xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Manrope"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`],
    ...workbook.sheets.map((sheet, index) => [`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet)]),
  ].map(([name, content]) => ({ name, bytes: encoder.encode(content) }));

  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;
  const write16 = (view: DataView, position: number, value: number) => view.setUint16(position, value, true);
  const write32 = (view: DataView, position: number, value: number) => view.setUint32(position, value, true);

  for (const file of files) {
    const nameBytes = encoder.encode(file.name);
    const checksum = crc32(file.bytes);
    const local = new Uint8Array(30 + nameBytes.length + file.bytes.length);
    const localView = new DataView(local.buffer);
    write32(localView, 0, 0x04034b50);
    write16(localView, 4, 20);
    write16(localView, 6, 0);
    write16(localView, 8, 0);
    write16(localView, 10, 0);
    write16(localView, 12, 0);
    write32(localView, 14, checksum);
    write32(localView, 18, file.bytes.length);
    write32(localView, 22, file.bytes.length);
    write16(localView, 26, nameBytes.length);
    write16(localView, 28, 0);
    local.set(nameBytes, 30);
    local.set(file.bytes, 30 + nameBytes.length);
    localParts.push(local);

    const central = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(central.buffer);
    write32(centralView, 0, 0x02014b50);
    write16(centralView, 4, 20);
    write16(centralView, 6, 20);
    write16(centralView, 8, 0);
    write16(centralView, 10, 0);
    write16(centralView, 12, 0);
    write16(centralView, 14, 0);
    write32(centralView, 16, checksum);
    write32(centralView, 20, file.bytes.length);
    write32(centralView, 24, file.bytes.length);
    write16(centralView, 28, nameBytes.length);
    write16(centralView, 30, 0);
    write16(centralView, 32, 0);
    write16(centralView, 34, 0);
    write16(centralView, 36, 0);
    write32(centralView, 38, 0);
    write32(centralView, 42, offset);
    central.set(nameBytes, 46);
    centralParts.push(central);
    offset += local.length;
  }

  const centralOffset = offset;
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  write32(endView, 0, 0x06054b50);
  write16(endView, 8, files.length);
  write16(endView, 10, files.length);
  write32(endView, 12, centralSize);
  write32(endView, 16, centralOffset);
  write16(endView, 20, 0);
  const parts = [...localParts, ...centralParts, end];
  const output = new Uint8Array(parts.reduce((sum, part) => sum + part.byteLength, 0));
  let cursor = 0;
  for (const part of parts) {
    output.set(part, cursor);
    cursor += part.byteLength;
  }
  return new Blob([output.buffer as ArrayBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function makeFileName(title: string) {
  const safe = title.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase();
  return `${safe || "weave-workbook"}.xlsx`;
}

function EmptyWorkspace({ onUseStarter }: { onUseStarter: (url: string) => void }) {
  return (
    <section className="weave-rise weave-delay-2 relative overflow-hidden rounded-[1.35rem] border border-border bg-card p-6 shadow-sm sm:p-10" data-testid="empty-workspace">
      <div className="absolute -right-20 -top-24 h-72 w-72 rounded-full border border-accent/30 bg-accent/10" />
      <div className="absolute right-20 top-12 h-28 w-28 rounded-full border border-primary/10" />
      <div className="relative max-w-2xl">
        <div className="mb-5 flex h-12 w-12 items-center justify-center rounded-2xl bg-primary text-accent shadow-lg shadow-primary/10">
          <WandSparkles size={22} strokeWidth={1.7} />
        </div>
        <p className="mb-3 font-mono text-[11px] uppercase tracking-[0.2em] text-muted-foreground">A quieter way to research</p>
        <h2 className="max-w-xl font-serif text-4xl leading-[0.98] tracking-[-0.035em] text-foreground sm:text-6xl">
          Start with a page. Leave with a workbook.
        </h2>
        <p className="mt-5 max-w-lg text-sm leading-6 text-muted-foreground sm:text-base">
          Weave finds the useful shape hiding in a public webpage, then gives you a clean place to question, refine, and carry it forward.
        </p>
        <div className="mt-8 flex flex-wrap gap-2">
          {starterUrls.map((starter) => (
            <button
              key={starter.url}
              type="button"
              onClick={() => onUseStarter(starter.url)}
              className="group flex items-center gap-2 rounded-full border border-border bg-background/75 px-3.5 py-2 text-xs font-semibold text-foreground transition hover:-translate-y-0.5 hover:border-accent hover:bg-accent/20"
              data-testid={`button-starter-${starter.label.toLowerCase().replaceAll(" ", "-")}`}
            >
              <Link2 size={13} className="text-muted-foreground transition group-hover:text-foreground" />
              {starter.label}
            </button>
          ))}
        </div>
      </div>
      <div className="pointer-events-none absolute bottom-8 right-10 hidden w-72 rotate-[-7deg] lg:block">
        <div className="rounded-lg border border-border/80 bg-background/85 p-3 shadow-xl shadow-primary/5 backdrop-blur">
          <div className="mb-3 flex items-center justify-between border-b border-border pb-2">
            <span className="font-mono text-[9px] uppercase tracking-[0.15em] text-muted-foreground">suggested sheet</span>
            <Check size={13} className="text-accent-foreground" />
          </div>
          <div className="grid grid-cols-3 gap-px overflow-hidden rounded border border-border bg-border">
            {["Name", "Type", "Signal", "Orchid", "Article", "Strong", "Mina", "Book", "Medium"].map((cell) => (
              <div key={cell} className="bg-card px-2 py-2 text-[9px] text-foreground/75">{cell}</div>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}

function SkeletonWorkspace() {
  return (
    <div className="space-y-4" data-testid="loading-extraction">
      <div className="h-28 animate-pulse rounded-[1.35rem] border border-border bg-card/70" />
      <div className="grid gap-4 lg:grid-cols-[1fr_260px]">
        <div className="h-80 animate-pulse rounded-[1.35rem] border border-border bg-card/70" />
        <div className="h-80 animate-pulse rounded-[1.35rem] border border-border bg-card/70" />
      </div>
    </div>
  );
}

export default function Workspace() {
  const extract = useExtractWebpage();
  const health = useHealthCheck();
  const [url, setUrl] = useState("");
  const [inputMode, setInputMode] = useState<"url" | "text">("url");
  const [pastedText, setPastedText] = useState("");
  const [maxRows, setMaxRows] = useState(50);
  const [workbook, setWorkbook] = useState<ExtractionResult>(() => readSavedWorkbook() ?? blankWorkbook);
  const [activeSheetId, setActiveSheetId] = useState("");
  const [selectedCell, setSelectedCell] = useState("0:0");
  const [notice, setNotice] = useState<Notice>(null);
  const [saved, setSaved] = useState(Boolean(readSavedWorkbook()));
  const [appendMode, setAppendMode] = useState(false);

  const activeSheet = useMemo(
    () => workbook.sheets.find((sheet) => sheet.id === activeSheetId) ?? workbook.sheets[0],
    [activeSheetId, workbook.sheets],
  );

  useEffect(() => {
    if (workbook.sourceUrl.startsWith("http") && !url) setUrl(workbook.sourceUrl);
  }, [workbook.sourceUrl, url]);

  useEffect(() => {
    if (!activeSheetId && workbook.sheets[0]) setActiveSheetId(workbook.sheets[0].id);
    if (activeSheetId && workbook.sheets.length > 0 && !workbook.sheets.some((sheet) => sheet.id === activeSheetId)) {
      setActiveSheetId(workbook.sheets[0].id);
    }
  }, [activeSheetId, workbook.sheets]);

  const updateActiveSheet = (update: (sheet: ExtractionSheet) => ExtractionSheet) => {
    if (!activeSheet) return;
    setWorkbook((current) => ({
      ...current,
      sheets: current.sheets.map((sheet) => (sheet.id === activeSheet.id ? update(sheet) : sheet)),
    }));
    setSaved(false);
  };

  const updateCell = (rowIndex: number, cellIndex: number, value: string) => {
    updateActiveSheet((sheet) => {
      const rows = sheet.rows.map((row) => [...row]);
      if (!rows[rowIndex]) rows[rowIndex] = [];
      rows[rowIndex][cellIndex] = value;
      return { ...sheet, rows };
    });
  };

  const updateHeader = (index: number, value: string) => {
    updateActiveSheet((sheet) => {
      const headers = [...sheet.headers];
      headers[index] = value;
      return { ...sheet, headers };
    });
  };

  const handleExtract = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const normalized = url.trim();
    const normalizedText = pastedText.trim();
    if (inputMode === "url" && !/^https?:\/\/.+/i.test(normalized)) {
      setNotice({ kind: "error", message: "Add a full public URL beginning with http:// or https://." });
      return;
    }
    if (inputMode === "text" && normalizedText.length < 20) {
      setNotice({ kind: "error", message: "Paste at least 20 characters of page text so Weave has enough context to structure it." });
      return;
    }
    setNotice({ kind: "info", message: inputMode === "url" ? "Reading the page and looking for its useful shape…" : "Reading the pasted page and looking for its useful shape…" });
    extract.mutate(
      {
        data: inputMode === "url"
          ? { url: normalized, maxRows }
          : { text: normalizedText, maxRows },
      },
      {
        onSuccess: (result) => {
          const appendedSheets = result.sheets.map((sheet, index) => ({
            ...sheet,
            id: `sheet-${workbook.sheets.length + index + 1}-${sheet.id}`,
          }));
          const nextWorkbook = appendMode && workbook.sheets.length
            ? {
                ...result,
                title: workbook.title || result.title,
                sheets: [...workbook.sheets, ...appendedSheets],
              }
            : result;
          setWorkbook(nextWorkbook);
          setActiveSheetId(appendMode ? appendedSheets[0]?.id ?? "" : result.sheets[0]?.id ?? "");
          setSelectedCell("0:0");
          setSaved(false);
          setAppendMode(false);
          setNotice({
            kind: "success",
            message: appendMode
              ? `${result.sheets.length} new sheet${result.sheets.length === 1 ? "" : "s"} added to this workbook.`
              : `${result.sheets.length} suggested sheet${result.sheets.length === 1 ? "" : "s"} ready to review.`,
          });
        },
        onError: (error) => {
          const typedError = error as {
            data?: { error?: string } | null;
            error?: string;
            message?: string;
          };
          const message = typedError.data?.error ?? typedError.error ?? typedError.message ?? "We couldn't read that page. Check the URL and try again.";
          setNotice({ kind: "error", message });
        },
      },
    );
  };

  const saveLocally = () => {
    if (!workbook.sourceUrl) return;
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook));
    setSaved(true);
    setNotice({ kind: "success", message: "Saved in this browser. Your workbook will be here when you return." });
  };

  const newWorkbook = () => {
    setWorkbook(blankWorkbook);
    setUrl("");
    setPastedText("");
    setInputMode("url");
    setActiveSheetId("");
    setSelectedCell("0:0");
    setSaved(false);
    setAppendMode(false);
    setNotice({ kind: "info", message: "A fresh workbook is ready." });
  };

  const addSheet = () => {
    setAppendMode(true);
    setUrl("");
    setPastedText("");
    setNotice({ kind: "info", message: "Paste another public page above. Its suggested sheets will be added here." });
    window.setTimeout(() => document.querySelector<HTMLInputElement>('[data-testid="input-source-url"]')?.focus(), 0);
  };

  const deleteActiveSheet = () => {
    if (!activeSheet || workbook.sheets.length <= 1) {
      setNotice({ kind: "info", message: "Keep at least one sheet in a workbook." });
      return;
    }
    const remaining = workbook.sheets.filter((sheet) => sheet.id !== activeSheet.id);
    setWorkbook((current) => ({ ...current, sheets: remaining }));
    setActiveSheetId(remaining[0].id);
    setSaved(false);
  };

  const addRow = () => {
    updateActiveSheet((sheet) => ({ ...sheet, rows: [...sheet.rows, sheet.headers.map(() => "")] }));
  };

  const downloadWorkbook = () => {
    if (workbook.sheets.length === 0) return;
    const anchor = document.createElement("a");
    anchor.href = URL.createObjectURL(createXlsx(workbook));
    anchor.download = makeFileName(workbook.title);
    anchor.click();
    URL.revokeObjectURL(anchor.href);
    setNotice({ kind: "success", message: "Your XLSX is downloading." });
  };

  const hasWorkbook = workbook.sheets.length > 0;
  const healthStatus = health.isLoading ? "Checking service" : health.isError ? "Local mode" : health.data?.status === "ok" ? "Service ready" : "Service connected";

  return (
    <div className="min-h-[100dvh] bg-background text-foreground">
      <aside className="fixed inset-y-0 left-0 z-20 hidden w-[248px] flex-col border-r border-sidebar-border bg-sidebar px-5 py-6 text-sidebar-foreground md:flex">
        <div className="flex items-center gap-3 px-1">
          <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-sidebar-primary text-sidebar-primary-foreground">
            <span className="font-serif text-xl leading-none">W</span>
          </div>
          <div>
            <p className="font-serif text-[22px] leading-none tracking-[-0.04em]">Weave</p>
            <p className="mt-1 font-mono text-[9px] uppercase tracking-[0.16em] text-sidebar-foreground/50">web to workbook</p>
          </div>
        </div>
        <div className="mt-14">
          <p className="px-1 font-mono text-[10px] uppercase tracking-[0.18em] text-sidebar-foreground/45">Your workspace</p>
          <div className="mt-3 rounded-xl bg-sidebar-accent px-3 py-3">
            <div className="flex items-center gap-3">
              <Table2 size={16} className="text-sidebar-primary" />
              <span className="text-sm font-semibold">Current workbook</span>
            </div>
            <p className="mt-3 truncate pl-7 text-xs text-sidebar-foreground/55">{workbook.title || "Untitled research"}</p>
          </div>
        </div>
        <div className="mt-auto space-y-4">
          <div className="rounded-xl border border-sidebar-border bg-sidebar-accent/50 p-3">
            <div className="flex items-center gap-2 text-xs font-semibold">
              <CircleHelp size={14} className="text-sidebar-primary" />
              A small note
            </div>
            <p className="mt-2 text-[11px] leading-5 text-sidebar-foreground/55">Weave suggests structure. You decide what belongs in the final workbook.</p>
          </div>
          <div className="flex items-center gap-2 px-1 font-mono text-[10px] uppercase tracking-[0.12em] text-sidebar-foreground/45">
            <span className={`h-1.5 w-1.5 rounded-full ${health.isError ? "bg-sidebar-foreground/40" : "bg-sidebar-primary"}`} />
            {healthStatus}
          </div>
        </div>
      </aside>

      <main className="md:pl-[248px]">
        <header className="sticky top-0 z-10 border-b border-border/80 bg-background/90 px-4 py-4 backdrop-blur sm:px-8 lg:px-10">
          <div className="mx-auto flex max-w-[1480px] items-center justify-between gap-4">
            <div className="flex items-center gap-3 md:hidden">
              <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-primary text-accent">
                <span className="font-serif text-lg">W</span>
              </div>
              <span className="font-serif text-xl tracking-[-0.04em]">Weave</span>
            </div>
            <div className="hidden items-center gap-2 md:flex">
              <span className="font-mono text-[10px] uppercase tracking-[0.17em] text-muted-foreground">Research desk</span>
              <ChevronDown size={13} className="text-muted-foreground" />
            </div>
            <div className="flex items-center gap-2">
              <span className="hidden items-center gap-2 rounded-full border border-border bg-card px-3 py-1.5 text-[11px] font-semibold text-muted-foreground sm:flex">
                <span className={`h-1.5 w-1.5 rounded-full ${health.isError ? "bg-muted-foreground" : "bg-emerald-600"}`} />
                {healthStatus}
              </span>
              <a
                href={`${import.meta.env.BASE_URL}weave-source.zip`}
                download
                className="flex items-center gap-2 rounded-lg border border-border bg-card px-3 py-2 text-xs font-semibold transition hover:border-primary/30 hover:bg-muted"
                data-testid="button-download-source"
              >
                <Code2 size={14} />
                <span className="hidden sm:inline">Download source</span>
              </a>
              <button type="button" onClick={newWorkbook} className="flex items-center gap-2 rounded-lg border border-border bg-card px-3 py-2 text-xs font-semibold transition hover:border-primary/30 hover:bg-muted" data-testid="button-new-workbook">
                <FilePlus2 size={14} />
                <span className="hidden sm:inline">New workbook</span>
              </button>
            </div>
          </div>
        </header>

        <div className="mx-auto max-w-[1480px] px-4 py-8 sm:px-8 lg:px-10 lg:py-10">
          <section className="weave-rise flex flex-col justify-between gap-6 lg:flex-row lg:items-end">
            <div>
              <p className="mb-3 font-mono text-[10px] uppercase tracking-[0.2em] text-muted-foreground">Webpage → workbook</p>
              <h1 className="max-w-3xl font-serif text-4xl leading-[0.98] tracking-[-0.045em] sm:text-6xl lg:text-7xl">
                Turn the web into <span className="text-muted-foreground">working knowledge.</span>
              </h1>
              <p className="mt-5 max-w-xl text-sm leading-6 text-muted-foreground sm:text-base">Paste a public page. Weave finds its tables, lists, and signals so you can spend your time thinking instead of copying.</p>
            </div>
            {hasWorkbook && (
              <div className="flex items-center gap-3 pb-1 text-xs text-muted-foreground">
                <Database size={15} />
                <span>{workbook.sheets.length} sheet{workbook.sheets.length === 1 ? "" : "s"} in this workbook</span>
              </div>
            )}
          </section>

          <form onSubmit={handleExtract} className="weave-rise weave-delay-1 mt-8 rounded-[1.15rem] border border-primary/15 bg-primary p-2 shadow-xl shadow-primary/10 sm:p-2.5" data-testid="form-extract">
            <div className="mb-2 flex items-center gap-1 rounded-[0.85rem] bg-primary-foreground/10 p-1">
              <button
                type="button"
                onClick={() => setInputMode("url")}
                className={`flex items-center gap-2 rounded-lg px-3 py-2 text-xs font-bold transition ${inputMode === "url" ? "bg-card text-foreground shadow-sm" : "text-primary-foreground/65 hover:text-primary-foreground"}`}
                aria-pressed={inputMode === "url"}
                data-testid="button-source-url"
              >
                <Link2 size={14} /> Use a URL
              </button>
              <button
                type="button"
                onClick={() => setInputMode("text")}
                className={`flex items-center gap-2 rounded-lg px-3 py-2 text-xs font-bold transition ${inputMode === "text" ? "bg-card text-foreground shadow-sm" : "text-primary-foreground/65 hover:text-primary-foreground"}`}
                aria-pressed={inputMode === "text"}
                data-testid="button-source-text"
              >
                <FileText size={14} /> Paste page text
              </button>
              <span className="ml-auto hidden pr-3 text-[10px] text-primary-foreground/50 sm:block">
                {inputMode === "url" ? "Best for public pages" : "Best when a page blocks access"}
              </span>
            </div>
            <div className="flex flex-col gap-2 lg:flex-row">
              <div className={`flex min-w-0 flex-1 gap-3 rounded-[0.85rem] bg-primary-foreground/10 px-4 py-3 text-primary-foreground ${inputMode === "url" ? "items-center" : "items-start"}`}>
                {inputMode === "url" ? <Link2 size={18} className="mt-0.5 shrink-0 text-accent" /> : <FileText size={18} className="mt-0.5 shrink-0 text-accent" />}
                {inputMode === "url" ? (
                  <input
                    value={url}
                    onChange={(event) => setUrl(event.target.value)}
                    placeholder="https://a public webpage you want to work with"
                    className="min-w-0 flex-1 bg-transparent text-sm outline-none placeholder:text-primary-foreground/45"
                    data-testid="input-source-url"
                    aria-label="Public webpage URL"
                  />
                ) : (
                  <textarea
                    value={pastedText}
                    onChange={(event) => setPastedText(event.target.value)}
                    placeholder="Paste the page text here — tables, lists, product details, or the full visible page"
                    className="min-h-28 min-w-0 flex-1 resize-y bg-transparent text-sm leading-6 outline-none placeholder:text-primary-foreground/45"
                    data-testid="input-source-text"
                    aria-label="Pasted webpage text"
                  />
                )}
                <div className="hidden items-center gap-1.5 border-l border-primary-foreground/15 pl-3 text-primary-foreground/50 sm:flex">
                  <span className="font-mono text-[10px]">ROWS</span>
                  <input
                    type="number"
                    min={1}
                    max={500}
                    value={maxRows}
                    onChange={(event) => setMaxRows(Math.max(1, Math.min(500, Number(event.target.value) || 1)))}
                    className="w-12 bg-transparent text-right font-mono text-xs text-primary-foreground outline-none"
                    data-testid="input-max-rows"
                    aria-label="Maximum rows"
                  />
                </div>
              </div>
              <button type="submit" disabled={extract.isPending} className="flex items-center justify-center gap-2 rounded-[0.85rem] bg-accent px-5 py-3 text-sm font-extrabold text-accent-foreground transition hover:brightness-95 disabled:cursor-wait disabled:opacity-70" data-testid="button-extract">
                {extract.isPending ? <Loader2 size={17} className="animate-spin" /> : <Sparkles size={17} />}
                {extract.isPending ? "Reading page" : appendMode ? "Add to workbook" : "Extract data"}
              </button>
            </div>
          </form>

          {notice && (
            <div className={`mt-4 flex items-start gap-2 rounded-lg border px-3 py-2.5 text-xs ${notice.kind === "error" ? "border-destructive/25 bg-destructive/5 text-destructive" : notice.kind === "success" ? "border-emerald-700/20 bg-emerald-700/5 text-emerald-800 dark:text-emerald-300" : "border-accent/40 bg-accent/10 text-foreground"}`} role="status" data-testid={`status-${notice.kind}`}>
              {notice.kind === "error" ? <AlertCircle size={15} className="mt-0.5 shrink-0" /> : notice.kind === "success" ? <CheckCircle2 size={15} className="mt-0.5 shrink-0" /> : <Loader2 size={15} className="mt-0.5 shrink-0 animate-spin" />}
              <span className="flex-1">{notice.message}</span>
              {notice.kind === "error" && notice.message.toLowerCase().includes("paste page text") && (
                <button
                  type="button"
                  onClick={() => {
                    setInputMode("text");
                    setPastedText("");
                    window.setTimeout(() => document.querySelector<HTMLTextAreaElement>('[data-testid="input-source-text"]')?.focus(), 0);
                  }}
                  className="shrink-0 font-bold underline underline-offset-2 hover:no-underline"
                  data-testid="button-switch-to-paste"
                >
                  Paste instead
                </button>
              )}
            </div>
          )}

          <div className="mt-8">
            {extract.isPending ? (
              <SkeletonWorkspace />
            ) : hasWorkbook && activeSheet ? (
              <section className="weave-rise weave-delay-2 space-y-4" data-testid="workbook-workspace">
                <div className="flex flex-col justify-between gap-4 rounded-[1.15rem] border border-border bg-card p-5 sm:flex-row sm:items-start sm:p-6">
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-2">
                      <h2 className="truncate font-serif text-2xl tracking-[-0.025em]" data-testid="text-workbook-title">{workbook.title || "Untitled workbook"}</h2>
                      <span className="rounded-full bg-accent/25 px-2.5 py-1 font-mono text-[10px] uppercase tracking-[0.1em] text-accent-foreground">{workbook.pageType || "webpage"}</span>
                    </div>
                    <p className="mt-2 max-w-3xl text-sm leading-5 text-muted-foreground">{workbook.description || "A workbook assembled from the source page."}</p>
                    <div className="mt-3 flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px] text-muted-foreground">
                      {isPastedSource(workbook.sourceUrl) ? (
                        <span className="inline-flex max-w-full items-center gap-1.5 truncate font-mono" data-testid="text-pasted-source">
                          <FileText size={12} /> Pasted page text
                        </span>
                      ) : (
                        <a href={workbook.sourceUrl} target="_blank" rel="noreferrer" className="inline-flex max-w-full items-center gap-1.5 truncate font-mono hover:text-foreground" data-testid="link-source">
                          <ExternalLink size={12} /> {workbook.sourceUrl}
                        </a>
                      )}
                      <span>Fetched {formatDate(workbook.fetchedAt)}</span>
                    </div>
                  </div>
                  <div className="flex shrink-0 items-center gap-2">
                    <button type="button" onClick={saveLocally} className="flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-xs font-semibold transition hover:border-primary/30 hover:bg-muted" data-testid="button-save-local">
                      {saved ? <Check size={14} className="text-emerald-700" /> : <Save size={14} />}
                      {saved ? "Saved locally" : "Save locally"}
                    </button>
                    <button type="button" onClick={downloadWorkbook} className="flex items-center gap-2 rounded-lg bg-primary px-3.5 py-2 text-xs font-semibold text-primary-foreground transition hover:bg-primary/90" data-testid="button-download-xlsx">
                      <Download size={14} /> XLSX
                    </button>
                  </div>
                </div>

                <div className="flex gap-2 overflow-x-auto border-b border-border pb-0 weave-scrollbar" data-testid="sheet-tabs">
                  {workbook.sheets.map((sheet) => (
                    <button
                      key={sheet.id}
                      type="button"
                      onClick={() => { setActiveSheetId(sheet.id); setSelectedCell("0:0"); }}
                      className={`group flex shrink-0 items-center gap-2 border-b-2 px-3 py-3 text-xs font-bold transition ${activeSheet.id === sheet.id ? "border-primary text-foreground" : "border-transparent text-muted-foreground hover:text-foreground"}`}
                      data-testid={`button-sheet-${sheet.id}`}
                    >
                      <Table2 size={14} className={activeSheet.id === sheet.id ? "text-accent-foreground" : ""} />
                      {sheet.name || "Untitled sheet"}
                      <span className="font-mono text-[10px] font-normal opacity-55">{sheet.rows.length}</span>
                    </button>
                  ))}
                   <button type="button" onClick={addSheet} className={`flex shrink-0 items-center gap-1.5 px-3 py-3 text-xs font-semibold transition ${appendMode ? "text-foreground" : "text-muted-foreground hover:text-foreground"}`} data-testid="button-add-sheet">
                     <Plus size={14} /> {appendMode ? "Paste next page" : "Add tab"}
                  </button>
                </div>

                <div className="grid min-w-0 gap-4 xl:grid-cols-[minmax(0,1fr)_280px]">
                  <div className="min-w-0 overflow-hidden rounded-[1.15rem] border border-border bg-card">
                    <div className="flex flex-wrap items-center justify-between gap-3 border-b border-border px-4 py-3 sm:px-5">
                      <div className="flex items-center gap-2">
                        <span className="h-2 w-2 rounded-full bg-accent" />
                        <span className="text-xs font-bold">{activeSheet.name || "Untitled sheet"}</span>
                        <span className="text-xs text-muted-foreground">· {activeSheet.rows.length} rows</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <button type="button" onClick={addRow} className="flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-[11px] font-semibold text-muted-foreground transition hover:bg-muted hover:text-foreground" data-testid="button-add-row"><Plus size={13} /> Row</button>
                        <button type="button" onClick={deleteActiveSheet} className="rounded-md p-1.5 text-muted-foreground transition hover:bg-destructive/10 hover:text-destructive" aria-label="Delete current sheet" data-testid="button-delete-sheet"><Trash2 size={14} /></button>
                        <button type="button" onClick={() => setNotice({ kind: "info", message: "Sheet options are available inline: rename the tab, add rows, or remove it from this toolbar." })} className="rounded-md p-1.5 text-muted-foreground transition hover:bg-muted hover:text-foreground" aria-label="More sheet options" data-testid="button-sheet-options"><MoreHorizontal size={15} /></button>
                      </div>
                    </div>
                    <div className="weave-scrollbar max-h-[560px] overflow-auto">
                      <table className="w-full min-w-[700px] border-collapse text-left" data-testid="table-workbook">
                        <thead>
                          <tr className="bg-muted/45">
                            <th className="sticky left-0 z-[1] w-12 border-b border-r border-border bg-muted/75 px-3 py-3 text-center font-mono text-[10px] font-normal text-muted-foreground">#</th>
                            {activeSheet.headers.map((header, index) => (
                              <th key={`${header}-${index}`} className="min-w-[150px] border-b border-border p-0 font-mono text-[10px] uppercase tracking-[0.08em] text-muted-foreground">
                                <input value={header} onChange={(event) => updateHeader(index, event.target.value)} className="cell-input font-mono text-[10px] uppercase tracking-[0.08em]" aria-label={`Edit header ${index + 1}`} data-testid={`input-header-${index}`} />
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {activeSheet.rows.length === 0 ? (
                            <tr><td colSpan={Math.max(activeSheet.headers.length + 1, 2)} className="px-5 py-12 text-center text-sm text-muted-foreground">This sheet is empty. Add a row to start shaping it.</td></tr>
                          ) : activeSheet.rows.map((row, rowIndex) => (
                            <tr key={`row-${rowIndex}`} className="group transition hover:bg-accent/5" data-testid={`row-workbook-${rowIndex}`}>
                              <td className="sticky left-0 z-[1] border-r border-b border-border bg-card px-3 py-2 text-center font-mono text-[10px] text-muted-foreground group-hover:bg-accent/5">{rowIndex + 1}</td>
                              {activeSheet.headers.map((_, cellIndex) => {
                                const key = `${rowIndex}:${cellIndex}`;
                                return (
                                  <td key={key} className={`border-b border-border p-0 ${selectedCell === key ? "bg-accent/10" : ""}`}>
                                    <input
                                      value={cellText(row[cellIndex])}
                                      onFocus={() => setSelectedCell(key)}
                                      onChange={(event) => updateCell(rowIndex, cellIndex, event.target.value)}
                                      className="cell-input text-sm"
                                      aria-label={`Edit row ${rowIndex + 1}, column ${cellIndex + 1}`}
                                      data-testid={`input-cell-${rowIndex}-${cellIndex}`}
                                    />
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="flex items-center gap-2 border-t border-border bg-muted/20 px-4 py-3 text-[11px] text-muted-foreground sm:px-5">
                      <Clipboard size={13} />
                      <span>Click any cell to edit. Headers are yours to rename.</span>
                    </div>
                  </div>

                  <aside className="rounded-[1.15rem] border border-border bg-card p-5" data-testid="sheet-inspector">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <PanelRight size={15} className="text-muted-foreground" />
                        <p className="font-mono text-[10px] uppercase tracking-[0.17em] text-muted-foreground">Sheet details</p>
                      </div>
                      <button type="button" className="rounded-md p-1 text-muted-foreground hover:bg-muted" aria-label="Copy sheet source" data-testid="button-copy-source" onClick={() => navigator.clipboard?.writeText(activeSheet.source)}><Clipboard size={14} /></button>
                    </div>
                    <label className="mt-6 block">
                      <span className="mb-2 block text-[11px] font-bold uppercase tracking-[0.1em] text-muted-foreground">Sheet name</span>
                      <input value={activeSheet.name} onChange={(event) => updateActiveSheet((sheet) => ({ ...sheet, name: event.target.value }))} className="w-full rounded-lg border border-input bg-background px-3 py-2 text-sm outline-none transition focus:border-accent focus:ring-2 focus:ring-accent/20" data-testid="input-sheet-name" />
                    </label>
                    <div className="mt-5 border-t border-border pt-5">
                      <p className="text-xs leading-5 text-muted-foreground">{activeSheet.description || "No description was supplied for this sheet."}</p>
                    </div>
                    <div className="mt-5 space-y-3 border-t border-border pt-5">
                      <div className="flex items-center justify-between text-xs"><span className="text-muted-foreground">Confidence</span><span className="font-semibold">{Math.round(activeSheet.confidence * 100)}%</span></div>
                      <div className="h-1.5 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full bg-accent transition-all" style={{ width: `${Math.max(4, activeSheet.confidence * 100)}%` }} /></div>
                      <p className="font-mono text-[10px] uppercase tracking-[0.11em] text-muted-foreground">{confidenceLabel(activeSheet.confidence)}</p>
                    </div>
                    <div className="mt-5 border-t border-border pt-5">
                      <p className="mb-2 text-[11px] font-bold uppercase tracking-[0.1em] text-muted-foreground">Source trail</p>
                      {isPastedSource(activeSheet.source) ? (
                        <span className="flex items-start gap-2 text-xs leading-5 text-foreground/70" data-testid="text-sheet-pasted-source"><FileText size={13} className="mt-1 shrink-0" /> <span>Page text pasted directly</span></span>
                      ) : (
                        <a href={activeSheet.source} target="_blank" rel="noreferrer" className="flex items-start gap-2 text-xs leading-5 text-foreground/70 hover:text-foreground" data-testid="link-sheet-source"><ExternalLink size={13} className="mt-1 shrink-0" /> <span className="break-all">{activeSheet.source}</span></a>
                      )}
                    </div>
                    <div className="mt-6 rounded-lg bg-accent/15 p-3 text-[11px] leading-5 text-foreground/70">
                      <div className="mb-1 flex items-center gap-1.5 font-bold text-foreground"><Sparkles size={13} /> Suggested, not sacred</div>
                      Keep the shape, rename the language, or start a new tab. This workbook belongs to you.
                    </div>
                  </aside>
                </div>

                <div className="flex flex-col justify-between gap-3 border-t border-border pt-4 text-xs text-muted-foreground sm:flex-row sm:items-center">
                  <span className="flex items-center gap-2"><RotateCcw size={13} /> Changes are local until you save or download.</span>
                  <button type="button" onClick={newWorkbook} className="flex items-center gap-2 self-start font-semibold text-foreground hover:text-muted-foreground sm:self-auto" data-testid="button-start-over"><FilePlus2 size={14} /> Start a new workbook</button>
                </div>
              </section>
            ) : (
              <EmptyWorkspace onUseStarter={setUrl} />
            )}
          </div>
        </div>
      </main>
    </div>
  );
}