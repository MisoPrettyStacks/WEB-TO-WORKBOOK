import { Router, type IRouter } from "express";
import dns from "node:dns/promises";
import net from "node:net";
import { ExtractWebpageBody } from "@workspace/api-zod";

type Cell = string | number | boolean | null;
type Sheet = {
  id: string;
  name: string;
  description: string;
  confidence: number;
  headers: string[];
  rows: Cell[][];
  source: string;
};

const router: IRouter = Router();
const MAX_HTML_BYTES = 5 * 1024 * 1024;
const USER_AGENT =
  "Mozilla/5.0 (compatible; WeaveWorkbook/1.0; +https://replit.com)";

const blockedHosts = new Set(["localhost", "localhost.localdomain"]);

function decodeEntities(value: string): string {
  return value
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code: string) =>
      String.fromCodePoint(Number.parseInt(code, 16)),
    );
}

function cleanText(value: string): string {
  return decodeEntities(
    value
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<svg[\s\S]*?<\/svg>/gi, " ")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim(),
  );
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function tagAttribute(tag: string, attribute: string): string | undefined {
  const match = tag.match(
    new RegExp(`${escapeRegExp(attribute)}\\s*=\\s*["']([^"']+)["']`, "i"),
  );
  return match?.[1];
}

function absoluteUrl(value: string, base: string): string {
  try {
    return new URL(value, base).toString();
  } catch {
    return value;
  }
}

function isPrivateAddress(address: string): boolean {
  if (net.isIP(address) === 4) {
    const octets = address.split(".").map(Number);
    return (
      octets[0] === 10 ||
      octets[0] === 127 ||
      (octets[0] === 169 && octets[1] === 254) ||
      (octets[0] === 172 && octets[1] >= 16 && octets[1] <= 31) ||
      (octets[0] === 192 && octets[1] === 168)
    );
  }
  if (net.isIP(address) === 6) {
    const normalized = address.toLowerCase();
    return (
      normalized === "::1" ||
      normalized.startsWith("fc") ||
      normalized.startsWith("fd") ||
      normalized.startsWith("fe80:")
    );
  }
  return false;
}

async function assertPublicUrl(rawUrl: string): Promise<URL> {
  const url = new URL(rawUrl);
  if (!["http:", "https:"].includes(url.protocol)) {
    throw new Error("Only public http and https pages can be read.");
  }
  const hostname = url.hostname.toLowerCase();
  if (blockedHosts.has(hostname) || isPrivateAddress(hostname)) {
    throw new Error("That address is not available for public page extraction.");
  }
  const records = await dns.lookup(hostname, { all: true });
  if (records.some((record) => isPrivateAddress(record.address))) {
    throw new Error("That address is not available for public page extraction.");
  }
  return url;
}

async function fetchPublicHtml(rawUrl: string): Promise<{ url: URL; html: string }> {
  let url = await assertPublicUrl(rawUrl);
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const response = await fetch(url, {
      headers: {
        "user-agent": USER_AGENT,
        accept: "text/html,application/xhtml+xml",
      },
      redirect: "manual",
      signal: AbortSignal.timeout(15000),
    });

    if ([301, 302, 303, 307, 308].includes(response.status)) {
      const location = response.headers.get("location");
      if (!location) throw new Error("The page returned an invalid redirect.");
      url = await assertPublicUrl(absoluteUrl(location, url.toString()));
      continue;
    }
    if (!response.ok) {
      if (response.status === 403) {
        throw new Error(
          "This website blocked automated access. Try switching to Paste page text instead.",
        );
      }
      if (response.status === 401) {
        throw new Error(
          "This page requires a login or blocked automated access. Try switching to Paste page text instead.",
        );
      }
      if (response.status === 429) {
        throw new Error(
          "This website is rate-limiting automated access. Try again later or use Paste page text instead.",
        );
      }
      throw new Error(`The page returned HTTP ${response.status}.`);
    }
    const bytes = await response.arrayBuffer();
    if (bytes.byteLength > MAX_HTML_BYTES) {
      throw new Error("That page is larger than the 5 MB extraction limit.");
    }
    return { url, html: new TextDecoder().decode(bytes) };
  }
  throw new Error("The page redirected too many times.");
}

function pageTitle(html: string, url: URL): string {
  const title = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i)?.[1];
  return cleanText(title ?? url.hostname) || url.hostname;
}

function pageDescription(html: string): string {
  const match = html.match(
    /<meta[^>]+(?:name|property)=["'](?:description|og:description)["'][^>]+content=["']([^"']*)["'][^>]*>/i,
  );
  return cleanText(match?.[1] ?? "");
}

function parseCellValue(value: string): Cell {
  const text = cleanText(value);
  if (!text) return null;
  if (/^(true|false)$/i.test(text)) return text.toLowerCase() === "true";
  const numeric = text.replace(/[$,%\s,]/g, "");
  if (/^-?\d+(?:\.\d+)?$/.test(numeric) && text.length < 40) {
    return Number(numeric);
  }
  return text;
}

function uniqueHeaders(values: string[], width: number): string[] {
  const seen = new Map<string, number>();
  return Array.from({ length: width }, (_, index) => {
    const base = cleanText(values[index] ?? "") || `Column ${index + 1}`;
    const count = (seen.get(base) ?? 0) + 1;
    seen.set(base, count);
    return count === 1 ? base : `${base} ${count}`;
  });
}

function extractTables(html: string, source: string, maxRows: number): Sheet[] {
  const sheets: Sheet[] = [];
  const tableMatches = html.match(/<table\b[\s\S]*?<\/table>/gi) ?? [];
  tableMatches.forEach((table, index) => {
    const rows: string[][] = [];
    const rowMatches = table.match(/<tr\b[\s\S]*?<\/tr>/gi) ?? [];
    rowMatches.forEach((row) => {
      const cells =
        row.match(/<(?:th|td)\b[^>]*>[\s\S]*?<\/(?:th|td)>/gi)?.map(cleanText) ??
        [];
      if (cells.some(Boolean)) rows.push(cells);
    });
    if (rows.length < 2) return;
    const hasHeadings = /<th\b/i.test(rowMatches[0] ?? "");
    const headerSource = hasHeadings ? rows.shift()! : rows.shift()!;
    const width = Math.max(
      headerSource.length,
      ...rows.map((row) => row.length),
    );
    const cleanRows = rows
      .slice(0, maxRows)
      .map((row) =>
        Array.from({ length: width }, (_, cellIndex) =>
          parseCellValue(row[cellIndex] ?? ""),
        ),
      );
    if (cleanRows.length === 0 || width === 0) return;
    sheets.push({
      id: `table-${index + 1}`,
      name: `Table ${index + 1}`,
      description: hasHeadings
        ? "A structured table detected on the page."
        : "A repeated row pattern detected on the page.",
      confidence: hasHeadings ? 0.98 : 0.86,
      headers: uniqueHeaders(headerSource, width),
      rows: cleanRows,
      source,
    });
  });
  return sheets;
}

function jsonLdBlocks(html: string): unknown[] {
  const blocks = html.match(
    /<script[^>]+type=["']application\/ld\+json["'][^>]*>[\s\S]*?<\/script>/gi,
  );
  const values: unknown[] = [];
  for (const block of blocks ?? []) {
    const raw = block.replace(/<script[^>]*>|<\/script>/gi, "").trim();
    try {
      values.push(JSON.parse(raw.replace(/\/\*[\s\S]*?\*\//g, "")));
    } catch {
      continue;
    }
  }
  return values;
}

function extractStructuredData(html: string, source: string, maxRows: number): Sheet[] {
  const records: Record<string, unknown>[] = [];
  const visit = (value: unknown) => {
    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }
    if (!value || typeof value !== "object") return;
    const record = value as Record<string, unknown>;
    if (Array.isArray(record.itemListElement)) {
      record.itemListElement.forEach(visit);
    }
    const type = String(record["@type"] ?? "");
    if (
      ["Product", "Article", "NewsArticle", "Recipe", "JobPosting", "Event", "Movie"].some(
        (candidate) => type.includes(candidate),
      )
    ) {
      records.push(record);
    }
    if (record.item && typeof record.item === "object") visit(record.item);
    if (record["@graph"]) visit(record["@graph"]);
  };
  jsonLdBlocks(html).forEach(visit);
  if (records.length < 2) return [];

  const keys = Array.from(
    new Set(
      records.flatMap((record) =>
        Object.keys(record).filter(
          (key) =>
            !key.startsWith("@") &&
            ["object", "string", "number", "boolean"].includes(typeof record[key]),
        ),
      ),
    ),
  ).slice(0, 10);
  const rows = records.slice(0, maxRows).map((record) =>
    keys.map((key) => {
      const value = record[key];
      if (value === null || typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
        return value as Cell;
      }
      return typeof value === "object" ? JSON.stringify(value) : String(value);
    }),
  );
  return [
    {
      id: "structured-data",
      name: "Structured data",
      description: "Product, article, recipe, event, or other records published by the page.",
      confidence: 0.93,
      headers: uniqueHeaders(keys, keys.length),
      rows,
      source,
    },
  ];
}

function extractArticles(html: string, source: string, maxRows: number): Sheet[] {
  const blocks = html.match(/<article\b[\s\S]*?<\/article>/gi) ?? [];
  const rows = blocks
    .map((article) => {
      const heading = article.match(/<h[1-6]\b[^>]*>[\s\S]*?<\/h[1-6]>/i)?.[0] ?? "";
      const link = article.match(/<a\b[^>]+href=["']([^"']+)["'][^>]*>/i)?.[1] ?? "";
      const paragraphs = article.match(/<p\b[\s\S]*?<\/p>/gi)?.map(cleanText) ?? [];
      return [cleanText(heading), absoluteUrl(link, source), paragraphs[0] ?? ""];
    })
    .filter(([title]) => title)
    .slice(0, maxRows);
  if (rows.length < 2) return [];
  return [
    {
      id: "articles",
      name: "Articles",
      description: "Repeated article or story cards detected on the page.",
      confidence: 0.78,
      headers: ["Title", "Link", "Summary"],
      rows,
      source,
    },
  ];
}

function splitPastedLine(line: string): string[] {
  if (line.includes("\t")) return line.split("\t").map((part) => part.trim());
  if (line.includes("|")) return line.split("|").map((part) => part.trim()).filter(Boolean);
  if (/\s{2,}/.test(line)) return line.split(/\s{2,}/).map((part) => part.trim()).filter(Boolean);
  return [line.trim()];
}

function extractPastedText(text: string, source: string, sourceTitle: string | undefined, maxRows: number): {
  title: string;
  description: string;
  pageType: string;
  sheets: Sheet[];
} {
  const trimmed = text.trim();
  const looksLikeHtml = /<\/?(?:table|article|main|section|h[1-6]|ul|ol|p)\b/i.test(trimmed);
  if (looksLikeHtml) {
    const title = sourceTitle || pageTitle(trimmed, new URL("https://pasted-content.local"));
    const description = pageDescription(trimmed);
    const sheets = [
      ...extractTables(trimmed, source, maxRows),
      ...extractStructuredData(trimmed, source, maxRows),
      ...extractArticles(trimmed, source, maxRows),
    ].sort((a, b) => b.confidence - a.confidence);
    return {
      title,
      description,
      pageType: inferPageType(trimmed, sheets),
      sheets: sheets.length ? sheets.slice(0, 5) : [fallbackSheet(trimmed, source, title, description)],
    };
  }

  const lines = trimmed
    .split(/\r?\n/)
    .map((line) => line.replace(/^\s*(?:[-*•]\s+|\d+[.)]\s+)/, "").trim())
    .filter(Boolean);
  const title = sourceTitle || (lines[0]?.length <= 180 ? lines[0] : "Pasted webpage");
  const description = lines.slice(sourceTitle ? 0 : 1, sourceTitle ? 2 : 3).join(" ").slice(0, 280);
  const splitRows = lines.map(splitPastedLine);
  const maxWidth = Math.max(...splitRows.map((row) => row.length), 0);
  const delimited = maxWidth >= 2 && splitRows.filter((row) => row.length >= 2).length >= 2;
  if (delimited) {
    const headers = uniqueHeaders(splitRows[0], maxWidth);
    const rows = splitRows.slice(1, maxRows).map((row) =>
      Array.from({ length: maxWidth }, (_, index) => parseCellValue(row[index] ?? "")),
    );
    return {
      title,
      description,
      pageType: "Pasted table",
      sheets: [{
        id: "pasted-table",
        name: "Pasted table",
        description: "Columns detected from tabs, pipes, or spaced text.",
        confidence: 0.9,
        headers,
        rows,
        source,
      }],
    };
  }

  const keyValueRows = lines
    .map((line) => line.match(/^([^:]{2,80}):\s*(.+)$/))
    .filter((match): match is RegExpMatchArray => Boolean(match));
  if (keyValueRows.length >= 3) {
    return {
      title,
      description,
      pageType: "Pasted details",
      sheets: [{
        id: "pasted-details",
        name: "Page details",
        description: "Labeled facts detected in the pasted page text.",
        confidence: 0.84,
        headers: ["Field", "Value"],
        rows: keyValueRows.slice(0, maxRows).map((match) => [match[1].trim(), parseCellValue(match[2])]),
        source,
      }],
    };
  }

  const contentLines = lines.slice(sourceTitle ? 0 : 1, maxRows);
  return {
    title,
    description,
    pageType: "Pasted webpage text",
    sheets: [{
      id: "pasted-content",
      name: "Page content",
      description: "The pasted page text, separated into reviewable rows.",
      confidence: 0.72,
      headers: ["Content"],
      rows: (contentLines.length ? contentLines : [trimmed]).map((line) => [line]),
      source,
    }],
  };
}

function fallbackSheet(html: string, source: string, title: string, description: string): Sheet {
  const headings = html
    .match(/<h[1-3]\b[\s\S]*?<\/h[1-3]>/gi)
    ?.map(cleanText)
    .filter(Boolean)
    .slice(0, 30) ?? [];
  const rows = (headings.length ? headings : [title]).map((heading) => [
    heading,
    source,
  ]);
  return {
    id: "page-overview",
    name: "Page overview",
    description: "Key page metadata and headings, ready to refine.",
    confidence: 0.58,
    headers: ["Heading", "Source"],
    rows,
    source,
  };
}

function inferPageType(html: string, sheets: Sheet[]): string {
  const types = jsonLdBlocks(html).flatMap((value) => {
    const found: string[] = [];
    const visit = (item: unknown) => {
      if (Array.isArray(item)) return item.forEach(visit);
      if (item && typeof item === "object") {
        const type = (item as Record<string, unknown>)["@type"];
        if (typeof type === "string") found.push(type);
        visit((item as Record<string, unknown>)["@graph"]);
      }
    };
    visit(value);
    return found;
  });
  if (types.length) return types[0];
  if (sheets.some((sheet) => sheet.id.startsWith("table"))) return "Table page";
  if (sheets.some((sheet) => sheet.id === "articles")) return "Article listing";
  return "Web page";
}

router.post("/extract", async (req, res) => {
  const parsed = ExtractWebpageBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "Add a URL or paste at least 20 characters of page text." });
    return;
  }
  try {
    const { url: requestedUrl, text: pastedText, sourceTitle, maxRows = 100 } = parsed.data;
    if (!requestedUrl && !pastedText?.trim()) {
      res.status(400).json({ error: "Add a URL or paste at least 20 characters of page text." });
      return;
    }
    if (pastedText?.trim()) {
      const source = "pasted-content://local";
      const pasted = extractPastedText(pastedText, source, sourceTitle, maxRows);
      res.json({
        sourceUrl: source,
        title: pasted.title,
        description: pasted.description,
        fetchedAt: new Date().toISOString(),
        pageType: pasted.pageType,
        sheets: pasted.sheets,
      });
      return;
    }
    if (!requestedUrl) {
      res.status(400).json({ error: "Add a URL or paste page text." });
      return;
    }
    const { url, html } = await fetchPublicHtml(requestedUrl);
    const source = url.toString();
    const title = pageTitle(html, url);
    const description = pageDescription(html);
    const tableSheets = extractTables(html, source, maxRows);
    const structuredSheets = extractStructuredData(html, source, maxRows);
    const articleSheets = extractArticles(html, source, maxRows);
    const sheets = [...tableSheets, ...structuredSheets, ...articleSheets];
    const selected = sheets.length
      ? sheets.sort((a, b) => b.confidence - a.confidence).slice(0, 5)
      : [fallbackSheet(html, source, title, description)];
    res.json({
      sourceUrl: source,
      title,
      description,
      fetchedAt: new Date().toISOString(),
      pageType: inferPageType(html, selected),
      sheets: selected,
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "We couldn't read that page.";
    req.log.warn({ err: error, url: req.body?.url }, "webpage extraction failed");
    res.status(502).json({ error: message });
  }
});

export default router;