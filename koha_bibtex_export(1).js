const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const inputPath = path.join(__dirname, "shelf_old.bibtex");
const outputPath = path.join(__dirname, "sorted_shelf.docx");

const rawData = fs.readFileSync(inputPath, "utf8");

// ============================
// Helper Functions
// ============================

const entries = rawData
  .split(/(?=^@)/m)
  .map((e) => e.trim())
  .filter(Boolean);

function getProps(entry, propName) {
  const regex = new RegExp(`${propName}\\s*=\\s*\\{([^}]*)\\}`, "gi");
  let match;
  const values = [];
  while ((match = regex.exec(entry)) !== null) {
    values.push(match[1].trim());
  }
  return values;
}

function getFirst(entry, propName) {
  const vals = getProps(entry, propName);
  return vals.length > 0 ? vals[0] : "";
}

function getJoined(entry, propName, sep = "; ") {
  const vals = getProps(entry, propName);
  const unique = [...new Set(vals)];
  return unique.join(sep);
}

function cleanNotes(entry) {
  const notesRaw = getProps(entry, "abstract");
  return notesRaw.map((n) => n.replace(/^Съдържа и:\s*/i, "").trim());
}

function normalizeAuthorName(name) {
  if (name.includes(",")) {
    const parts = name.split(",").map(p => p.trim());
    if (parts.length === 2) {
      return `${parts[1]} ${parts[0]}`;
    }
  }
  return name.trim();
}

function formatResponsibility(rawResp) {
  if (!rawResp) return "";
  const authors = rawResp.split(/\s*(?:and|;)\s*/i).map(a => normalizeAuthorName(a.trim()));
  return authors.join(", ");
}

function getSortWord(entry) {
  const base = getFirst(entry, "sort_word") || "";
  const rawResp = getFirst(entry, "responsibility") || "";
  if (!rawResp) return base;
  const parts = rawResp.split(/\s*(?:and|;)\s*/i).filter(Boolean);
  return parts.length > 1 ? `${base} и др.` : base;
}

// ============================
// NEW Helper Function: robust pairing
// ============================
function getAlsoPairsSingleLine(entry) {
  const sources = getProps(entry, "also_source").map(s => s.trim());
  let descriptions = getProps(entry, "also_description")
    .flatMap(d => d.split(";")) // split multiple entries in one field
    .map(d => d.trim())
    .filter(Boolean);

  const pairs = [];
  let descIndex = 0;

  for (let i = 0; i < sources.length; i++) {
    const src = sources[i];
    let desc = "";

    if (descIndex < descriptions.length) {
      // For last source, attach all remaining descriptions
      if (i === sources.length - 1) {
        desc = descriptions.slice(descIndex).join(";"); // <- no space after semicolon
        descIndex = descriptions.length;
      } else {
        desc = descriptions[descIndex];
        descIndex++;
      }
    }

    // only add comma if desc does NOT start with ',' or '('
    const sep = desc.startsWith(",") || desc.startsWith("(") || desc === "" ? "" : ", ";
    pairs.push((src + sep + desc).trim());
  }

  if (pairs.length === 0) return "";

  return pairs.join(";") + ";"; // <- join without extra space
}

// ============================
// Formatting Functions
// ============================

function formatBook(entry) {
  const main_sig = getFirst(entry, "main_sig") || "";
  const dep_sig = getFirst(entry, "dep_sig") || "";
  const sortWord = getSortWord(entry);
  const rawResp = getFirst(entry, "responsibility") || "";
  const responsibility = formatResponsibility(rawResp);
  let title = getFirst(entry, "title") || "";
  const subtitle = getFirst(entry, "subtitle") || getFirst(entry, "substitle") || "";
  const edition = getJoined(entry, "edition");
  const place = getFirst(entry, "address") || getFirst(entry, "place") || "";
  const publisher = getFirst(entry, "publisher") || "";
  const year = getFirst(entry, "year") || "";
  const extent = getFirst(entry, "page_count") || getFirst(entry, "extent") || "";
  const dimensions = getFirst(entry, "illustrations") || getFirst(entry, "dimensions") || "";
  const series = getFirst(entry, "series") || "";
  const isbn = getFirst(entry, "isbn") || "";
  const book_info = getFirst(entry, "book_info") || "";

  const otherSources = getAlsoPairsSingleLine(entry);

  const itemTypes = [...new Set(getProps(entry, "item_type").map((v) => v.toUpperCase()))];
  const notes = cleanNotes(entry);

  const line1 = itemTypes.includes("GOI") ? "" : sortWord;
  const line01 = `${main_sig}       ${dep_sig}`;

  if (itemTypes.includes("CDD")) title += " [CD-ROM]";

  let line2 = `  ${title}`;
  if (subtitle) line2 += ` : ${subtitle}`;
  if (responsibility && !itemTypes.includes("GOI")) line2 += ` / ${responsibility}`;
  if (edition) line2 += `. – ${edition}`;
  if (book_info) line2 += `. – ${book_info}`;

  if (place || publisher || year) {
    let pubBlock = "";
    if (place) pubBlock += place;
    if (publisher) pubBlock += (pubBlock ? " : " : "") + publisher;
    if (year) pubBlock += (pubBlock ? ", " : "") + year;
    line2 += `. – ${pubBlock}`;
  }

  if (extent || dimensions) {
    let physBlock = "";
    if (extent) physBlock += extent;
    if (dimensions) physBlock += (physBlock ? " ; " : "") + dimensions;
    line2 += `. – ${physBlock}`;
  }

  if (series) line2 += `. – (${series})`;
  if (isbn) line2 += `. – ISBN ${isbn}`;

  const itemTypeLine = itemTypes.length > 0 ? `Item types: ${itemTypes.join(", ")}` : "";

  return {
    mainLine: `${line01}\n${line1}\n${line2}`,
    notes,
    itemTypeLine,
    otherSources,
  };
}

function formatArticle(entry) {
  const itemTypes = getProps(entry, "item_type").map(v => v.toUpperCase());
  const sortWord = getSortWord(entry);
  const rawResp = getFirst(entry, "responsibility") || "";
  const responsibility = formatResponsibility(rawResp);

  const line1 = itemTypes.includes("GOI") ? "" : sortWord;

  let title = getFirst(entry, "title") || "";
  const source = getFirst(entry, "source") || "";
  const issue = getFirst(entry, "issue") || "";
  const year = getFirst(entry, "year") || "";
  const pages = getFirst(entry, "art_pages") || "";
  const column = getFirst(entry, "column") || "";
  const journalCity = getFirst(entry, "journal_city") || "";

  const otherSources = getAlsoPairsSingleLine(entry);
  const notes = cleanNotes(entry);

  let line2 = `  ${title}`;
  if (responsibility && !itemTypes.includes("GOI")) {
    line2 += ` / ${responsibility}`;
    if (column) line2 += `.(${column})`;
  }

  if (source) line2 += `. - В: ${source}`;
  if (journalCity) line2 += ` (${journalCity})`;
  if (issue) line2 += ` , бр. ${issue}`;
  if (year) line2 += ` , (${year})`;
  if (pages) line2 += ` , ${pages}`;

  const itemTypeLine = itemTypes.length > 0 ? `Item types: ${itemTypes.join(", ")}` : "";

  return {
    mainLine: `${line1}\n${line2}`,
    notes,
    itemTypeLine,
    otherSources
  };
}

function formatOther(entry) {
  const rawResp = getFirst(entry, "responsibility") || "";
  const responsibility = formatResponsibility(rawResp);
  let title = getFirst(entry, "title") || "";
  const year = getFirst(entry, "year") || "";
  const itemTypes = [...new Set(getProps(entry, "item_type").map((v) => v.toUpperCase()))];
  const notes = cleanNotes(entry);

  let line1 = itemTypes.includes("GOI") ? "" : responsibility;
  let line2 = `  ${title}`;
  if (year) line2 += ` (${year})`;

  const itemTypeLine = itemTypes.length > 0 ? `Item types: ${itemTypes.join(", ")}` : "";

  const otherSources = getAlsoPairsSingleLine(entry);

  return { mainLine: `${line1}\n${line2}`, notes, itemTypeLine, otherSources };
}

// ============================
// Parsing and Sorting
// ============================

const parsed = entries.map((entry) => {
  const itemTypes = getProps(entry, "item_type").map((v) => v.toUpperCase());
  const primaryType = itemTypes.length > 0 ? itemTypes[0] : "";
  const year = parseInt(getFirst(entry, "year"), 10) || 0;
  return { entry, itemTypes, itemType: primaryType, year };
});

const bookTypes = ["KNG", "CDD"];
parsed.sort((a, b) => {
  const aIsBook = bookTypes.includes((a.itemType || "").toUpperCase());
  const bIsBook = bookTypes.includes((b.itemType || "").toUpperCase());
  if (aIsBook && !bIsBook) return -1;
  if (!aIsBook && bIsBook) return 1;
  return a.year - b.year;
});

const articleTypes = ["JOU", "KRA", "ARTICLE", "DRU", "NSP", "GOI"];
const books = parsed.filter((p) => bookTypes.includes((p.itemType || "").toUpperCase()));
const articles = parsed.filter((p) => articleTypes.includes((p.itemType || "").toUpperCase()));
const others = parsed.filter(
  (p) =>
    !bookTypes.includes((p.itemType || "").toUpperCase()) &&
    !articleTypes.includes((p.itemType || "").toUpperCase())
);

// ============================
// Build DOCX
// ============================

const total = parsed.length;
const children = [];

children.push(
  new Paragraph({
    children: [
      new TextRun({
        text: `Общо записи: ${total} (Книги: ${books.length}, Статии: ${articles.length}, Други: ${others.length})`,
        bold: true,
        size: 24,
      }),
    ],
  }),
  new Paragraph({ text: "" })
);

function pushEntry(formatFn, entries) {
  entries.forEach((e) => {
    const { mainLine, notes, itemTypeLine, otherSources } = formatFn(e.entry);

    mainLine.split("\n").forEach((line) => {
      children.push(new Paragraph({ children: [new TextRun({ text: line, size: 24 })] }));
    });

    if (itemTypeLine)
      children.push(new Paragraph({ children: [new TextRun({ text: itemTypeLine, size: 24 })] }));

    if (notes)
      notes.forEach((n) =>
        children.push(
          new Paragraph({ children: [new TextRun({ text: "\t" + n, italics: true, size: 24 })] })
        )
      );

    // Only add Вж. и: if there is content
    if (otherSources && otherSources.trim()) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: "\tВж. и: " + otherSources, size: 24 })],
        })
      );
    }

    children.push(new Paragraph({ text: "" }));
  });
}

if (books.length > 0) {
  children.push(new Paragraph({ children: [new TextRun({ text: "КНИГИ", bold: true, size: 24 })] }));
  pushEntry(formatBook, books);
}

if (articles.length > 0) {
  children.push(new Paragraph({ children: [new TextRun({ text: "СТАТИИ", bold: true, size: 24 })] }));
  pushEntry(formatArticle, articles);
}

if (others.length > 0) {
  children.push(new Paragraph({ children: [new TextRun({ text: "ДРУГИ", bold: true, size: 24 })] }));
  pushEntry(formatOther, others);
}

const doc = new Document({ sections: [{ children }] });

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Sorted DOCX saved to ${outputPath}`);
});
