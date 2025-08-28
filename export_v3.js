const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const inputPath = path.join(__dirname, "test_bobi.bibtex");
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

function formatAuthors(rawAuthors) {
  if (!rawAuthors) return "";
  return rawAuthors
    .split(" and ")
    .map(a => normalizeAuthorName(a))
    .join(", ");
}

// ============================
// Formatting Functions
// ============================

function formatBook(entry) {
  const main_sig = getFirst(entry, "main_sig") || "";
  const dep_sig = getFirst(entry, "dep_sig") || "";
  const sortWord = getFirst(entry, "sort_word") || getFirst(entry, "author") || "";
  const rawAuthor = getFirst(entry, "author") || "";
  const author = formatAuthors(rawAuthor);
  let title = getFirst(entry, "title") || "";
  const subtitle = getFirst(entry, "subtitle") || getFirst(entry, "substitle") || "";
  const responsibility = getFirst(entry, "responsibility") || author;
  const edition = getJoined(entry, "edition");
  const place = getFirst(entry, "address") || getFirst(entry, "place") || "";
  const publisher = getFirst(entry, "publisher") || "";
  const year = getFirst(entry, "year") || "";
  const extent = getFirst(entry, "page_count") || getFirst(entry, "extent") || "";
  const dimensions = getFirst(entry, "illustrations") || getFirst(entry, "dimensions") || "";
  const series = getFirst(entry, "series") || "";
  const isbn = getFirst(entry, "isbn") || "";
  const book_info = getFirst(entry, "book_info") || "";

  const otherSources = getProps(entry, "other_sources"); // keep as array
  const itemTypes = [...new Set(getProps(entry, "item_type").map((v) => v.toUpperCase()))];
  const notes = cleanNotes(entry);

  const line1 = itemTypes.includes("GOI") ? "" : sortWord;
  const line01 = `${main_sig}       ${dep_sig}`;

  if (itemTypes.includes("CDD")) title += " [CD-ROM]";

  let line2 = `  ${title}`;
  if (subtitle) line2 += ` : ${subtitle}`;
  if (responsibility) line2 += ` / ${responsibility}`;
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
  const sortWord = getFirst(entry, "sort_word") || "";
  const line1 = itemTypes.includes("GOI") ? "" : sortWord;

  const rawAuthor = getFirst(entry, "author") || "";
  const authorStr = formatAuthors(rawAuthor);

  let title = getFirst(entry, "title") || "";
  const source = getFirst(entry, "source") || "";
  const issue = getFirst(entry, "issue") || "";
  const year = getFirst(entry, "year") || "";
  const pages = getFirst(entry, "art_pages") || "";
  const column = getFirst(entry, "column") || "";
  const journalCity = getFirst(entry, "journal_city") || "";

  const otherSources = getProps(entry, "other_sources"); // array 
  const notes = cleanNotes(entry);

  let line2 = `  ${title}`;
  if (authorStr) {
    line2 += ` / ${authorStr}`;
    if (column) line2 += `. — (${column})`;
  }

  // Include source as placeholder for italics
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
  const rawAuthor = getFirst(entry, "author") || "";
  const author = formatAuthors(rawAuthor);
  let title = getFirst(entry, "title") || "";
  const year = getFirst(entry, "year") || "";
  const itemTypes = [...new Set(getProps(entry, "item_type").map((v) => v.toUpperCase()))];
  const notes = cleanNotes(entry);
  
  let line1 = `${author}`;
  let line2 = `  ${title}`;
  if (year) line2 += ` (${year})`;

  const itemTypeLine = itemTypes.length > 0 ? `Item types: ${itemTypes.join(", ")}` : "";

  return { mainLine: `${line1}\n${line2}`, notes, itemTypeLine,  otherSources: [] };
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

const bookTypes = ["KNG", "GOI", "CDD"];
parsed.sort((a, b) => {
  const aIsBook = bookTypes.includes((a.itemType || "").toUpperCase());
  const bIsBook = bookTypes.includes((b.itemType || "").toUpperCase());
  if (aIsBook && !bIsBook) return -1;
  if (!aIsBook && bIsBook) return 1;
  return a.year - b.year;
});

const articleTypes = ["JOU", "KRA", "ARTICLE", "DRU", "NSP"];
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
      const sourceMatch = line.match(/В: ([^,]+)/); // match only source
      if (sourceMatch) {
        const beforeSource = line.slice(0, sourceMatch.index + 3);
        const sourceText = sourceMatch[1];
        const afterSource = line.slice(sourceMatch.index + 3 + sourceText.length);
        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: beforeSource, size: 24 }),
              new TextRun({ text: sourceText, italics: true, size: 24 }),
              new TextRun({ text: afterSource, size: 24 }),
            ],
          })
        );
      } else {
        children.push(new Paragraph({ children: [new TextRun({ text: line, size: 24 })] }));
      }
    });

    if (itemTypeLine)
      children.push(new Paragraph({ children: [new TextRun({ text: itemTypeLine, size: 24 })] }));

    

    // Notes indented one tab
    if (notes)
      notes.forEach((n) =>
        children.push(
          new Paragraph({ children: [new TextRun({ text: "\t" + n, italics: true, size: 24 })] })
        )
      );

    // Other sources indented one tab
    if (otherSources && otherSources.length > 0) {
      otherSources.forEach((src, idx) => {
        const prefix = idx === 0 ? "\tВж. и: " : "\t";
        children.push(
          new Paragraph({
            children: [new TextRun({ text: `${prefix}${src}`, size: 24 })],
          })
        );
      });
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
