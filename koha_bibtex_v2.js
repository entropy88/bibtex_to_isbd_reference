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
    const parts = name.split(",").map((p) => p.trim());
    if (parts.length === 2) {
      return `${parts[1]} ${parts[0]}`;
    }
  }
  return name.trim();
}

function formatResponsibility(rawResp) {
  if (!rawResp) return "";
  const authors = rawResp
    .split(/\s*(?:and|;)\s*/i)
    .map((a) => normalizeAuthorName(a.trim()));
  return authors.join(", ");
}

function getSortWord(entry) {
  const base = getFirst(entry, "sort_word") || "";
  const rawResp = getFirst(entry, "responsibility") || "";
  const itemTypes = getProps(entry, "item_type").map((v) => v.toUpperCase());

  if (!rawResp) return base;

  const parts = rawResp.split(/\s*(?:and|;)\s*/i).filter(Boolean);

  // ✅ Append "и др." only if there are multiple authors AND no GOI item type
  if (parts.length > 1 && !itemTypes.includes("GOI")) {
    return `${base} и др.`;
  }

  return base;
}

// ============================
// NEW Helper Function: robust pairing
// ============================
function getAlsoPairsSingleLine(entry) {
  const sources = getProps(entry, "also_source").map((s) => s.trim());
  let descriptions = getProps(entry, "also_description")
    .flatMap((d) => d.split(";"))
    .map((d) => d.trim())
    .filter(Boolean);

  const pairs = [];
  let descIndex = 0;

  for (let i = 0; i < sources.length; i++) {
    const src = sources[i];
    let desc = "";

    if (descIndex < descriptions.length) {
      if (i === sources.length - 1) {
        desc = descriptions.slice(descIndex).join(";");
        descIndex = descriptions.length;
      } else {
        desc = descriptions[descIndex];
        descIndex++;
      }
    }

    const sep = desc.startsWith(",") || desc.startsWith("(") || desc === "" ? "" : ", ";
    pairs.push((src + sep + desc).trim());
  }

  if (pairs.length === 0) return "";

  return pairs.join(";") + "; ";
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
  const itemTypes = getProps(entry, "item_type").map((v) => v.toUpperCase());
  const sortWord = getSortWord(entry);
  const rawResp = getFirst(entry, "responsibility") || "";
  const responsibility = formatResponsibility(rawResp);

  const main_sig = getFirst(entry, "main_sig") || "";
  const dep_sig = getFirst(entry, "dep_sig") || "";

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

  // build line2 as styled chunks
  const line2Parts = [];

  line2Parts.push({ text: `  ${title}` });
  if (column) line2Parts.push({ text: `. (${column})` });
  if (responsibility && !itemTypes.includes("GOI")) {
    line2Parts.push({ text: ` / ${responsibility}` });

  }

  if (source) {
    line2Parts.push({ text: `. – В: ` });
    line2Parts.push({ text: source, italics: true }); // ✅ italicized
  }

  if (journalCity) line2Parts.push({ text: ` (${journalCity})` });
  if (issue) line2Parts.push({ text: ` , бр. ${issue}` });
  if (year) line2Parts.push({ text: ` , (${year})` });
  if (pages) line2Parts.push({ text: ` , ${pages}` });

  const itemTypeLine = itemTypes.length > 0 ? `Item types: ${itemTypes.join(", ")}` : "";

  // ✅ Only show signatures when there are exactly two item types and one of them is GOI
  let sigLine = "";
  if (itemTypes.length === 2 && itemTypes.includes("GOI")) {
    sigLine = `${main_sig}       ${dep_sig}`;
  }

  return {
    mainLine: `${sigLine ? sigLine + "\n" : ""}${line1}`, // only non-styled part
    styledLine2: line2Parts, // send styled chunks separately
    notes,
    itemTypeLine,
    otherSources,
  };
}
// ============================
// NEW: Yearbook Formatter; handles multiple cases because structure SUCKS MASSIVE ASS
// ============================
function formatYearbook(entry) {
  // Extract relevant fields from the BibTeX entry
  const mainSig = getFirst(entry, "main_sig") || "";
  const title = getFirst(entry, "title") || "";
  const subtitle = getFirst(entry, "subtitle") || getFirst(entry, "substitle") || "";
  const responsibility = getFirst(entry, "responsibility") || "";
  const editionRaw = getFirst(entry, "edition") || ""; // can be publication info OR extent
  const publisher = getFirst(entry, "publisher") || "";
  const address = getFirst(entry, "address") || getFirst(entry, "place") || "";
  const year = getFirst(entry, "year") || "";
  let extent = getFirst(entry, "page_count") || getFirst(entry, "extent") || "";
  const aboutPersons = [...new Set(getProps(entry, "about_person"))]; // deduplicate

  const lines = [];

  // 1. Add library signature (main_sig) on its own line
  if (mainSig) lines.push(mainSig);

  // 2. Decide how to interpret "edition"
  let editionPub = "";    // for publication info
  let editionExtent = ""; // for extent info (pages, illustrations, etc.)

  if (editionRaw) {
    if (/с\./i.test(editionRaw)) {
      // If the string contains "с." → it's describing pages or illustrations
      editionExtent = editionRaw;
    } else {
      // Otherwise → assume it's publication statement (place, publisher, year)
      editionPub = editionRaw;
    }
  }

  // 3. Build the ISBD description line
  let isbdLine = "";

  // Title and subtitle
  if (title) isbdLine += title;
  if (subtitle) isbdLine += " : " + subtitle;

  // Statement of responsibility
  if (responsibility) isbdLine += " / " + responsibility;

  // Publication information
  if (editionPub) {
    // Directly use edition if it looks like publication info
    isbdLine += ". – " + editionPub;
  } else if (address || publisher || year) {
    // Or construct publication info from address, publisher, year
    let pubBlock = "";
    if (address) pubBlock += address;
    if (publisher) pubBlock += (pubBlock ? " : " : "") + publisher;
    if (year) pubBlock += (pubBlock ? ", " : "") + year;
    if (pubBlock) isbdLine += ". – " + pubBlock;
  }

  // Extent (priority: editionExtent > extent field)
  const finalExtent = editionExtent || extent;
  if (finalExtent) {
    isbdLine += ". – " + finalExtent;
  }

  // Add the ISBD record line
  if (isbdLine) lines.push(isbdLine);

  // 4. Add "about persons" (if any) as final line
  if (aboutPersons.length > 0) {
    lines.push("Имена на лица, за които става дума: " + aboutPersons.join(", "));
  }

  // 5. Return in the same structure used by other formatters
  return {
    mainLine: lines.join("\n"),
    notes: [],
    itemTypeLine: "Item types: GOI",
    otherSources: "",
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

const yearbooks = parsed.filter((p) => p.itemTypes.length === 1 && p.itemTypes[0] === "GOI");
const books = parsed.filter(
  (p) => bookTypes.includes((p.itemType || "").toUpperCase()) || yearbooks.includes(p)
);
const articles = parsed.filter(
  (p) =>
    (p.itemTypes.includes("GOI") && p.itemTypes.length > 1) ||
    ["JOU", "KRA", "ARTICLE", "DRU", "NSP"].includes((p.itemType || "").toUpperCase())
);
const others = parsed.filter(
  (p) => !books.includes(p) && !articles.includes(p)
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

function pushEntry(formatFn, rawEntries) {
  rawEntries.forEach((entry) => {
    const { mainLine, styledLine2, notes, itemTypeLine, otherSources } = formatFn(entry);

    // handle multi-line mainLine (plain text)
    if (mainLine) {
      mainLine.split("\n").forEach((line) => {
        if (line.trim()) {
          children.push(
            new Paragraph({ children: [new TextRun({ text: line, size: 24 })] })
          );
        }
      });
    }

    // handle styled line (e.g., italic source for articles)
    if (styledLine2 && styledLine2.length > 0) {
      children.push(
        new Paragraph({
          children: styledLine2.map(
            (part) =>
              new TextRun({
                text: part.text,
                italics: part.italics || false,
                bold: part.bold || false,
                size: 24,
              })
          ),
        })
      );
    }

    // item types
    if (itemTypeLine) {
      children.push(
        new Paragraph({ children: [new TextRun({ text: itemTypeLine, size: 24 })] })
      );
    }

    // notes
    if (notes && notes.length > 0) {
      notes.forEach((n) =>
        children.push(
          new Paragraph({
            children: [new TextRun({ text: "\t" + n, italics: true, size: 24 })],
          })
        )
      );
    }

    // other sources
    if (otherSources && otherSources.trim()) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: "\tВж. и: " + otherSources, size: 24 })],
        })
      );
    }

    // spacer line
    children.push(new Paragraph({ text: "" }));
  });
}
if (books.length > 0) {
  children.push(
    new Paragraph({ children: [new TextRun({ text: "КНИГИ", bold: true, size: 24 })] })
  );
  books.forEach((b) => {
    if (yearbooks.includes(b)) {
      pushEntry(formatYearbook, [b.entry]); // raw entry
    } else {
      pushEntry(formatBook, [b.entry]); // raw entry
    }
  });
}

if (articles.length > 0) {
  children.push(
    new Paragraph({ children: [new TextRun({ text: "СТАТИИ", bold: true, size: 24 })] })
  );
  pushEntry(formatArticle, articles.map((a) => a.entry)); // unwrap entries
}

if (others.length > 0) {
  children.push(
    new Paragraph({ children: [new TextRun({ text: "ДРУГИ", bold: true, size: 24 })] })
  );
  pushEntry(formatOther, others.map((o) => o.entry)); // unwrap entries
}

const doc = new Document({ sections: [{ children }] });

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Sorted DOCX saved to ${outputPath}`);
});
