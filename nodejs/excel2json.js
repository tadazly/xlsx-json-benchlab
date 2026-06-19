import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { StringDecoder } from "node:string_decoder";
import { performance } from "node:perf_hooks";
import unzipper from "unzipper";
import { SaxesParser } from "saxes";

const options = parseArgs(process.argv.slice(2));

if (!options.input) {
  console.error("missing input file: use --input data.xlsx [--output data.json] [--sheet Sheet1]");
  process.exit(1);
}

if (path.extname(options.input).toLowerCase() !== ".xlsx") {
  console.error(`input must be an .xlsx file: ${options.input}`);
  process.exit(1);
}

if (!options.output) {
  options.output = defaultOutputPath(options.input);
}

const startedAt = performance.now();

try {
  const rowCount = await convert(options);
  const elapsedMs = performance.now() - startedAt;
  console.log(`Converted ${rowCount} data rows to ${options.output} in ${formatDuration(elapsedMs)}`);
} catch (error) {
  console.error(error instanceof Error ? error.stack || error.message : error);
  process.exit(1);
}

async function convert({ input, output, sheet }) {
  fs.mkdirSync(path.dirname(output), { recursive: true });

  const zip = await unzipper.Open.file(input);
  const entries = new Map(zip.files.map((entry) => [entry.path.replaceAll("\\", "/"), entry]));
  const workbook = await readWorkbook(entries);
  const sheetInfo = selectSheet(workbook.sheets, sheet);
  const relTarget = workbook.relationships.get(sheetInfo.relationshipId);

  if (!relTarget) {
    throw new Error(`worksheet relationship not found: ${sheetInfo.relationshipId}`);
  }

  const sheetPath = normalizeZipPath("xl", relTarget);
  const sheetEntry = entries.get(sheetPath);
  if (!sheetEntry) {
    throw new Error(`worksheet entry not found: ${sheetPath}`);
  }

  const sharedStringsEntry = entries.get("xl/sharedStrings.xml");
  const sharedStrings = sharedStringsEntry ? await readSharedStrings(sharedStringsEntry) : [];

  return streamWorksheetToJSON(sheetEntry, sharedStrings, output);
}

async function readWorkbook(entries) {
  const workbookEntry = entries.get("xl/workbook.xml");
  const relsEntry = entries.get("xl/_rels/workbook.xml.rels");
  if (!workbookEntry) {
    throw new Error("workbook entry not found: xl/workbook.xml");
  }
  if (!relsEntry) {
    throw new Error("workbook relationships entry not found: xl/_rels/workbook.xml.rels");
  }

  const sheets = [];
  await parseXMLStream(workbookEntry.stream(), {
    opentag(tag) {
      if (tag.name !== "sheet") {
        return;
      }
      sheets.push({
        name: String(tag.attributes.name ?? ""),
        relationshipId: String(tag.attributes["r:id"] ?? "")
      });
    }
  });

  const relationships = new Map();
  await parseXMLStream(relsEntry.stream(), {
    opentag(tag) {
      if (tag.name !== "Relationship") {
        return;
      }
      relationships.set(String(tag.attributes.Id ?? ""), String(tag.attributes.Target ?? ""));
    }
  });

  return { sheets, relationships };
}

function selectSheet(sheets, requestedName) {
  if (sheets.length === 0) {
    throw new Error("workbook has no sheets");
  }

  if (!requestedName) {
    return sheets[0];
  }

  const found = sheets.find((item) => item.name === requestedName);
  if (!found) {
    throw new Error(`sheet not found: ${requestedName}`);
  }
  return found;
}

async function readSharedStrings(entry) {
  const sharedStrings = [];
  let currentText = "";
  let inStringItem = false;
  let inText = false;

  await parseXMLStream(entry.stream(), {
    opentag(tag) {
      if (tag.name === "si") {
        inStringItem = true;
        currentText = "";
      } else if (inStringItem && tag.name === "t") {
        inText = true;
      }
    },
    text(text) {
      if (inText) {
        currentText += text;
      }
    },
    closetag(tagName) {
      if (tagName === "t") {
        inText = false;
      } else if (tagName === "si") {
        sharedStrings.push(currentText);
        currentText = "";
        inStringItem = false;
      }
    }
  });

  return sharedStrings;
}

async function streamWorksheetToJSON(entry, sharedStrings, output) {
  const writer = fs.createWriteStream(output, { encoding: "utf8" });
  let outputChain = Promise.resolve();
  const queueWrite = (chunk) => {
    outputChain = outputChain.then(() => writeChunk(writer, chunk));
    return outputChain;
  };

  let columns = null;
  let rowCount = 0;
  let firstObject = true;
  let currentRow = null;
  let currentCell = null;
  let inValue = false;
  let inInlineText = false;

  await queueWrite("[\n");

  try {
    await parseXMLStream(entry.stream(), {
      opentag(tag) {
        if (tag.name === "row") {
          currentRow = [];
        } else if (tag.name === "c") {
          currentCell = {
            index: cellRefToColumnIndex(String(tag.attributes.r ?? "")),
            type: String(tag.attributes.t ?? ""),
            value: "",
            inlineText: ""
          };
        } else if (currentCell && tag.name === "v") {
          inValue = true;
        } else if (currentCell && tag.name === "t") {
          inInlineText = true;
        }
      },
      text(text) {
        if (!currentCell) {
          return;
        }
        if (inValue) {
          currentCell.value += text;
        } else if (inInlineText) {
          currentCell.inlineText += text;
        }
      },
      async closetag(tagName) {
        if (tagName === "v") {
          inValue = false;
        } else if (tagName === "t") {
          inInlineText = false;
        } else if (tagName === "c" && currentCell && currentRow) {
          if (currentCell.index >= 0) {
            currentRow[currentCell.index] = resolveCellValue(currentCell, sharedStrings);
          }
          currentCell = null;
        } else if (tagName === "row" && currentRow) {
          const rowValues = currentRow;
          currentRow = null;

          if (columns === null) {
            columns = selectHeaderColumns(rowValues);
            if (columns.length === 0) {
              throw new Error("header row has no non-empty columns");
            }
          } else if (!isEmptyRow(rowValues, columns)) {
            const prefix = firstObject ? "" : ",\n";
            firstObject = false;
            await queueWrite(prefix + buildJSONObject(columns, rowValues));
            rowCount += 1;
          }
        }
      }
    });

    if (columns === null) {
      throw new Error("sheet is empty");
    }

    await queueWrite("\n]\n");
    await outputChain;
    await closeWriter(writer);
    return rowCount;
  } catch (error) {
    writer.destroy();
    try {
      fs.rmSync(output, { force: true });
    } catch {
      // Best effort cleanup; preserve the original conversion error.
    }
    throw error;
  }
}

function resolveCellValue(cell, sharedStrings) {
  if (cell.type === "s") {
    const index = Number.parseInt(cell.value, 10);
    return Number.isInteger(index) ? String(sharedStrings[index] ?? "") : "";
  }
  if (cell.type === "inlineStr") {
    return cell.inlineText;
  }
  return cell.value;
}

function selectHeaderColumns(row) {
  const seen = new Map();
  const columns = [];

  row.forEach((value, index) => {
    const trimmed = String(value ?? "").trim();
    if (!trimmed) {
      return;
    }

    const header = uniqueHeader(trimmed, seen);
    columns.push({
      index,
      header,
      quotedHeader: JSON.stringify(header)
    });
  });

  return columns;
}

function uniqueHeader(header, seen) {
  const count = (seen.get(header) ?? 0) + 1;
  seen.set(header, count);
  return count === 1 ? header : `${header}_${count}`;
}

function isEmptyRow(row, columns) {
  return columns.every((column) => String(row[column.index] ?? "").trim() === "");
}

function buildJSONObject(columns, row) {
  let json = "{";
  for (let index = 0; index < columns.length; index += 1) {
    const column = columns[index];
    if (index > 0) {
      json += ",";
    }
    json += column.quotedHeader;
    json += ":";
    json += JSON.stringify(row[column.index] ?? "");
  }
  json += "}";
  return json;
}

async function parseXMLStream(stream, handlers) {
  const parser = new SaxesParser({ xmlns: false });
  const decoder = new StringDecoder("utf8");
  const pending = [];

  parser.on("opentag", (tag) => {
    if (handlers.opentag) {
      pending.push(Promise.resolve(handlers.opentag(tag)));
    }
  });
  parser.on("text", (text) => {
    if (handlers.text) {
      pending.push(Promise.resolve(handlers.text(text)));
    }
  });
  parser.on("closetag", (tag) => {
    if (handlers.closetag) {
      pending.push(Promise.resolve(handlers.closetag(typeof tag === "string" ? tag : tag.name)));
    }
  });

  for await (const chunk of stream) {
    parser.write(decoder.write(chunk));
    await drainPending(pending);
  }
  const tail = decoder.end();
  if (tail) {
    parser.write(tail);
  }
  parser.close();
  await drainPending(pending);
}

async function drainPending(pending) {
  while (pending.length > 0) {
    await pending.shift();
  }
}

function cellRefToColumnIndex(cellRef) {
  let index = 0;
  let seenLetter = false;

  for (const char of cellRef.toUpperCase()) {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) {
      break;
    }
    seenLetter = true;
    index = index * 26 + (code - 64);
  }

  return seenLetter ? index - 1 : -1;
}

function normalizeZipPath(baseDir, target) {
  if (target.startsWith("/")) {
    return target.slice(1).replaceAll("\\", "/");
  }

  const parts = `${baseDir}/${target}`.replaceAll("\\", "/").split("/");
  const normalized = [];
  for (const part of parts) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      normalized.pop();
    } else {
      normalized.push(part);
    }
  }
  return normalized.join("/");
}

function closeWriter(writer) {
  return new Promise((resolve, reject) => {
    writer.once("finish", resolve);
    writer.once("error", reject);
    writer.end();
  });
}

function writeChunk(writer, chunk) {
  return new Promise((resolve, reject) => {
    if (writer.write(chunk)) {
      resolve();
      return;
    }

    const onDrain = () => {
      writer.off("error", onError);
      resolve();
    };
    const onError = (error) => {
      writer.off("drain", onDrain);
      reject(error);
    };

    writer.once("drain", onDrain);
    writer.once("error", onError);
  });
}

function parseArgs(args) {
  const parsed = {};

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    const next = args[index + 1];

    if (arg === "--input" || arg === "-i") {
      parsed.input = next;
      index += 1;
    } else if (arg === "--output" || arg === "-o") {
      parsed.output = next;
      index += 1;
    } else if (arg === "--sheet" || arg === "-s") {
      parsed.sheet = next;
      index += 1;
    } else if (!parsed.input) {
      parsed.input = arg;
    } else if (!parsed.output) {
      parsed.output = arg;
    }
  }

  return parsed;
}

function defaultOutputPath(input) {
  return input.slice(0, -path.extname(input).length) + ".json";
}

function formatDuration(ms) {
  if (ms >= 1000) {
    return `${(ms / 1000).toFixed(2)}s`;
  }
  return `${Math.round(ms)}ms`;
}
