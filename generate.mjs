import fs from "fs";
import path from "path";
import { parse } from "@babel/parser";
import _traverseMod from "@babel/traverse";
const traverse = _traverseMod.default || _traverseMod;

const ORIG = "/tmp/App.jsx.orig.bak";
const src = fs.readFileSync(ORIG, "utf8");
const lines = src.split(/\r\n|\n/);
const deps = JSON.parse(fs.readFileSync("/home/claude/deps.json", "utf8"));

const buckets = [
  ["docTypes", 1, 45, "src/data/docTypes.jsx", "module"],
  ["parsers", 46, 1441, "src/lib/parsers.js", "module"],
  ["helpers", 1442, 1610, "src/lib/helpers.js", "module"],
  ["templatesConst", 1611, 1645, "src/data/templates.js", "module"],
  ["docGenerators", 1646, 2609, "src/lib/docGenerators.js", "module"],
  ["App", 2610, 2973, "src/App.jsx", "app"],
  ["formPanels", 2974, 3411, "src/components/FormPanels.jsx", "module"],
  ["docForm", 3412, 5049, "src/components/DocForm.jsx", "module"],
  ["docxPreviews", 5050, 5244, "src/components/DocxPreviews.jsx", "module"],
  ["bast", 5245, 5372, "src/lib/bast.js", "module"],
  ["berkasPembayaran", 5373, 6963, "src/lib/berkasPembayaran.js", "module"],
  ["suratKepala", 6964, 7108, "src/lib/suratKepala.js", "module"],
  ["docPreview", 7109, 7140, "src/components/DocPreview.jsx", "module"],
];

const fileByBucket = Object.fromEntries(buckets.map((b) => [b[0], b[3]]));

function bucketFor(line) {
  for (const [name, s, e] of buckets) {
    if (line >= s && line <= e) return name;
  }
  return null;
}

// ---- Parse original source to find TRUE top-level statements & their lines ----
const ast = parse(src, { sourceType: "module", plugins: ["jsx"] });
let programPath;
traverse(ast, {
  Program(p) {
    programPath = p;
    p.stop();
  },
});

// exact top-level statement start lines (1-indexed), excluding ImportDeclaration
const topLevelStarts = new Set();
for (const stmt of programPath.node.body) {
  if (stmt.type === "ImportDeclaration") continue;
  topLevelStarts.add(stmt.loc.start.line);
}

function relImport(fromFile, toFile) {
  let rel = path.relative(path.dirname(fromFile), toFile);
  rel = rel.replace(/\\/g, "/");
  rel = rel.replace(/\.jsx?$/, "");
  if (!rel.startsWith(".")) rel = "./" + rel;
  return rel;
}

function buildImports(bucketName, filePath) {
  const needs = deps[bucketName] || [];
  const externalBySource = new Map();
  const localByOwner = new Map();

  for (const item of needs) {
    if (item.type === "import") {
      const info = item.info;
      if (!info) continue;
      if (!externalBySource.has(info.source)) {
        externalBySource.set(info.source, { default: null, namespace: null, named: new Set() });
      }
      const bucket = externalBySource.get(info.source);
      if (info.kind === "default") bucket.default = info.localName;
      else if (info.kind === "namespace") bucket.namespace = info.localName;
      else if (info.kind === "named") {
        bucket.named.add(info.imported === info.localName ? info.imported : `${info.imported} as ${info.localName}`);
      }
    } else if (item.type === "local") {
      if (!localByOwner.has(item.ownerBucket)) localByOwner.set(item.ownerBucket, new Set());
      localByOwner.get(item.ownerBucket).add(item.name);
    }
  }

  const out = [];
  const sourceOrder = ["react", "framer-motion", "lucide-react", "xlsx", "pizzip", "jszip", "docxtemplater", "docxtemplater-image-module-free", "file-saver", "docx-preview"];
  const sources = [...externalBySource.keys()].sort((a, b) => {
    const ia = sourceOrder.indexOf(a);
    const ib = sourceOrder.indexOf(b);
    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
  });
  for (const source of sources) {
    const b = externalBySource.get(source);
    const parts = [];
    if (b.default) parts.push(b.default);
    if (b.namespace) parts.push(`* as ${b.namespace}`);
    if (b.named.size) parts.push(`{ ${[...b.named].sort().join(", ")} }`);
    out.push(`import ${parts.join(", ")} from "${source}";`);
  }
  if (sources.length) out.push("");

  const owners = [...localByOwner.keys()].sort();
  for (const owner of owners) {
    const names = [...localByOwner.get(owner)].sort();
    const rel = relImport(filePath, fileByBucket[owner]);
    out.push(`import { ${names.join(", ")} } from "${rel}";`);
  }
  if (owners.length) out.push("");

  return out.join("\n");
}

// Precisely add "export " only at verified top-level statement start lines
function exportifyPrecise(bodyLines, bucketStartOffset) {
  return bodyLines
    .map((line, idx) => {
      const originalLineNo = bucketStartOffset + idx;
      if (!topLevelStarts.has(originalLineNo)) return line;
      if (/^export\s/.test(line)) return line; // already exported (App default)
      if (/^(function|async function|const|let|class)\s/.test(line)) {
        return "export " + line;
      }
      return line;
    })
    .join("\n");
}

const HEADER_IMPORT_END_LINE = 20;

for (const [name, startLine, endLine, filePath] of buckets) {
  const sliceStart = Math.max(startLine, HEADER_IMPORT_END_LINE + 1);
  const bodyLines = lines.slice(sliceStart - 1, endLine).map((l) => l.replace(/\r$/, ""));
  let body =
    name === "App"
      ? bodyLines.join("\n")
      : exportifyPrecise(bodyLines, sliceStart);
  body = body.replace(/\n+$/, "\n");

  const header =
    name === "App"
      ? `// ============================================================\n// Portal Administrasi SE2026 — BPS Kota Jakarta Timur\n// Dependencies: npm install xlsx docxtemplater pizzip file-saver docx-preview\n// ============================================================\n`
      : `// ============================================================\n// Portal Administrasi SE2026 — bagian: ${name}\n// ============================================================\n`;

  const importsBlock = buildImports(name, filePath);
  const fullContent = header + "\n" + (importsBlock ? importsBlock + "\n" : "") + body + "\n";

  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, fullContent, "utf8");
  console.log("wrote", filePath, body.split("\n").length, "lines");
}
