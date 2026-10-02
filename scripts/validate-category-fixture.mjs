import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import XLSX from "xlsx";
import PizZip from "pizzip";

const root = path.resolve("tests/fixtures");
const fields = ["mis_padres", "mis_padrinos", "abuelos_paternos", "el_sr_y_la_sra", "yo"];
const excel = XLSX.read(await fs.readFile(path.join(root, "datos_sinteticos_familia.xlsx")), { type: "buffer" });
const sheet = excel.Sheets.Datos_sinteticos;
const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
assert.equal(rows.length, 50, "La fixture debe conservar 50 registros sintéticos");
assert.deepEqual(Object.keys(rows[0]), fields, "Los encabezados Excel cambiaron o están en otro orden");

const templateBuffer = await fs.readFile(path.join(root, "plantilla_categorias_familia.docx"));
const baseZip = new PizZip(templateBuffer);
const baseXml = baseZip.file("word/document.xml").asText();
const plain = (value) => value.replace(/<[^>]+>/g, "");
const tables = Array.from(baseXml.matchAll(/<w:tbl\b[^>]*>[\s\S]*?<\/w:tbl>/g)).map((match) => match[0]);
assert.ok(fields.every((field) => baseXml.includes(field)), "La plantilla debe contener las cinco categorías visibles");

const escapeXml = (value) => String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");

function render(row) {
  const zip = new PizZip(templateBuffer);
  let xml = zip.file("word/document.xml").asText();
  for (const field of fields) {
    const table = tables.find((candidate) => plain(candidate).includes(field));
    const rowsXml = Array.from(table.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)).map((match) => match[0]);
    const headerIndex = rowsXml.findIndex((candidate) => plain(candidate).includes(field));
    const headerCells = Array.from(rowsXml[headerIndex].matchAll(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g)).map((match) => match[0]);
    const columnIndex = headerCells.findIndex((candidate) => plain(candidate).includes(field));
    const targetCells = Array.from(rowsXml[headerIndex + 1].matchAll(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g)).map((match) => match[0]);
    const cell = targetCells[columnIndex];
    let replaced = false;
    const nextCell = cell.replace(/(<w:t\b[^>]*>)([^<]*_{3,}[^<]*)(<\/w:t>)/g, (_match, start, _placeholder, end) => {
      if (replaced) return `${start}${end}`;
      replaced = true;
      return `${start}${escapeXml(row[field])}${end}`;
    });
    assert.ok(replaced, `No se encontró línea de captura para ${field}`);
    xml = xml.replace(cell, nextCell);
  }
  zip.file("word/document.xml", xml);
  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

const rendered = rows.map(render);
const outputZip = new PizZip(rendered[0]);
const outputXml = outputZip.file("word/document.xml").asText();
const bodies = rendered.map((buffer) => new PizZip(buffer).file("word/document.xml").asText().match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/)[1].replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/, ""));
const sectPr = outputXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)[0];
outputZip.file("word/document.xml", outputXml.replace(/<w:body[^>]*>[\s\S]*?<\/w:body>/, `<w:body>${bodies.join('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')}${sectPr}</w:body>`));

const temporaryOutput = path.join(await fs.mkdtemp(path.join(os.tmpdir(), "generador-documentos-")), "resultado.docx");
await fs.writeFile(temporaryOutput, outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" }));
const finalXml = outputZip.file("word/document.xml").asText();
for (const field of fields) {
  for (const row of rows) assert.ok(finalXml.includes(escapeXml(row[field])), `No se insertó el valor de ${field}`);
}
assert.equal((finalXml.match(/w:type="page"/g) || []).length, 49, "Deben existir 49 separadores para 50 registros");
console.log(`category-fixture-ok: ${rows.length} filas, ${fields.length} categorías, salida temporal ${temporaryOutput}`);
