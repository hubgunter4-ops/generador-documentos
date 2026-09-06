import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

const out = path.resolve("sample-templates");
const xml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

async function createDoc(filename, title, lines) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
  const paragraphs = [title, "", ...lines].map((line, index) => index === 0 ? `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${xml(line)}</w:t></w:r></w:p>` : `<w:p><w:r><w:t xml:space="preserve">${xml(line)}</w:t></w:r></w:p>`).join("");
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`);
  await fs.writeFile(path.join(out, filename), await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

await fs.mkdir(out, { recursive: true });
await createDoc("plantilla-prueba-1.docx", "DOCUMENTO 1 · DATOS FAMILIARES", [
  "Yo: {persona_nombre}",
  "Papá: {padre_nombre}",
  "Mamá: {madre_nombre}",
  "Abuelos maternos:",
  "{abuelos_maternos}",
  "Abuelos paternos:",
  "{abuelos_paternos}",
  "Padrino: {padrino_nombre}",
  "Madrina: {madrina_nombre}",
]);
await createDoc("plantilla-prueba-2.docx", "DOCUMENTO 2 · GRUPOS COMPUESTOS", [
  "Yo: {persona_nombre}",
  "Mis padres:",
  "{mis_padres}",
  "Mis padrinos:",
  "{mis_padrinos}",
]);
await createDoc("plantilla-prueba-3.docx", "DOCUMENTO 3 · CONSTANCIA", [
  "Se hace constar que {persona_nombre}.",
  "Padre: {padre_nombre} · Madre: {madre_nombre}",
  "Padrino: {padrino_nombre} · Madrina: {madrina_nombre}",
  "Abuelos maternos: {abuelos_maternos}",
  "Abuelos paternos: {abuelos_paternos}",
]);
console.log(`Created sample templates in ${out}`);
