import fs from "node:fs/promises";
import path from "node:path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const root = path.resolve("sample-templates");
const data = { persona_nombre: "Juan Pérez López", padre_nombre: "Carlos Pérez", madre_nombre: "Ana López", abuelo_paterno: "José Pérez", abuela_paterna: "Rosa Hernández", abuelo_materno: "Roberto López", abuela_materna: "Elena Martínez", abuelos_maternos: "Roberto López\nElena Martínez", abuelos_paternos: "José Pérez\nRosa Hernández", padrino_nombre: "Miguel Torres", madrina_nombre: "Laura Gómez", el_sr: "Miguel Torres", la_sra: "Laura Gómez", mis_padres: "Carlos Pérez\nAna López", mis_padrinos: "Miguel Torres\nLaura Gómez" };
for (const filename of ["plantilla-prueba-1.docx", "plantilla-prueba-2.docx", "plantilla-prueba-3.docx"]) {
  const source = await fs.readFile(path.join(root, filename));
  const engine = new Docxtemplater(new PizZip(source), { paragraphLoop: true, linebreaks: true, nullGetter: () => "" });
  engine.render(data);
  const rendered = engine.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
  if (!rendered || rendered.length < 500) throw new Error(`DOCX inválido: ${filename}`);
  console.log(`${filename}: OK (${rendered.length} bytes)`);
}
