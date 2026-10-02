import { useMemo, useState } from "react";
import { Download, FileSpreadsheet, FileText, Info, ShieldCheck, Upload } from "lucide-react";
import * as XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { toast } from "sonner";

const FIELD_LABELS: Record<string, string> = {
  mis_padres: "Mis padres",
  mis_padrinos: "Mis padrinos",
  abuelos_paternos: "Abuelos paternos",
  el_sr_y_la_sra: "El Sr. y la Sra.",
  yo: "Yo",
};
const FIELD_KEYS = Object.keys(FIELD_LABELS);
type ExcelRow = Record<string, string>;
type ExcelData = { rows: ExcelRow[]; columns: string[]; ignoredColumns: string[] };

function normalizeField(value: string) {
  const field = value.trim().replace(/^\{+|\}+$/g, "").toLowerCase();
  return field === "persona_nombre" ? "yo" : field;
}

function cellText(value: unknown) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toLocaleDateString("es-MX");
  return String(value).trim();
}

async function readExcel(file: File): Promise<ExcelData> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("El archivo Excel no contiene hojas.");
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
  if (!rawRows.length) throw new Error("El archivo Excel no contiene filas de datos.");

  const originalColumns = Object.keys(rawRows[0]);
  const normalizedColumns = originalColumns.map(normalizeField);
  const duplicates = normalizedColumns.filter((field, index) => normalizedColumns.indexOf(field) !== index);
  if (duplicates.length) throw new Error(`Hay columnas repetidas en Excel: ${Array.from(new Set(duplicates)).join(", ")}. Cada campo debe tener una sola columna.`);

  const recognized = normalizedColumns.filter(field => FIELD_KEYS.includes(field));
  if (!recognized.length) throw new Error(`No se encontró ninguna columna válida. Usa: ${FIELD_KEYS.join(", ")}.`);
  const ignoredColumns = normalizedColumns.filter(field => !FIELD_KEYS.includes(field));
  const rows = rawRows.map((raw) => {
    const result: ExcelRow = {};
    originalColumns.forEach((column, index) => {
      const field = normalizedColumns[index];
      if (FIELD_KEYS.includes(field)) result[field] = cellText(raw[column]);
    });
    return result;
  });
  return { rows, columns: recognized, ignoredColumns };
}

function extractMarkers(buffer: ArrayBuffer) {
  const zip = new PizZip(buffer.slice(0));
  const xml = Object.keys(zip.files)
    .filter((name) => name.endsWith(".xml") && (name.includes("document") || name.includes("header") || name.includes("footer")))
    .map((name) => zip.file(name)?.asText() ?? "")
    .join(" ");
  const explicit = Array.from(new Set(Array.from(xml.matchAll(/\{([^{}]+)\}/g)).map((match) => normalizeField(match[1])))).filter(Boolean);
  if (explicit.length) return explicit;
  return FIELD_KEYS.filter((field) => xml.includes(field));
}

function hasExplicitMarkers(buffer: ArrayBuffer) {
  const zip = new PizZip(buffer.slice(0));
  return Object.keys(zip.files)
    .filter((name) => name.endsWith(".xml") && (name.includes("document") || name.includes("header") || name.includes("footer")))
    .some((name) => /\{([^{}]+)\}/.test(zip.file(name)?.asText() ?? ""));
}

function escapeXml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function injectCategoryValues(zip: PizZip, row: ExcelRow) {
  const file = zip.file("word/document.xml");
  const xml = file?.asText();
  if (!xml) throw new Error("La plantilla Word no contiene word/document.xml.");
  let updated = xml;
  const tables = Array.from(xml.matchAll(/<w:tbl\b[^>]*>[\s\S]*?<\/w:tbl>/g)).map((match) => match[0]);
  const plain = (value: string) => value.replace(/<[^>]+>/g, "");
  for (const field of FIELD_KEYS) {
    const table = tables.find((candidate) => plain(candidate).includes(field));
    if (!table) continue;
    const rows = Array.from(table.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)).map((match) => match[0]);
    const headerIndex = rows.findIndex((candidate) => plain(candidate).includes(field));
    const headerCells = headerIndex >= 0 ? Array.from(rows[headerIndex].matchAll(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g)).map((match) => match[0]) : [];
    const columnIndex = headerCells.findIndex((candidate) => plain(candidate).includes(field));
    const targetRow = rows[headerIndex + 1];
    const targetCells = targetRow ? Array.from(targetRow.matchAll(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g)).map((match) => match[0]) : [];
    const cell = columnIndex >= 0 ? targetCells[columnIndex] : undefined;
    if (!cell) continue;
    let replaced = false;
    const nextCell = cell.replace(/(<w:t\b[^>]*>)([^<]*_{3,}[^<]*)(<\/w:t>)/g, (_match, start, _placeholder, end) => {
      if (replaced) return `${start}${end}`;
      replaced = true;
      return `${start}${escapeXml(row[field] ?? "")}${end}`;
    });
    updated = updated.replace(cell, nextCell);
  }
  zip.file("word/document.xml", updated);
  return zip;
}

function renderRow(template: ArrayBuffer, row: ExcelRow, hasExplicitMarkers: boolean) {
  if (!hasExplicitMarkers) return injectCategoryValues(new PizZip(template.slice(0)), row).generate({ type: "arraybuffer", compression: "DEFLATE" }) as ArrayBuffer;
  const engine = new Docxtemplater(new PizZip(template.slice(0)), { paragraphLoop: true, linebreaks: true, nullGetter: () => "" });
  const values = Object.fromEntries(FIELD_KEYS.map((key) => [key, row[key] ?? ""]));
  values.persona_nombre = row.yo ?? "";
  engine.render(values);
  return engine.getZip().generate({ type: "arraybuffer", compression: "DEFLATE" }) as ArrayBuffer;
}

function mergeDocxDocuments(documents: ArrayBuffer[]) {
  if (!documents.length) throw new Error("No hay documentos para combinar.");
  const first = new PizZip(documents[0].slice(0));
  const firstXml = first.file("word/document.xml")?.asText();
  if (!firstXml) throw new Error("La plantilla Word no contiene word/document.xml.");
  const bodies = documents.map((buffer) => {
    const xml = new PizZip(buffer.slice(0)).file("word/document.xml")?.asText();
    const match = xml?.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/);
    if (!match) throw new Error("No se pudo leer el cuerpo de la plantilla Word.");
    return match[1].replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/, "");
  });
  const sectPr = firstXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0] ?? "";
  const body = bodies.join('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  first.file("word/document.xml", firstXml.replace(/<w:body[^>]*>[\s\S]*?<\/w:body>/, `<w:body>${body}${sectPr}</w:body>`));
  return first.generate({ type: "blob", compression: "DEFLATE" }) as Blob;
}

function download(blob: Blob, name: string) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

export default function App() {
  const [template, setTemplate] = useState<ArrayBuffer | null>(null);
  const [templateName, setTemplateName] = useState("");
  const [markers, setMarkers] = useState<string[]>([]);
  const [explicitMarkers, setExplicitMarkers] = useState(false);
  const [excelName, setExcelName] = useState("");
  const [excel, setExcel] = useState<ExcelData | null>(null);
  const [outputName, setOutputName] = useState("documentos_generados.docx");
  const [busy, setBusy] = useState(false);

  const missingColumns = useMemo(() => markers.filter((marker) => FIELD_KEYS.includes(marker) && !excel?.columns.includes(marker)), [markers, excel]);
  const unknownMarkers = useMemo(() => markers.filter((marker) => !FIELD_KEYS.includes(marker)), [markers]);

  const chooseTemplate = async (file: File) => {
    try {
      if (!file.name.toLowerCase().endsWith(".docx")) throw new Error("Selecciona un archivo Word .docx.");
      const buffer = await file.arrayBuffer();
      setTemplate(buffer); setTemplateName(file.name); setMarkers(extractMarkers(buffer)); setExplicitMarkers(hasExplicitMarkers(buffer));
      toast.success("Plantilla Word cargada.");
    } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudo leer la plantilla."); }
  };

  const chooseExcel = async (file: File) => {
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error("Selecciona un archivo Excel .xlsx o .xls.");
      setExcel(await readExcel(file)); setExcelName(file.name);
      toast.success("Excel cargado y validado.");
    } catch (error) { setExcel(null); toast.error(error instanceof Error ? error.message : "No se pudo leer el Excel."); }
  };

  const generate = async () => {
    if (!template || !excel) return toast.error("Carga una plantilla Word y un archivo Excel.");
    if (!outputName.toLowerCase().endsWith(".docx")) return toast.error("El nombre de salida debe terminar en .docx.");
    setBusy(true);
    try {
      const documents = excel.rows.map((row) => renderRow(template, row, explicitMarkers));
      download(mergeDocxDocuments(documents), outputName);
      toast.success(`Se generó un Word único con ${documents.length} registro(s).`);
    } catch (error) { toast.error(error instanceof Error ? error.message : "No se pudo generar el documento."); }
    finally { setBusy(false); }
  };

  return <div className="min-h-screen bg-[#f4f6f4] text-slate-900">
    <header className="border-b bg-[#0f2f2b] text-white"><div className="container py-7"><div className="flex items-center gap-3"><div className="rounded-xl bg-[#d2f34c] p-2 text-[#0f2f2b]"><FileText size={22} /></div><span className="text-xs font-semibold uppercase tracking-[.24em] text-[#d2f34c]">Herramienta local</span></div><h1 className="mt-3 text-3xl font-semibold tracking-tight">Generador desde Excel</h1><p className="mt-1 text-sm text-emerald-100/75">Una plantilla Word de varias páginas · un registro por fila · un documento final</p></div></header>
    <main className="container space-y-6 py-8">
      <section className="grid gap-6 lg:grid-cols-[1.1fr_.9fr]">
        <Card className="border-0 shadow-sm"><CardHeader><CardTitle>1. Cargar archivos</CardTitle><CardDescription>Usa una sola plantilla .docx y un Excel con una fila por documento.</CardDescription></CardHeader><CardContent className="space-y-4">
          <label className="block cursor-pointer rounded-xl border border-dashed border-slate-300 bg-slate-50 p-4 hover:border-emerald-700 hover:bg-emerald-50"><input type="file" accept=".docx" className="sr-only" onChange={(event) => event.target.files?.[0] && chooseTemplate(event.target.files[0])} /><div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-slate-500"><FileText size={15} /> Plantilla Word</div><div className="mt-2 truncate text-sm font-medium">{templateName || "Seleccionar plantilla de varias páginas"}</div><div className="mt-2 flex items-center gap-1 text-xs text-emerald-800"><Upload size={14} /> Elegir archivo .docx</div></label>
          <label className="block cursor-pointer rounded-xl border border-dashed border-slate-300 bg-slate-50 p-4 hover:border-emerald-700 hover:bg-emerald-50"><input type="file" accept=".xlsx,.xls" className="sr-only" onChange={(event) => event.target.files?.[0] && chooseExcel(event.target.files[0])} /><div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-slate-500"><FileSpreadsheet size={15} /> Datos Excel</div><div className="mt-2 truncate text-sm font-medium">{excelName || "Seleccionar archivo Excel"}</div><div className="mt-2 flex items-center gap-1 text-xs text-emerald-800"><Upload size={14} /> Elegir .xlsx o .xls</div></label>
          <div><label className="text-sm font-medium" htmlFor="output">Nombre del documento final</label><input id="output" value={outputName} onChange={(event) => setOutputName(event.target.value)} className="mt-2 h-10 w-full rounded-md border bg-white px-3 text-sm" /></div>
          <Button className="w-full" onClick={generate} disabled={busy || !template || !excel}><Download size={16} /> {busy ? "Generando…" : "Generar un Word único"}</Button>
        </CardContent></Card>
        <Card className="border-0 bg-[#e8efe8] shadow-sm"><CardHeader><CardTitle>Campos aceptados</CardTitle><CardDescription>Los nombres de las columnas Excel deben coincidir con estos campos.</CardDescription></CardHeader><CardContent className="space-y-2">{FIELD_KEYS.map((field) => <div key={field} className="flex items-center justify-between rounded-lg bg-white/70 px-3 py-2 text-sm"><span>{FIELD_LABELS[field]}</span><code className="text-[11px] text-emerald-800">{`{${field}}`}</code></div>)}<div className="mt-4 flex gap-2 text-xs text-slate-600"><Info size={15} className="mt-0.5 shrink-0" /><span>Los campos compuestos vienen directamente de Excel. La aplicación no los vuelve a construir ni los mezcla con otros nombres.</span></div></CardContent></Card>
      </section>
      {(template || excel) && <section className="grid gap-6 lg:grid-cols-[1fr_1fr]"><Card className="border-0 shadow-sm"><CardHeader><div className="flex items-center justify-between"><div><CardTitle className="text-base">2. Validación de correspondencia</CardTitle><CardDescription>Revisión antes de generar.</CardDescription></div><Badge variant="secondary">{excel ? `${excel.rows.length} fila(s)` : "Sin Excel"}</Badge></div></CardHeader><CardContent className="space-y-3">{template && <div><b className="text-sm">Marcadores detectados en Word</b><div className="mt-2 flex flex-wrap gap-2">{markers.length ? markers.map((marker) => <code key={marker} className={`rounded px-2 py-1 text-xs ${excel?.columns.includes(marker) ? "bg-emerald-100 text-emerald-900" : "bg-amber-100 text-amber-900"}`}>{`{${marker}}`}</code>) : <span className="text-sm text-slate-500">No se detectaron marcadores.</span>}</div></div>}{missingColumns.length > 0 && <div className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-sm text-amber-900">Faltan en Excel: <b>{missingColumns.map((field) => `{${field}}`).join(", ")}</b>. Se dejarán vacíos.</div>}{unknownMarkers.length > 0 && <div className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-sm text-amber-900">Marcadores no definidos en el flujo: <b>{unknownMarkers.map((field) => `{${field}}`).join(", ")}</b>. Revisa el nombre para evitar confusiones.</div>}{excel?.ignoredColumns.length ? <div className="rounded-lg border border-slate-200 bg-slate-50 p-3 text-xs text-slate-600">Columnas no utilizadas: {excel.ignoredColumns.join(", ")}</div> : null}</CardContent></Card><Card className="border-0 shadow-sm"><CardHeader><CardTitle className="text-base">Muestra de Excel</CardTitle><CardDescription>Primera fila que se pasará a Word.</CardDescription></CardHeader><CardContent>{excel ? <div className="space-y-2">{FIELD_KEYS.map((field) => <div key={field} className="flex justify-between gap-4 border-b py-2 text-sm"><span className="text-slate-500">{field}</span><span className="max-w-[60%] whitespace-pre-line text-right font-medium">{excel.rows[0]?.[field] || "(vacío)"}</span></div>)}</div> : <span className="text-sm text-slate-500">Carga un Excel para ver sus datos.</span>}</CardContent></Card></section>}
      <section className="rounded-xl border border-emerald-200 bg-emerald-50/70 p-4 text-sm text-emerald-950"><div className="flex gap-3"><ShieldCheck className="mt-0.5 shrink-0" size={18} /><div><b>Privacidad:</b> el Excel y el DOCX se procesan dentro del navegador. No se envían a un servidor ni se guardan en el repositorio.</div></div></section>
      <Separator /><p className="text-xs text-slate-500">Cada fila de Excel produce una copia completa de la plantilla, incluyendo todas sus páginas. Las copias se unen con un salto de página en un único archivo Word.</p>
    </main>
  </div>;
}
