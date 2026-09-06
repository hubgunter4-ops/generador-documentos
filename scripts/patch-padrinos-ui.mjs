import fs from "node:fs/promises";
const file = "client/src/App.tsx";
let source = await fs.readFile(file, "utf8");
source = source.replace(/placeholder=\{\'\(yo\): Juan Pérez López\\n\(papá\): Carlos Pérez Hernández\\n\(mamá\): Ana López Martínez\\n\(abuelo paterno\): José Pérez\\n\(abuela paterna\): Rosa Hernández\\n\(padrino\): Miguel Torres\\n\(madrina\): Laura Gómez\'\}/, "placeholder={'(yo): Juan Pérez López\\n(papá): Carlos Pérez Hernández\\n(mamá): Ana López Martínez\\n(padrinos): Miguel Torres\\nLaura Gómez'}");
source = source.replace(/\[\"yo\", \"papá\", \"mamá\", \"abuelo paterno\", \"abuela paterna\", \"padrino\", \"madrina\"\]/, '["yo", "papá", "mamá", "padrinos"]');
source = source.replace(/<span>Etiquetas:<\/span>\{\[\"yo\", \"papá\", \"mamá\", \"padrinos\"\]\.map\(label => <code key=\{label\} className=\"rounded bg-slate-100 px-1\.5 py-0\.5\">\(\{label\}\)<\/code>\)\}<\/div>/, '<span>Etiquetas:</span>{["yo", "papá", "mamá", "padrinos"].map(label => <code key={label} className="rounded bg-slate-100 px-1.5 py-0.5">({label})</code>)}<span className="ml-1">Dos nombres, en líneas separadas o unidos por «E».</span></div>');
await fs.writeFile(file, source);
