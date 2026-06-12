const fs = require("fs");
const dir = "./_xtmp";
const ss = fs.readFileSync(dir + "/xl/sharedStrings.xml", "utf8");
const strings = [];
const re = /<si>([\s\S]*?)<\/si>/g;
let m;
while ((m = re.exec(ss))) {
  const tx = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join("");
  strings.push(tx);
}
console.log("num shared strings:", strings.length);
const sheet = fs.readFileSync(dir + "/xl/worksheets/sheet1.xml", "utf8");
const dim = sheet.match(/<dimension ref="([^"]+)"/);
console.log("dimension:", dim && dim[1]);
const rows = [...sheet.matchAll(/<row [^>]*?r="(\d+)"[^>]*?>([\s\S]*?)<\/row>/g)];
console.log("num row elements:", rows.length);
function decode(s) {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}
function cellVal(c) {
  const t = c.match(/ t="([^"]+)"/);
  const v = c.match(/<v>([\s\S]*?)<\/v>/);
  if (!v) return "";
  if (t && t[1] === "s") return decode(strings[+v[1]] || "");
  return v[1];
}
function printRow(r) {
  const rn = r[1];
  const cells = [...r[2].matchAll(/<c r="([A-Z]+)(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)];
  const out = cells.map((c) => c[1] + ":" + cellVal(c[0]));
  console.log("ROW", rn, "::", out.join(" | "));
}
const byNum = {};
for (const r of rows) byNum[+r[1]] = r;
console.log("--- first rows ---");
for (const r of rows.slice(0, 6)) printRow(r);
console.log("--- last rows ---");
for (const r of rows.slice(-6)) printRow(r);
console.log("--- around 520-535 ---");
for (let i = 515; i <= 535; i++) if (byNum[i]) printRow(byNum[i]);
const maxRow = Math.max(...rows.map((r) => +r[1]));
console.log("max row number:", maxRow);
