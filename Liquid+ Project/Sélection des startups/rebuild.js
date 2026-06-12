const fs = require("fs");
const P = "C:/Users/jeanp/AppData/Local/Temp/";
const j = (f) => JSON.parse(fs.readFileSync(P + f, "utf8"));
const an = j("rannees.json"), ce = j("rcentres.json");
const name = (arr) => Object.fromEntries(arr.map((t) => [t.id, t.name]));
const anN = name(an), ceN = name(ce);
let projs = [...j("rproj_1.json"), ...j("rproj_2.json")];
function strip(html) {
  return (html || "").replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim();
}
const rows = projs.map((p) => {
  const years = (p.annees || []).map((id) => anN[id]).filter(Boolean);
  const yearNum = Math.max(0, ...years.map((y) => parseInt(y)).filter((n) => !isNaN(n)));
  return {
    name: strip(p.title && p.title.rendered),
    year: yearNum,
    centre: (p.centres || []).map((id) => ceN[id]).filter(Boolean).join(", "),
    link: p.link,
  };
});
const since2024 = rows.filter((r) => r.year >= 2024).sort((a, b) => b.year - a.year || a.name.localeCompare(b.name));
fs.writeFileSync("c:/Users/jeanp/Documents/Liquid-/Liquid+ Project/Sélection des startups/iss_since2024.json", JSON.stringify(since2024, null, 1));
console.log("rebuilt:", since2024.length);
since2024.forEach((r, i) => console.log((530 + i) + "\t" + r.name + "\t" + r.year + "\t" + r.centre));
