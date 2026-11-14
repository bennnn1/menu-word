/* menu-word.js
   Reads the locked Word menu template (menu.docx) and turns it into menu data.
   Any page can use it:
     <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
     <script src="menu-word.js"></script>
     const menu = await MenuFromWord.load("menu.docx");
     menu.csv          -> same columns as the old menu.csv
     menu.community    -> community name typed in the Word file
     menu.week1Starts  -> "YYYY-MM-DD" (the Sunday Week 1 began), or "" if not filled in
     menu.cycleWeeks   -> number of weeks in the Word file
*/
(function () {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const KEYS = ["item", "appetizer", "entree", "sides", "dessert"];

  function textOf(p) {
    let out = "";
    const walk = n => {
      for (const c of n.childNodes) {
        if (c.namespaceURI === W) {
          if (c.localName === "t") out += c.textContent;
          else if (c.localName === "tab") out += " ";
          else if (c.localName === "br" || c.localName === "cr") out += "\n";
          else if (c.localName === "delText" || c.localName === "instrText") continue;
          else walk(c);
        } else walk(c);
      }
    };
    walk(p);
    return out;
  }

  function linesOf(sdt) {
    const pr = sdt.getElementsByTagNameNS(W, "sdtPr")[0];
    if (pr && pr.getElementsByTagNameNS(W, "showingPlcHdr").length) return [];
    const content = sdt.getElementsByTagNameNS(W, "sdtContent")[0];
    if (!content) return [];
    return [...content.getElementsByTagNameNS(W, "p")]
      .flatMap(p => textOf(p).split("\n"))
      .map(s => s.replace(/\s+/g, " ").trim())
      .filter(Boolean);
  }

  function toSunday(text) {
    const s = (text || "").trim();
    if (!s) return "";
    let d = null, m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) d = new Date(+m[1], m[2] - 1, +m[3]);
    else if ((m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/))) d = new Date(m[3].length === 2 ? 2000 + +m[3] : +m[3], m[1] - 1, +m[2]);
    else { const t = Date.parse(s); if (!isNaN(t)) d = new Date(t); }
    if (!d || isNaN(d)) return "";
    d.setDate(d.getDate() - d.getDay());   // back up to that week's Sunday
    const pad = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  const q = v => /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;

  async function load(url) {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error("Menu file not found");
    const zip = await JSZip.loadAsync(await res.arrayBuffer());
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");

    const data = {}; let community = "", start = "", weeks = 0;
    for (const sdt of doc.getElementsByTagNameNS(W, "sdt")) {
      const tagEl = sdt.getElementsByTagNameNS(W, "tag")[0];
      const tag = tagEl && tagEl.getAttributeNS(W, "val");
      if (!tag) continue;
      const lines = linesOf(sdt);
      if (tag === "Community") { community = lines.join(" "); continue; }
      if (tag === "Week1Starts") { start = toSunday(lines.join(" ")); continue; }
      const [wk, day, meal, key] = tag.split("|");
      const week = parseInt(wk.slice(1), 10);
      if (!week || !day || !meal || !KEYS.includes(key)) continue;
      weeks = Math.max(weeks, week);
      if (!lines.length) continue;
      ((data[week] ??= {})[day] ??= {})[meal] ??= {};
      data[week][day][meal][key] = lines;
    }

    const rows = [["Week", "Day", "Meal", "Item", "Appetizer", "Entree", "Sides", "Dessert", "HolidayGif"]];
    for (const week of Object.keys(data).sort((a, b) => a - b))
      for (const day of Object.keys(data[week]))
        for (const meal of Object.keys(data[week][day])) {
          const m = data[week][day][meal];
          rows.push([week, day, meal, ...KEYS.map(k => (m[k] || []).join("\n")), ""]);
        }
    const csv = rows.map(r => r.map(v => q(String(v))).join(",")).join("\n");
    return { csv, community, week1Starts: start, cycleWeeks: weeks };
  }

  window.MenuFromWord = { load };
})();
