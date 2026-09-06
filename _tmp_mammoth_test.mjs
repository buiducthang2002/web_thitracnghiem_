import mammoth from "mammoth";
import fs from "fs";
const file = process.argv[2] || "1-cau-hoi-de-chan.docx";
const buf = fs.readFileSync(file);
const ab = { buffer: buf };

try {
  const h = await mammoth.convertToHtml(ab, { styleMap: ['u => u'] });
  console.log("HTML OK, len =", h.value.length);
  console.log("first 300:", h.value.slice(0, 300));
  console.log("messages:", h.messages.slice(0, 5).map(m => m.message));
} catch (e) { console.log("HTML FAIL:", e.message); }
try {
  const r = await mammoth.extractRawText(ab);
  console.log("RAW OK, len =", r.value.length);
  console.log(JSON.stringify(r.value.slice(0, 400)));
} catch (e) { console.log("RAW FAIL:", e.message); }
