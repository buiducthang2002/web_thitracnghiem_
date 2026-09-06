// Chạy đúng nhánh code của trình duyệt (JSZip + arrayBuffer) để xem lỗi có tái hiện không.
const fs = require("fs");
const mammoth = require("./node_modules/mammoth/mammoth.browser.js");

const file = process.argv[2] || "1-cau-hoi-de-chan.docx";
const buf = fs.readFileSync(file);
const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

mammoth.convertToHtml({ arrayBuffer: ab }, { styleMap: ["u => u"] })
  .then(r => console.log("browser convertToHtml OK, len =", r.value.length))
  .catch(e => console.log("browser convertToHtml FAIL:", e.message))
  .then(() => mammoth.extractRawText({ arrayBuffer: ab }))
  .then(r => console.log("browser extractRawText OK, len =", r.value.length))
  .catch(e => console.log("browser extractRawText FAIL:", e.message));
