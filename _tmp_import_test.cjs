// Kiểm thử các hàm đọc file trong App.jsx bằng cách trích thẳng đoạn source (không JSX).
const fs = require("fs");
const JSZip = require("./node_modules/jszip");
const mammoth = require("./node_modules/mammoth/mammoth.browser.js");

const src = fs.readFileSync("src/App.jsx", "utf8");
const start = src.indexOf("  const markStyledRuns");
const end = src.indexOf("  const wordImportError");
if (start < 0 || end < 0) throw new Error("không tìm thấy đoạn helper trong App.jsx");
const body = src.slice(start, end);

const BOLD = "\u0001";
const OPT_LETTERS = "ABCDEFGHIJ";
const MAX_OPTS = OPT_LETTERS.length;
const make = new Function("BOLD", "OPT_LETTERS", "MAX_OPTS", "JSZip", "TextDecoder",
  body + "\n return {decodeBytes, sniffFileKind, wordXmlToMarkedText, zipDocumentToMarkedText, rtfToText, htmlToMarkedText, parseWordText};");
const H = make(BOLD, OPT_LETTERS, MAX_OPTS, JSZip, TextDecoder);

const toAb = buf => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const show = (name, res) => console.log(`  ${name}: found=${res.found}` +
  (res.found ? ` | câu 1 = "${res.parsed[0].text.slice(0, 45)}" | opts=${res.parsed[0].opts.length} | đáp án=${OPT_LETTERS[res.parsed[0].ans]}` : ""));

const NOI_DUNG = `Câu 1: Chi bộ có những hình thức sinh hoạt nào?
A. Sinh hoạt lãnh đạo
B. Sinh hoạt chuyên đề
C. Cả A và B
D. Không có đáp án đúng
Đáp án: C
Chủ đề: Nội quy
Mức độ: Dễ

Câu 2: Đảng viên dự bị có thời gian bao lâu?
A. 6 tháng
B. 12 tháng
C. 18 tháng
D. 24 tháng
Đáp án: B`;

async function main() {
  // 1. File .docx thật của người dùng
  const real = fs.readFileSync("1-cau-hoi-de-chan.docx");
  console.log("[1] file .docx thật (đường mammoth):");
  const html = await mammoth.convertToHtml({ arrayBuffer: toAb(real) }, { styleMap: ["u => u"] });
  show("mammoth html", H.parseWordText(H.htmlToMarkedText(html.value)));
  console.log("  sniff =", H.sniffFileKind(toAb(real)));
  console.log("[2] file .docx thật (đường dự phòng bóc XML):");
  show("xml fallback", H.parseWordText(await H.zipDocumentToMarkedText(toAb(real))));

  // 3. docx lệch chuẩn: có document.xml nhưng không có w:body -> mammoth bó tay
  const mkZip = async (docXml) => {
    const z = new JSZip();
    z.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    z.folder("_rels").file(".rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    z.folder("word").file("document.xml", docXml);
    return toAb(await z.generateAsync({ type: "nodebuffer" }));
  };
  const paras = (lines, boldLine) => lines.map(l =>
    `<w:p><w:r>${l === boldLine ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${l.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</w:t></w:r></w:p>`).join("");
  const lines = NOI_DUNG.split("\n");

  const badBody = await mkZip(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:notbody>${paras(lines)}</w:notbody></w:document>`);
  console.log("[3] docx thiếu thẻ w:body (đúng lỗi đang gặp):");
  try { await mammoth.extractRawText({ arrayBuffer: badBody }); console.log("  mammoth: OK (?)"); }
  catch (e) { console.log("  mammoth FAIL:", e.message); }
  console.log("  sniff =", H.sniffFileKind(badBody));
  show("xml fallback", H.parseWordText(await H.zipDocumentToMarkedText(badBody)));

  // 4. docx không khai báo namespace w
  const noNs = await mkZip(`<?xml version="1.0"?><document><body>${paras(lines).replace(/w:/g, "")}</body></document>`);
  console.log("[4] docx không có namespace:");
  try { await mammoth.extractRawText({ arrayBuffer: noNs }); console.log("  mammoth: OK (?)"); }
  catch (e) { console.log("  mammoth FAIL:", e.message); }
  show("xml fallback", H.parseWordText(await H.zipDocumentToMarkedText(noNs)));

  // 5. đáp án đánh dấu bằng in đậm thay vì dòng "Đáp án:"
  const boldOnly = await mkZip(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:nobody>${paras(lines.filter(l => !/^Đáp án/.test(l)), "C. Cả A và B")}</w:nobody></w:document>`);
  console.log("[5] docx lệch chuẩn, đáp án đúng chỉ được bôi đậm:");
  show("xml fallback", H.parseWordText(await H.zipDocumentToMarkedText(boldOnly)));

  // 6. file HTML mang đuôi .doc/.docx (chính là file mẫu tải về)
  const esc = s => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const htmlDoc = Buffer.from("\ufeff<!DOCTYPE html><html xmlns:o=\"urn:schemas-microsoft-com:office:office\"><head><meta charset=\"utf-8\"></head><body>" +
    lines.map(l => `<p style="margin:0">${esc(l) || "&nbsp;"}</p>`).join("") + "</body></html>", "utf8");
  console.log("[6] file HTML đặt tên .docx:", "sniff =", H.sniffFileKind(toAb(htmlDoc)));
  show("html", H.parseWordText(H.htmlToMarkedText(H.decodeBytes(new Uint8Array(htmlDoc)))));

  // 7. RTF
  const rtf = Buffer.from("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times New Roman;}}\\pard " +
    lines.map(l => l.replace(/[^\x00-\x7f]/g, c => `\\u${c.charCodeAt(0)}?`)).join("\\par ") + "\\par }", "utf8");
  console.log("[7] file RTF đặt tên .doc:", "sniff =", H.sniffFileKind(toAb(rtf)));
  show("rtf", H.parseWordText(H.rtfToText(H.decodeBytes(new Uint8Array(rtf)))));

  // 8. file text thuần
  const txt = Buffer.from(NOI_DUNG, "utf8");
  console.log("[8] file text thuần:", "sniff =", H.sniffFileKind(toAb(txt)));
  show("text", H.parseWordText(H.decodeBytes(new Uint8Array(txt))));

  // 9. ODT
  const z = new JSZip();
  z.file("mimetype", "application/vnd.oasis.opendocument.text");
  z.file("content.xml", `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis" xmlns:text="urn:oasis:text"><office:body><office:text>${lines.map(l => `<text:p>${esc(l)}</text:p>`).join("")}</office:text></office:body></office:document-content>`);
  const odt = toAb(await z.generateAsync({ type: "nodebuffer" }));
  console.log("[9] file .odt (LibreOffice):", "sniff =", H.sniffFileKind(odt));
  show("xml fallback", H.parseWordText(await H.zipDocumentToMarkedText(odt)));
}

main().catch(e => { console.error("TEST LỖI:", e); process.exit(1); });
