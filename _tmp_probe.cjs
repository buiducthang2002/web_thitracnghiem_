// Thử các loại file "giả docx" để biết mammoth báo lỗi gì với từng loại.
const mammoth = require("./node_modules/mammoth/mammoth.browser.js");
const JSZip = require("./node_modules/jszip");

const toAb = buf => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

async function probe(name, buf) {
  try {
    const r = await mammoth.extractRawText({ arrayBuffer: toAb(buf) });
    console.log(`[${name}] OK len=${r.value.length}`);
  } catch (e) {
    console.log(`[${name}] FAIL: ${e.message}`);
  }
}

async function main() {
  const htmlDoc = Buffer.from('﻿<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8"></head><body><p>Câu 1: abc?</p></body></html>', "utf8");
  await probe("word-html-.doc", htmlDoc);

  const rtf = Buffer.from("{\\rtf1\\ansi Câu 1: abc?\\par }", "utf8");
  await probe("rtf", rtf);

  const ole = Buffer.concat([Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]), Buffer.alloc(2000)]);
  await probe("ole-.doc", ole);

  await probe("plain-text", Buffer.from("Câu 1: abc?\nA. x\nB. y\nĐáp án: A", "utf8"));

  const emptyZip = new JSZip();
  emptyZip.file("hello.txt", "hi");
  await probe("zip-khong-phai-docx", await emptyZip.generateAsync({ type: "nodebuffer" }));

  // zip đúng cấu trúc docx nhưng document.xml không có w:body
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.folder("_rels").file(".rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.folder("word").file("document.xml", '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:notbody><w:p><w:r><w:t>Câu 1: abc?</w:t></w:r></w:p></w:notbody></w:document>');
  await probe("docx-thieu-w:body", await z.generateAsync({ type: "nodebuffer" }));

  // document.xml không khai báo namespace w (file do công cụ khác sinh ra)
  const z2 = new JSZip();
  z2.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z2.folder("_rels").file(".rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z2.folder("word").file("document.xml", '<?xml version="1.0"?><document><body><p><r><t>Câu 1: abc?</t></r></p></body></document>');
  await probe("docx-khong-namespace", await z2.generateAsync({ type: "nodebuffer" }));
}

main();
