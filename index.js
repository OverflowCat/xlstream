import * as zip from "https://deno.land/x/zipjs/index.js";
import * as DATA from "./data.js";

const { configure, TextReader, ZipReader, ZipWriter } = zip;

configure({
  useWebWorkers: true,
});

export async function getZipFileBlob() {
  console.log(0);
  const zipFileStream = new TransformStream();
  const transformStream = new TransformStream();
  const readable = transformStream.readable;
  const writable = transformStream.writable;
  const zipWriter = new ZipWriter(zipFileStream.writable);
  console.log(1);

  new Promise(async (resolve) => {
    await zipWriter.add("docProps/app.xml", new TextReader(DATA.DOCPROPS_APP));
    await zipWriter.add(
      "docProps/core.xml",
      new TextReader(DATA.DOCPROPS_CORE)
    );
    await zipWriter.add(
      "[Content_Types].xml",
      new TextReader(DATA.CONTENT_TYPES_XML)
    );
    await zipWriter.add("_rels/.rels", new TextReader(DATA.RELS_DOT_RELS));
    await zipWriter.add("xl/workbook.xml", new TextReader(DATA.WORKBOOK_XML));
    await zipWriter.add(
      "xl/_rels/workbook.xml.rels",
      new TextReader(DATA.WORKBOOK_XML_RELS)
    );
    await zipWriter.add("xl/styles.xml", new TextReader(DATA.STYLES_XML));
    await zipWriter.add(
      "xl/sharedStrings.xml",
      new TextReader(DATA.SHARED_STRINGS_XML)
    );
    await zipWriter.add("xl/theme/theme1.xml", new TextReader(DATA.THEME_XML));
    await zipWriter.add(
      "xl/worksheets/sheet2.xml",
      new TextReader(DATA.SHEET2_XML)
    );
    await zipWriter.add(
      "xl/worksheets/sheet3.xml",
      new TextReader(DATA.SHEET3_XML)
    );
    writeDataToStream(writable);
    await zipWriter.add("xl/worksheets/sheet1.xml", readable);
    // await zipWriter.add("hello.txt", new TextReader("Hello world!"));
    await zipWriter.close();
  });

  console.log(2);
  return zipFileStream.readable;
}

/**
 * Writes data to the specified stream.
 *
 * @param {WritableStream} stream - The stream to write data to.
 * @return {Promise<void>}
 */
async function writeDataToStream(stream) {
  const enc = new TextEncoder(); // always utf-8
  const writer = stream.getWriter();
  const lines = 114514;
  await writer.write(
    enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">
  <dimension ref="A1:C${lines}" />
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="C2" sqref="C2" /></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.25" />
  <cols><col min="3" max="3" width="18.72" style="1" /></cols>
  <sheetData>`)
  );
  const progress = document.querySelector("#progress");
  const rowContent = document.querySelector("#row");
  for (let i = 1; i <= lines; i++) {
    const data = genRandomData(i);
    await writer.write(enc.encode(data));
    progress.textContent = `已写入 ${i} / ${lines} 行…`;
    rowContent.textContent = data;
    console.info("Writing" + i);
  }
  await writer.write(enc.encode(DATA.TEMPLATE_END));
  await writer.close();
  progress.textContent = `已写入 ${lines} / ${lines} 行。下载完成！`;
}

const pick = (arr) => arr[Math.floor(Math.random() * arr.length)];

/**
 * Generate random data.
 *
 * @param {number} i - The parameter i.
 * @return {string} The generated random data.
 */
function genRandomData(i) {
  i = i.toString();
  // create a long string by repeating i
  return `<row r="${i}" spans="1:3" ht="14" x14ac:dyDescent="0.25">
<c r="A${i}"><v>${i}</v></c>
<c r="B${i}"><v>uuid ${crypto.randomUUID().repeat(100)}</v></c>
<c r="C${i}"><v>${pick(DATA.el1) + pick(DATA.el2) + pick(DATA.el3)}</v></c>
</row>`;
}
