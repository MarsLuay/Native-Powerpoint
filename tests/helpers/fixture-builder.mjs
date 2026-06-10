import { Buffer } from "node:buffer";
import { crc32 } from "pptx-svg";

const encoder = new TextEncoder();

const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";

const MAIN_CONTENT_TYPES = {
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  ppsx: "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
  potx: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
  pptm: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
  ppsm: "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
  potm: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
};

const MACRO_FORMATS = new Set(["pptm", "ppsm", "potm"]);
const FIXTURE_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR42mNk+M/wHwMDAwMjiAUAHgwCAWcOPK0AAAAASUVORK5CYII=",
  "base64",
);

function toBytes(value) {
  if (typeof value === "string") {
    return encoder.encode(value);
  }

  return value instanceof Uint8Array ? value : new Uint8Array(value);
}

function concat(parts) {
  return Buffer.concat(parts.map((part) => Buffer.from(part)));
}

export function buildStoredZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const name = encoder.encode(entry.name);
    const data = toBytes(entry.data);
    const checksum = crc32(data);
    const local = Buffer.alloc(30 + name.byteLength + data.byteLength);
    const central = Buffer.alloc(46 + name.byteLength);

    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0x0021, 12);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(data.byteLength, 18);
    local.writeUInt32LE(data.byteLength, 22);
    local.writeUInt16LE(name.byteLength, 26);
    local.writeUInt16LE(0, 28);
    Buffer.from(name).copy(local, 30);
    Buffer.from(data).copy(local, 30 + name.byteLength);

    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0x0021, 14);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(data.byteLength, 20);
    central.writeUInt32LE(data.byteLength, 24);
    central.writeUInt16LE(name.byteLength, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    Buffer.from(name).copy(central, 46);

    localParts.push(local);
    centralParts.push(central);
    offset += local.byteLength;
  }

  const centralDirectory = concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralDirectory.byteLength, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);

  return new Uint8Array(concat([...localParts, centralDirectory, end]));
}

function xml(strings, ...values) {
  return strings.reduce((result, part, index) => result + part + (values[index] ?? ""), "");
}

function entry(name, data) {
  return { name, data: toBytes(data) };
}

function contentTypes(format, slideCount, macro) {
  const slideOverrides = Array.from(
    { length: slideCount },
    (_, index) =>
      `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
  ).join("");
  const macroOverride = macro
    ? '<Override PartName="/ppt/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
    : "";

  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="${MAIN_CONTENT_TYPES[format]}"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
  <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  ${slideOverrides}
  ${macroOverride}
</Types>`;
}

function rootRelationships() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdOfficeDocument" Type="${OFFICE_RELS_NS}/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rIdCoreProperties" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rIdAppProperties" Type="${OFFICE_RELS_NS}/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rIdUnknownExtension" Type="https://native-powerpoint.invalid/relationships/fixture-extension" Target="customXml/native-powerpoint-extension.xml"/>
</Relationships>`;
}

function presentation(slideCount) {
  const slides = Array.from(
    { length: slideCount },
    (_, index) => `<p:sldId id="${256 + index}" r:id="rIdSlide${index + 1}"/>`,
  ).join("");

  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}" saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst>${slides}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function presentationRelationships(slideCount, macro) {
  const slides = Array.from(
    { length: slideCount },
    (_, index) =>
      `<Relationship Id="rIdSlide${index + 1}" Type="${OFFICE_RELS_NS}/slide" Target="slides/slide${index + 1}.xml"/>`,
  ).join("");
  const macroRelationship = macro
    ? `<Relationship Id="rIdVbaProject" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`
    : "";

  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdMaster" Type="${OFFICE_RELS_NS}/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdPresProps" Type="${OFFICE_RELS_NS}/presProps" Target="presProps.xml"/>
  <Relationship Id="rIdViewProps" Type="${OFFICE_RELS_NS}/viewProps" Target="viewProps.xml"/>
  <Relationship Id="rIdTheme" Type="${OFFICE_RELS_NS}/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rIdTableStyles" Type="${OFFICE_RELS_NS}/tableStyles" Target="tableStyles.xml"/>
  ${slides}
  ${macroRelationship}
</Relationships>`;
}

function baseShape(id, name, text, x, y, cx, cy) {
  return xml`<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2200"/><a:t>${text}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`;
}

function slideTreePrefix() {
  return xml`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;
}

function featureSlide() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${DRAWING_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}" xmlns:np="https://native-powerpoint.invalid/fixture">
  <p:cSld><p:spTree>
    ${slideTreePrefix()}
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Fixture title"><a:hlinkClick r:id="rIdHyperlink"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="342900"/><a:ext cx="5943600" cy="685800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Native PowerPoint fixture</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody>
    </p:sp>
    <p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="3" name="Fixture group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="2743200" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="914400"/></a:xfrm></p:grpSpPr>
      ${baseShape(4, "Grouped shape", "Grouped shape", 0, 0, 2743200, 914400)}
    </p:grpSp>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="5" name="Fixture image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="7772400" y="457200"/><a:ext cx="1828800" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="6" name="Fixture chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="457200" y="2971800"/><a:ext cx="3657600" cy="2286000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </p:graphicFrame>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="7" name="Fixture table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="5029200" y="2971800"/><a:ext cx="5943600" cy="1828800"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="2971800"/><a:gridCol w="2971800"/></a:tblGrid>
          <a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Metric</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
          <a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Slides</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
        </a:tbl>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
  <p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/></p:par></p:tnLst></p:timing>
  <p:extLst><p:ext uri="{57B35768-491A-4F9E-9E38-F0F5933D1274}"><np:feature preserve="unknown-ooxml"/></p:ext></p:extLst>
</p:sld>`;
}

function simpleSlide(index) {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld><p:spTree>
    ${slideTreePrefix()}
    ${baseShape(2, `Slide ${index} title`, `Large deck slide ${index}`, 457200, 457200, 7315200, 914400)}
  </p:spTree></p:cSld>
</p:sld>`;
}

function featureSlideRelationships() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdLayout" Type="${OFFICE_RELS_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdNotes" Type="${OFFICE_RELS_NS}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rIdImage" Type="${OFFICE_RELS_NS}/image" Target="../media/image1.png"/>
  <Relationship Id="rIdChart" Type="${OFFICE_RELS_NS}/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rIdHyperlink" Type="${OFFICE_RELS_NS}/hyperlink" Target="https://example.com/native-powerpoint" TargetMode="External"/>
</Relationships>`;
}

function simpleSlideRelationships() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdLayout" Type="${OFFICE_RELS_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
}

function slideMaster() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld name="Fixture master"><p:spTree>${slideTreePrefix()}</p:spTree></p:cSld>
  <p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rIdLayout"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`;
}

function slideLayout() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>${slideTreePrefix()}</p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
}

function theme() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${DRAWING_NS}" name="Native PowerPoint Fixture Theme">
  <a:themeElements>
    <a:clrScheme name="Fixture colors">
      <a:dk1><a:srgbClr val="1F2937"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="334155"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="0F766E"/></a:accent1><a:accent2><a:srgbClr val="B45309"/></a:accent2>
      <a:accent3><a:srgbClr val="2563EB"/></a:accent3><a:accent4><a:srgbClr val="BE123C"/></a:accent4>
      <a:accent5><a:srgbClr val="7C3AED"/></a:accent5><a:accent6><a:srgbClr val="15803D"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Fixture fonts"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Fixture formatting"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

function notesSlide() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld><p:spTree>${slideTreePrefix()}<p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Speaker notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Fixture speaker notes survive round trip.</a:t></a:r></a:p></p:txBody>
  </p:sp></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>`;
}

function chart() {
  return xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_NS}">
  <c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
    <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Fixture data</c:v></c:tx>
      <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Slides</c:v></c:pt></c:strLit></c:cat>
      <c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
    </c:ser><c:axId val="123456"/><c:axId val="654321"/>
  </c:barChart><c:catAx><c:axId val="123456"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="654321"/></c:catAx><c:valAx><c:axId val="654321"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="123456"/></c:valAx></c:plotArea></c:chart>
</c:chartSpace>`;
}

export function createDeckEntries({ format = "pptx", slideCount = 1, richFirstSlide = true } = {}) {
  if (!(format in MAIN_CONTENT_TYPES)) {
    throw new Error(`Unsupported fixture format: ${format}`);
  }
  if (!Number.isInteger(slideCount) || slideCount < 1) {
    throw new Error("Fixture decks must contain at least one slide.");
  }

  const macro = MACRO_FORMATS.has(format);
  const entries = [
    entry("[Content_Types].xml", contentTypes(format, slideCount, macro)),
    entry("_rels/.rels", rootRelationships()),
    entry("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Native PowerPoint fixture</dc:title></cp:coreProperties>'),
    entry("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Native PowerPoint tests</Application></Properties>'),
    entry("ppt/presentation.xml", presentation(slideCount)),
    entry("ppt/_rels/presentation.xml.rels", presentationRelationships(slideCount, macro)),
    entry("ppt/presProps.xml", `<?xml version="1.0" encoding="UTF-8"?><p:presentationPr xmlns:p="${PRESENTATION_NS}"/>`),
    entry("ppt/viewProps.xml", `<?xml version="1.0" encoding="UTF-8"?><p:viewPr xmlns:p="${PRESENTATION_NS}"/>`),
    entry("ppt/tableStyles.xml", `<?xml version="1.0" encoding="UTF-8"?><a:tblStyleLst xmlns:a="${DRAWING_NS}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`),
    entry("ppt/theme/theme1.xml", theme()),
    entry("ppt/slideMasters/slideMaster1.xml", slideMaster()),
    entry("ppt/slideMasters/_rels/slideMaster1.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rIdLayout" Type="${OFFICE_RELS_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rIdTheme" Type="${OFFICE_RELS_NS}/theme" Target="../theme/theme1.xml"/></Relationships>`),
    entry("ppt/slideLayouts/slideLayout1.xml", slideLayout()),
    entry("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rIdMaster" Type="${OFFICE_RELS_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`),
    entry("ppt/notesMasters/notesMaster1.xml", `<?xml version="1.0" encoding="UTF-8"?><p:notesMaster xmlns:a="${DRAWING_NS}" xmlns:p="${PRESENTATION_NS}"><p:cSld><p:spTree>${slideTreePrefix()}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/></p:notesMaster>`),
    entry("ppt/notesMasters/_rels/notesMaster1.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rIdTheme" Type="${OFFICE_RELS_NS}/theme" Target="../theme/theme1.xml"/></Relationships>`),
    entry("ppt/notesSlides/notesSlide1.xml", notesSlide()),
    entry("ppt/notesSlides/_rels/notesSlide1.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rIdNotesMaster" Type="${OFFICE_RELS_NS}/notesMaster" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rIdSlide" Type="${OFFICE_RELS_NS}/slide" Target="../slides/slide1.xml"/></Relationships>`),
    entry("ppt/charts/chart1.xml", chart()),
    entry("ppt/media/image1.png", FIXTURE_PNG),
    entry("customXml/native-powerpoint-extension.xml", '<?xml version="1.0" encoding="UTF-8"?><np:fixture xmlns:np="https://native-powerpoint.invalid/fixture" preserve="unknown-ooxml"><np:data>Do not discard this part.</np:data></np:fixture>'),
  ];

  for (let index = 1; index <= slideCount; index += 1) {
    const isRichSlide = index === 1 && richFirstSlide;
    entries.push(entry(`ppt/slides/slide${index}.xml`, isRichSlide ? featureSlide() : simpleSlide(index)));
    entries.push(entry(`ppt/slides/_rels/slide${index}.xml.rels`, isRichSlide ? featureSlideRelationships() : simpleSlideRelationships()));
  }

  if (macro) {
    entries.push(entry("ppt/vbaProject.bin", "Native PowerPoint inert macro-preservation fixture.\n"));
  }

  return entries;
}

export function createDeck(options) {
  return buildStoredZip(createDeckEntries(options));
}
