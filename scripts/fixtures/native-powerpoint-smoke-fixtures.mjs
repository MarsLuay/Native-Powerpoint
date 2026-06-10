import { mkdir, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const JSZip = require('jszip');

const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(SCRIPT_DIR, '..', '..');
const FIXTURE_DIR = path.join(PROJECT_ROOT, 'test-results', 'native-powerpoint-fixtures');
const TABLE_CHART_FIXTURE = path.join(FIXTURE_DIR, 'table-and-editable-chart.pptx');
const EXTERNAL_CHART_FIXTURE = path.join(FIXTURE_DIR, 'external-chart-link.pptx');

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_RELS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const SLIDE_CX = 9144000;
const SLIDE_CY = 5143500;

function xmlDeclaration(contents) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${contents}`;
}

function zipText(zip, name, contents) {
  zip.file(name, xmlDeclaration(contents));
}

async function buildWorkbook() {
  const workbook = new JSZip();
  zipText(workbook, '[Content_Types].xml', `
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`);
  zipText(workbook, '_rels/.rels', `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zipText(workbook, 'xl/workbook.xml', `
<workbook xmlns="${SHEET_NS}" xmlns:r="${OFFICE_RELS_NS}">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <calcPr calcId="191029"/>
</workbook>`);
  zipText(workbook, 'xl/_rels/workbook.xml.rels', `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);
  zipText(workbook, 'xl/worksheets/sheet1.xml', `
<worksheet xmlns="${SHEET_NS}">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Category</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Samples</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>1g</t></is></c>
      <c r="B2"><v>1.04</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>2g</t></is></c>
      <c r="B3"><v>2.08</v></c>
    </row>
  </sheetData>
</worksheet>`);

  return workbook.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function contentTypes(slideCount) {
  const slides = Array.from({ length: slideCount }, (_, index) => `
  <Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');

  return `
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>${slides}
</Types>`;
}

function presentation(slideCount) {
  const slideEntries = Array.from({ length: slideCount }, (_, index) => `
    <p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`).join('');

  return `
<p:presentation xmlns:a="${A_NS}" xmlns:p="${P_NS}" xmlns:r="${OFFICE_RELS_NS}">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>${slideEntries}
  </p:sldIdLst>
  <p:sldSz cx="${SLIDE_CX}" cy="${SLIDE_CY}" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function presentationRelationships(slideCount) {
  const slideRelationships = Array.from({ length: slideCount }, (_, index) => `
  <Relationship Id="rId${index + 2}" Type="${OFFICE_RELS_NS}/slide" Target="slides/slide${index + 1}.xml"/>`).join('');

  return `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideMaster" Target="slideMasters/slideMaster1.xml"/>${slideRelationships}
</Relationships>`;
}

function slideMaster() {
  return `
<p:sldMaster xmlns:a="${A_NS}" xmlns:p="${P_NS}" xmlns:r="${OFFICE_RELS_NS}">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle/>
    <p:bodyStyle/>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>`;
}

function slideMasterRelationships() {
  return `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="${OFFICE_RELS_NS}/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
}

function slideLayout() {
  return `
<p:sldLayout xmlns:a="${A_NS}" xmlns:p="${P_NS}" xmlns:r="${OFFICE_RELS_NS}" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
}

function slideLayoutRelationships() {
  return `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}

function theme() {
  return `
<a:theme xmlns:a="${A_NS}" name="Smoke Fixtures">
  <a:themeElements>
    <a:clrScheme name="Smoke">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2>
      <a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="70AD47"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="A5A5A5"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Smoke">
      <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Smoke">
      <a:fillStyleLst><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:srgbClr val="D9E2F3"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`;
}

function slide(slideNumber, bodyXml = '') {
  return `
<p:sld xmlns:a="${A_NS}" xmlns:p="${P_NS}" xmlns:r="${OFFICE_RELS_NS}">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
        </a:xfrm>
      </p:grpSpPr>${bodyXml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function slideRelationships(extraRelationships = '') {
  return `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>${extraRelationships}
</Relationships>`;
}

function tableShape() {
  return `
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="2" name="Fixture Table"/>
          <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="914400" y="914400"/>
          <a:ext cx="7315200" cy="1828800"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr firstRow="1" bandRow="1"/>
              <a:tblGrid>
                <a:gridCol w="3657600"/>
                <a:gridCol w="3657600"/>
              </a:tblGrid>
              <a:tr h="609600">
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Samples</a:t></a:r><a:endParaRPr lang="en-US" sz="1800"/></a:p>
                  </a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="C5FEDC"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Value</a:t></a:r><a:endParaRPr lang="en-US" sz="1800"/></a:p>
                  </a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="EAF7F0"/></a:solidFill></a:tcPr>
                </a:tc>
              </a:tr>
              <a:tr h="609600">
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>1g</a:t></a:r><a:endParaRPr lang="en-US" sz="1800"/></a:p>
                  </a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>1.04</a:t></a:r><a:endParaRPr lang="en-US" sz="1800"/></a:p>
                  </a:txBody>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr>
                </a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>`;
}

function chartShape() {
  return `
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="2" name="Fixture Chart"/>
          <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="914400" y="685800"/>
          <a:ext cx="7315200" cy="3657600"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="${C_NS}" xmlns:r="${OFFICE_RELS_NS}" r:id="rId2"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>`;
}

function chartXml() {
  return `
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${OFFICE_RELS_NS}">
  <c:lang val="en-US"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Samples</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
          </c:spPr>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$3</c:f>
              <c:strCache>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>1g</c:v></c:pt>
                <c:pt idx="1"><c:v>2g</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$3</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>1.04</c:v></c:pt>
                <c:pt idx="1"><c:v>2.08</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:gapWidth val="150"/>
        <c:axId val="12345678"/>
        <c:axId val="87654321"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="12345678"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="87654321"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="87654321"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="12345678"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:externalData r:id="rId1">
    <c:autoUpdate val="0"/>
  </c:externalData>
</c:chartSpace>`;
}

function chartRelationships(external) {
  const target = external ? 'file:///tmp/native-powerpoint-smoke-external.xlsx' : '../embeddings/Microsoft_Excel_Worksheet.xlsx';
  const targetMode = external ? ' TargetMode="External"' : '';

  return `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/package" Target="${target}"${targetMode}/>
</Relationships>`;
}

async function buildPresentation({ externalWorkbook = false, slideCount = 5 } = {}) {
  const zip = new JSZip();
  const workbook = externalWorkbook ? null : await buildWorkbook();

  zipText(zip, '[Content_Types].xml', contentTypes(slideCount));
  zipText(zip, '_rels/.rels', `
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);
  zipText(zip, 'ppt/presentation.xml', presentation(slideCount));
  zipText(zip, 'ppt/_rels/presentation.xml.rels', presentationRelationships(slideCount));
  zipText(zip, 'ppt/slideMasters/slideMaster1.xml', slideMaster());
  zipText(zip, 'ppt/slideMasters/_rels/slideMaster1.xml.rels', slideMasterRelationships());
  zipText(zip, 'ppt/slideLayouts/slideLayout1.xml', slideLayout());
  zipText(zip, 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', slideLayoutRelationships());
  zipText(zip, 'ppt/theme/theme1.xml', theme());
  zipText(zip, 'ppt/charts/chart1.xml', chartXml());
  zipText(zip, 'ppt/charts/_rels/chart1.xml.rels', chartRelationships(externalWorkbook));
  if (workbook) {
    zip.file('ppt/embeddings/Microsoft_Excel_Worksheet.xlsx', workbook);
  }

  for (let slideNumber = 1; slideNumber <= slideCount; slideNumber += 1) {
    const hasTable = !externalWorkbook && slideNumber === 4;
    const hasChart = slideNumber === (externalWorkbook ? 1 : 5);
    const body = `${hasTable ? tableShape() : ''}${hasChart ? chartShape() : ''}`;
    const chartRelationship = hasChart
      ? `
  <Relationship Id="rId2" Type="${OFFICE_RELS_NS}/chart" Target="../charts/chart1.xml"/>`
      : '';
    zipText(zip, `ppt/slides/slide${slideNumber}.xml`, slide(slideNumber, body));
    zipText(zip, `ppt/slides/_rels/slide${slideNumber}.xml.rels`, slideRelationships(chartRelationship));
  }

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function prepareNativePowerPointSmokeFixtures() {
  await mkdir(FIXTURE_DIR, { recursive: true });
  await Promise.all([
    writeFile(TABLE_CHART_FIXTURE, await buildPresentation()),
    writeFile(EXTERNAL_CHART_FIXTURE, await buildPresentation({ externalWorkbook: true, slideCount: 1 }))
  ]);

  return {
    tableSample: TABLE_CHART_FIXTURE,
    chartSample: TABLE_CHART_FIXTURE,
    chartDataSample: TABLE_CHART_FIXTURE,
    externalChartSample: EXTERNAL_CHART_FIXTURE
  };
}
