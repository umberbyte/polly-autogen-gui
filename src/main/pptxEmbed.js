const fs = require('node:fs/promises');
const path = require('node:path');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const mm = require('music-metadata');

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

const REL_TYPE_AUDIO = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio';
const REL_TYPE_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media';
const REL_TYPE_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const ICON_MEDIA_PATH = 'ppt/media/pollyAudioIcon.png';
// Generic play-button icon (dark circle, white triangle), 128x128 RGBA PNG
const ICON_PNG_BASE64 =
  "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";

const ICON_SIZE_EMU = 762000;
const ICON_MARGIN_EMU = 137160;
const DEFAULT_SLIDE_WIDTH_EMU = 12192000;
const DEFAULT_SLIDE_HEIGHT_EMU = 6858000;

const ADVANCE_MARGIN_MS = 3000;

function parseXml(xml) {
  return new DOMParser().parseFromString(xml, 'text/xml');
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

async function createWorkingCopy(pptxPath) {
  const parsed = path.parse(pptxPath);
  const workingPath = path.join(parsed.dir, `${parsed.name}_音声埋め込み.pptx`);
  await fs.copyFile(pptxPath, workingPath);
  return workingPath;
}

async function listMp3sByPageNumber(mp3Folder) {
  const entries = await fs.readdir(mp3Folder, { withFileTypes: true });
  const byPage = new Map();
  const unparsable = [];
  for (const entry of entries) {
    if (!entry.isFile() || path.extname(entry.name).toLowerCase() !== '.mp3') continue;
    const baseName = path.basename(entry.name, path.extname(entry.name));
    const lastUnderscore = baseName.lastIndexOf('_');
    if (lastUnderscore === -1) {
      unparsable.push(entry.name);
      continue;
    }
    const pageNumber = baseName.slice(lastUnderscore + 1);
    byPage.set(pageNumber, path.join(mp3Folder, entry.name));
  }
  return { byPage, unparsable };
}

async function getOrderedSlideParts(zip) {
  const presentationXml = await zip.file('ppt/presentation.xml').async('string');
  const relsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  const presDoc = parseXml(presentationXml);
  const relsDoc = parseXml(relsXml);

  const relIdToTarget = {};
  const relNodes = relsDoc.getElementsByTagNameNS(NS_RELS, 'Relationship');
  for (let i = 0; i < relNodes.length; i += 1) {
    relIdToTarget[relNodes[i].getAttribute('Id')] = relNodes[i].getAttribute('Target');
  }

  const sldIdNodes = presDoc.getElementsByTagNameNS(NS_P, 'sldId');
  const parts = [];
  for (let i = 0; i < sldIdNodes.length; i += 1) {
    const rId = sldIdNodes[i].getAttributeNS(NS_R, 'id');
    const target = relIdToTarget[rId];
    if (!target) continue;
    const normalized = target.replace(/^\/?ppt\//, '').replace(/^\.\.\//, '');
    parts.push(`ppt/${normalized}`);
  }
  return parts;
}

async function getSlideSize(zip) {
  const presentationXml = await zip.file('ppt/presentation.xml').async('string');
  const presDoc = parseXml(presentationXml);
  const sldSzNodes = presDoc.getElementsByTagNameNS(NS_P, 'sldSz');
  if (sldSzNodes.length === 0) {
    return { width: DEFAULT_SLIDE_WIDTH_EMU, height: DEFAULT_SLIDE_HEIGHT_EMU };
  }
  const cx = Number(sldSzNodes[0].getAttribute('cx'));
  const cy = Number(sldSzNodes[0].getAttribute('cy'));
  return {
    width: Number.isFinite(cx) && cx > 0 ? cx : DEFAULT_SLIDE_WIDTH_EMU,
    height: Number.isFinite(cy) && cy > 0 ? cy : DEFAULT_SLIDE_HEIGHT_EMU,
  };
}

function ensureContentTypes(zip, contentTypesXml) {
  const doc = parseXml(contentTypesXml);
  const root = doc.documentElement;
  const defaults = doc.getElementsByTagNameNS(NS_CT, 'Default');
  const existing = new Set();
  for (let i = 0; i < defaults.length; i += 1) {
    existing.add(defaults[i].getAttribute('Extension').toLowerCase());
  }
  const needed = [
    ['mp3', 'audio/mpeg'],
    ['png', 'image/png'],
  ];
  let changed = false;
  for (const [ext, contentType] of needed) {
    if (existing.has(ext)) continue;
    const el = doc.createElementNS(NS_CT, 'Default');
    el.setAttribute('Extension', ext);
    el.setAttribute('ContentType', contentType);
    root.insertBefore(el, root.firstChild);
    changed = true;
  }
  return changed ? serializeXml(doc) : contentTypesXml;
}

function nextRelId(relsDoc) {
  const relNodes = relsDoc.getElementsByTagNameNS(NS_RELS, 'Relationship');
  let max = 0;
  for (let i = 0; i < relNodes.length; i += 1) {
    const match = /^rId(\d+)$/.exec(relNodes[i].getAttribute('Id') || '');
    if (match) max = Math.max(max, Number(match[1]));
  }
  return `rId${max + 1}`;
}

function addRelationship(relsDoc, type, target) {
  const id = nextRelId(relsDoc);
  const el = relsDoc.createElementNS(NS_RELS, 'Relationship');
  el.setAttribute('Id', id);
  el.setAttribute('Type', type);
  el.setAttribute('Target', target);
  relsDoc.documentElement.appendChild(el);
  return id;
}

function nextShapeId(slideDoc) {
  const cNvPrNodes = slideDoc.getElementsByTagNameNS(NS_P, 'cNvPr');
  let max = 0;
  for (let i = 0; i < cNvPrNodes.length; i += 1) {
    const id = Number(cNvPrNodes[i].getAttribute('id'));
    if (!Number.isNaN(id)) max = Math.max(max, id);
  }
  return max + 1;
}

function buildPicFragment({ shapeId, name, rIdAudio, rIdMedia, rIdImage, offX, offY, size }) {
  return `<p:pic xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
  <p:nvPicPr>
    <p:cNvPr id="${shapeId}" name="${name}">
      <a:hlinkClick r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr>
      <a:audioFile r:link="${rIdAudio}"/>
      <p:extLst>
        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
          <p14:media xmlns:p14="${NS_P14}" r:embed="${rIdMedia}"/>
        </p:ext>
      </p:extLst>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="${rIdImage}"/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${offX}" y="${offY}"/>
      <a:ext cx="${size}" cy="${size}"/>
    </a:xfrm>
    <a:prstGeom prst="ellipse">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
</p:pic>`;
}

function buildTimingFragment({ shapeId, durationMs }) {
  return `<p:timing xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <p:par>
                  <p:cTn id="3" fill="hold">
                    <p:stCondLst>
                      <p:cond delay="indefinite"/>
                      <p:cond evt="onBegin" delay="0">
                        <p:tn val="2"/>
                      </p:cond>
                    </p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="4" fill="hold">
                          <p:stCondLst>
                            <p:cond delay="0"/>
                          </p:stCondLst>
                          <p:childTnLst>
                            <p:par>
                              <p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="withEffect">
                                <p:stCondLst>
                                  <p:cond delay="0"/>
                                </p:stCondLst>
                                <p:childTnLst>
                                  <p:cmd type="call" cmd="playFrom(0.0)">
                                    <p:cBhvr>
                                      <p:cTn id="6" dur="${durationMs}" fill="hold"/>
                                      <p:tgtEl>
                                        <p:spTgt spid="${shapeId}"/>
                                      </p:tgtEl>
                                    </p:cBhvr>
                                  </p:cmd>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
            <p:prevCondLst>
              <p:cond evt="onPrev" delay="0">
                <p:tgtEl>
                  <p:sldTgt/>
                </p:tgtEl>
              </p:cond>
            </p:prevCondLst>
            <p:nextCondLst>
              <p:cond evt="onNext" delay="0">
                <p:tgtEl>
                  <p:sldTgt/>
                </p:tgtEl>
              </p:cond>
            </p:nextCondLst>
          </p:seq>
          <p:audio>
            <p:cMediaNode vol="80000">
              <p:cTn id="7" fill="hold" display="0">
                <p:stCondLst>
                  <p:cond delay="indefinite"/>
                </p:stCondLst>
                <p:endCondLst>
                  <p:cond evt="onStopAudio" delay="0">
                    <p:tgtEl>
                      <p:sldTgt/>
                    </p:tgtEl>
                  </p:cond>
                </p:endCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt spid="${shapeId}"/>
              </p:tgtEl>
            </p:cMediaNode>
          </p:audio>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>`;
}

function importFragment(targetDoc, fragmentXml) {
  const fragDoc = parseXml(fragmentXml);
  return targetDoc.importNode(fragDoc.documentElement, true);
}

function firstChildNS(el, ns, localName) {
  const nodes = el.getElementsByTagNameNS(ns, localName);
  return nodes.length > 0 ? nodes[0] : null;
}

function setTransitionAdvance(slideDoc, advTmMs) {
  const sld = slideDoc.documentElement;
  let transition = null;
  for (let node = sld.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1 && node.namespaceURI === NS_P && node.localName === 'transition') {
      transition = node;
      break;
    }
  }
  if (!transition) {
    transition = slideDoc.createElementNS(NS_P, 'p:transition');
    // schema order: cSld, clrMapOvr?, transition?, timing?, extLst?
    let insertBefore = null;
    for (let node = sld.firstChild; node; node = node.nextSibling) {
      if (
        node.nodeType === 1 &&
        node.namespaceURI === NS_P &&
        (node.localName === 'timing' || node.localName === 'extLst')
      ) {
        insertBefore = node;
        break;
      }
    }
    if (insertBefore) sld.insertBefore(transition, insertBefore);
    else sld.appendChild(transition);
  }
  transition.setAttribute('advTm', String(advTmMs));
}

function insertTiming(slideDoc, timingNode) {
  const sld = slideDoc.documentElement;
  let insertBefore = null;
  for (let node = sld.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1 && node.namespaceURI === NS_P && node.localName === 'extLst') {
      insertBefore = node;
      break;
    }
  }
  if (insertBefore) sld.insertBefore(timingNode, insertBefore);
  else sld.appendChild(timingNode);
}

function hasExistingTiming(slideDoc) {
  const sld = slideDoc.documentElement;
  for (let node = sld.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1 && node.namespaceURI === NS_P && node.localName === 'timing') {
      return true;
    }
  }
  return false;
}

async function embedAudioIntoPptx(workingPptxPath, mp3Folder, onProgress = () => {}) {
  const pptxBuffer = await fs.readFile(workingPptxPath);
  const zip = await JSZip.loadAsync(pptxBuffer);

  const contentTypesXml = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ensureContentTypes(zip, contentTypesXml));

  const slideParts = await getOrderedSlideParts(zip);
  const { byPage, unparsable } = await listMp3sByPageNumber(mp3Folder);
  const slideSize = await getSlideSize(zip);
  const iconOffX = slideSize.width - ICON_SIZE_EMU - ICON_MARGIN_EMU;
  const iconOffY = slideSize.height - ICON_SIZE_EMU - ICON_MARGIN_EMU;

  const warnings = [];
  for (const name of unparsable) {
    warnings.push(`ファイル名からページ番号を判別できないためスキップしました: ${name}`);
  }

  if (!zip.file(ICON_MEDIA_PATH)) {
    zip.file(ICON_MEDIA_PATH, Buffer.from(ICON_PNG_BASE64, 'base64'));
  }

  const total = slideParts.length;
  onProgress({ type: 'start', total });

  const usedPages = new Set();
  let embeddedCount = 0;
  let mediaCounter = 0;

  for (let i = 0; i < slideParts.length; i += 1) {
    const slidePart = slideParts[i];
    const pageNumber = String(i + 1);
    onProgress({ type: 'slide-start', index: i + 1, total, pageNumber });

    const mp3Path = byPage.get(pageNumber);
    if (!mp3Path) {
      warnings.push(`スライド${pageNumber}に対応するmp3が見つからないためスキップしました`);
      onProgress({ type: 'slide-skip', index: i + 1, total, pageNumber });
      continue;
    }
    usedPages.add(pageNumber);

    try {
      const meta = await mm.parseFile(mp3Path);
      const durationSec = meta.format.duration || 0;
      const durationMs = Math.round(durationSec * 1000);

      mediaCounter += 1;
      const mediaFileName = `audioEmbed${mediaCounter}.mp3`;
      const mediaPartPath = `ppt/media/${mediaFileName}`;
      zip.file(mediaPartPath, await fs.readFile(mp3Path));

      const slideFileName = path.posix.basename(slidePart);
      const relsPartPath = `ppt/slides/_rels/${slideFileName}.rels`;
      const relsFile = zip.file(relsPartPath);
      const relsXml = relsFile
        ? await relsFile.async('string')
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_RELS}"></Relationships>`;
      const relsDoc = parseXml(relsXml);

      const rIdAudio = addRelationship(relsDoc, REL_TYPE_AUDIO, `../media/${mediaFileName}`);
      const rIdMedia = addRelationship(relsDoc, REL_TYPE_MEDIA, `../media/${mediaFileName}`);
      const rIdImage = addRelationship(relsDoc, REL_TYPE_IMAGE, `../media/${path.posix.basename(ICON_MEDIA_PATH)}`);
      zip.file(relsPartPath, serializeXml(relsDoc));

      const slideXml = await zip.file(slidePart).async('string');
      const slideDoc = parseXml(slideXml);

      const shapeId = nextShapeId(slideDoc);
      const spTree = firstChildNS(slideDoc.documentElement, NS_P, 'spTree');
      if (!spTree) throw new Error('スライドXMLに<p:spTree>が見つかりません');
      spTree.appendChild(
        importFragment(
          slideDoc,
          buildPicFragment({
            shapeId,
            name: `Audio ${pageNumber}`,
            rIdAudio,
            rIdMedia,
            rIdImage,
            offX: iconOffX,
            offY: iconOffY,
            size: ICON_SIZE_EMU,
          }),
        ),
      );

      if (hasExistingTiming(slideDoc)) {
        warnings.push(
          `スライド${pageNumber}には既存のアニメーション(タイミング)設定があるため、自動再生の設定は追加しませんでした。音声自体は埋め込まれています(自動再生はPowerPoint側で手動設定してください)`,
        );
      } else {
        insertTiming(slideDoc, importFragment(slideDoc, buildTimingFragment({ shapeId, durationMs })));
      }

      setTransitionAdvance(slideDoc, durationMs + ADVANCE_MARGIN_MS);
      zip.file(slidePart, serializeXml(slideDoc));

      embeddedCount += 1;
      onProgress({ type: 'slide-done', index: i + 1, total, pageNumber, durationMs });
    } catch (error) {
      const message = error.message || String(error);
      warnings.push(`スライド${pageNumber}の埋め込みでエラー: ${message}`);
      onProgress({ type: 'slide-error', index: i + 1, total, pageNumber, error: message });
    }
  }

  for (const pageNumber of byPage.keys()) {
    if (!usedPages.has(pageNumber)) {
      warnings.push(`ページ番号${pageNumber}のmp3に対応するスライドが無いためスキップしました`);
    }
  }

  const outBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await fs.writeFile(workingPptxPath, outBuffer);

  onProgress({ type: 'done', total, embeddedCount });
  return { embeddedCount, warnings };
}

module.exports = {
  createWorkingCopy,
  getOrderedSlideParts,
  listMp3sByPageNumber,
  embedAudioIntoPptx,
};
