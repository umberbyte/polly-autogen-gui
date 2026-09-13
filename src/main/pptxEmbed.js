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
// 1x1 transparent PNG
const ICON_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

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

function buildPicFragment({ shapeId, name, rIdAudio, rIdMedia, rIdImage }) {
  return `<p:pic xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
  <p:nvPicPr>
    <p:cNvPr id="${shapeId}" name="${name}" hidden="1">
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
      <a:off x="0" y="0"/>
      <a:ext cx="609600" cy="609600"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
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
