import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from 'xmldom';

const typeConverters = {
  str: (e) => e,
  int: (e) => e,
  float: (e) => e,
  Date: (e) => new Date(e).toString(),
  enumDocSecurity: (e) =>
    e == 0
      ? 'None'
      : e == 1
        ? 'Document is password protected.'
        : e == 2
          ? 'Document is recommended to be opened as read-only.'
          : e == 4
            ? 'Document is enforced to be opened as read-only.'
            : e == 8
              ? 'Document is locked for annotation.'
              : 'Unknown', //default
  bool: (e) => (e == 'false' ? 'No' : e == 'true' ? 'Yes' : 'Unknown'),
  ISO8601: (e) => (
    (e = /^P((\d+Y)?(\d+M)?(\d+W)?(\d+D)?)?(T(\d+H)?(\d+M)?(\d+S)?)?$/
      .exec(e)
      .map((t) =>
        Number(
          typeof t === 'undefined'
            ? 0
            : ((t = t.replace(/[^\d]/g, '')), t == '' ? 0 : t),
        ),
      )),
    (e = Math.floor(
      (e[2] * 31104000 +
        e[3] * 2592000 +
        e[4] * 604800 +
        e[5] * 246060 +
        e[7] * 3600 +
        e[8] * 60 +
        e[9]) /
        60,
    )),
    e == 1 ? e + ' minute' : e + ' minutes'
  ),
  intMinutes: (e) => (e == 1 ? e + ' minute' : e + ' minutes'),
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.extendedproperties(v=office.14).aspx
const properties = {
  'cp:category': { name: 'category', type: 'str' },
  Manager: { name: 'manager', type: 'str' },
  'cp:contentStatus': { name: 'contentStatus', type: 'str' },
  'dc:subject': { name: 'subject', type: 'str' },
  HyperlinkBase: { name: 'hyperlinkBase', type: 'str' },
  'Slide Titles': { name: 'slideTitles', type: 'str' },
  Theme: { name: 'theme', type: 'str' },
  Title: { name: 'titles', type: 'str' },
  'dc:title': { name: 'title', type: 'str' },
  'dc:creator': { name: 'creator', type: 'str' },
  'cp:keywords': { name: 'keywords', type: 'str' },
  'dc:description': { name: 'description', type: 'str' },
  'cp:lastModifiedBy': { name: 'lastModifiedBy', type: 'str' },
  'cp:revision': { name: 'revisionNumber', type: 'int' },
  'dcterms:created': { name: 'created', type: 'Date' },
  'dcterms:modified': { name: 'modified', type: 'Date' },
  Template: { name: 'template', type: 'str' },
  TotalTime: { name: 'totalTime', type: 'intMinutes' },
  Pages: { name: 'pages', type: 'int' },
  Words: { name: 'words', type: 'int' },
  Characters: { name: 'characters', type: 'int' },
  Application: { name: 'application', type: 'str' },
  DocSecurity: { name: 'docSecurity', type: 'enumDocSecurity' },
  Lines: { name: 'lines', type: 'int' },
  Paragraphs: { name: 'paragraphs', type: 'int' },
  ScaleCrop: { name: 'scaleCrop', type: 'bool' },
  Company: { name: 'company', type: 'str' },
  LinksUpToDate: { name: 'linksUpToDate', type: 'bool' },
  SharedDoc: { name: 'sharedDoc', type: 'bool' },
  HyperlinksChanged: { name: 'hyperlinksChanged', type: 'bool' },
  AppVersion: { name: 'appVersion', type: 'float' },
  CharactersWithSpaces: { name: 'charactersWithSpaces', type: 'int' },
  Slides: { name: 'slides', type: 'int' },
  Notes: { name: 'notes', type: 'str' },
  HiddenSlides: { name: 'hiddenSlides', type: 'int' },
  'dc:language': { name: 'language', type: 'str' },
  MMClips: { name: 'mmClips', type: 'str' },
  'cp:lastPrinted': { name: 'lastPrinted', type: 'Date' },
  PresentationFormat: { name: 'presentationFormat', type: 'str' },
  Worksheets: { name: 'worksheets', type: 'str' },
  'office:meta/meta:initial-creator': { name: 'creator', type: 'str' },
  'office:meta/dc:creator': { name: 'lastModifiedBy', type: 'str' },
  'office:meta/meta:creation-date': { name: 'created', type: 'Date' },
  'office:meta/dc:date': { name: 'modified', type: 'Date' },
  'office:meta/meta:template': { name: 'template', type: 'str' },
  'office:meta/meta:editing-cycles': { name: 'revision', type: 'int' },
  'office:meta/meta:editing-duration': { name: 'totalTime', type: 'ISO8601' },
  'office:meta/meta:document-statistic/@meta:page-count': {
    name: 'pages',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:paragraph-count': {
    name: 'paragraphs',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:word-count': {
    name: 'words',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:character-count': {
    name: 'characters',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:row-count': {
    name: 'rows',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:non-whitespace-character-count': {
    name: 'whitespaceCharacters',
    type: 'str',
  },
  'office:meta/meta:template/@xlink:href': { name: 'template', type: 'str' },
  'office:meta/meta:template/@xlink:type': {
    name: 'templateType',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:table-count': {
    name: 'tables',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:image-count': {
    name: 'images',
    type: 'str',
  },
  'office:meta/meta:document-statistic/@meta:object-count': {
    name: 'objects',
    type: 'str',
  },
  'office:meta/meta:generator': { name: 'application', type: 'str' },
};

export const mimeTypes: Record<string, string> = {
  //https://stackoverflow.com/questions/4212861/what-is-a-correct-mime-type-for-docx-pptx-etc
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  docm: 'application/vnd.ms-word.document.macroEnabled.12',
  dotm: 'application/vnd.ms-word.template.macroEnabled.12',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  xlsb: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
  ppsm: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12',
  pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
  xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
  xltm: 'application/vnd.ms-excel.template.macroEnabled.12',
  potx: 'application/vnd.openxmlformats-officedocument.presentationml.template',
  potm: 'application/vnd.ms-powerpoint.template.macroEnabled.12',
  odt: 'application/vnd.oasis.opendocument.text',
  odp: 'application/vnd.oasis.opendocument.presentation',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  ots: 'application/vnd.oasis.opendocument.spreadsheet-template',
  otp: 'application/vnd.oasis.opendocument.presentation-template',
  ott: 'application/vnd.oasis.opendocument.text-template',
};

async function getMetadataAsXML(zip: JSZip) {
  const OPformat = getFormat(zip);
  if (OPformat === 'office') {
    return [
      await getXmlFromZip(zip, 'docProps/core.xml'),
      await getXmlFromZip(zip, 'docProps/app.xml'),
    ];
  } else if (OPformat === 'openoffice') {
    return [await getXmlFromZip(zip, 'meta.xml')];
  }
  return null;
}

async function getXmlFromZip(zip: JSZip, fileName: string): Promise<Document> {
  const text = await zip.file(fileName).async('text');
  const xmlDoc = new DOMParser().parseFromString(text, 'text/xml');
  return xmlDoc;
}

async function loadFile(officeFile: Buffer<ArrayBufferLike>): Promise<JSZip> {
  const zip = await JSZip.loadAsync(officeFile);
  const OPformat = getFormat(zip);
  if (OPformat) {
    return zip;
  } else {
    throw new Error('Error: File not valid');
  }
}

function getFormat(zip: JSZip): 'office' | 'openoffice' | null {
  if (
    Object.prototype.hasOwnProperty.call(zip.files, 'docProps/core.xml') &&
    Object.prototype.hasOwnProperty.call(zip.files, 'docProps/app.xml')
  ) {
    return 'office';
  } else if (Object.prototype.hasOwnProperty.call(zip.files, 'meta.xml')) {
    return 'openoffice';
  }
  return null;
}

function translateMetadata(
  textObjects: any,
  names: any,
): { editable: any; readOnly: any } {
  let headingPairsAndParts: any[] = [];
  textObjects.forEach((e: any, i: number, a: any[]) => {
    if (e.path == 'HeadingPairs/vt:vector/vt:variant/vt:lpstr') {
      headingPairsAndParts.push({
        name: Object.prototype.hasOwnProperty.call(names, e.value)
          ? names[e.value].name
          : e.value.replace(/ /g, ''),
        length: a[i + 1].value,
        value: [],
      });
    }
    if (e.path == 'TitlesOfParts/vt:vector/vt:lpstr') {
      for (var i = 0; i < headingPairsAndParts.length; i++) {
        if (
          headingPairsAndParts[i]['value'].length <
          headingPairsAndParts[i]['length']
        ) {
          headingPairsAndParts[i]['value'].push(e.value);
          break;
        }
      }
    }
  });

  textObjects = textObjects.filter(
    (e) =>
      e.path != 'HeadingPairs/vt:vector/vt:variant/vt:lpstr' &&
      e.path != 'TitlesOfParts/vt:vector/vt:lpstr' &&
      e.path != 'HeadingPairs/vt:vector/vt:variant/vt:i4',
  );

  var createPropertyOrArray = (object: any, property: any, val: any) => {
    if (object.hasOwnProperty(property)) {
      if (object[property].value instanceof Array) {
        object[property].value.push(val.value);
        object[property].rvalue.push(val.tvalue);
      } else {
        object[property].tvalue = [object[property].tvalue, val.tvalue];
        object[property].value = [object[property].value, val.value];
      }
    } else {
      object[property] = val;
    }
  };

  const editable = {};
  textObjects.forEach((e) => {
    if (names.hasOwnProperty(e['path'])) {
      const tvalue = typeConverters[names[e['path']].type](e.value);
      createPropertyOrArray(editable, names[e['path']].name, {
        value: e.value,
        tvalue,
        xmlPath: e.path,
      });
    } else {
      createPropertyOrArray(editable, e['path'], {
        value: e.value,
        tvalue: e.value,
        xmlPath: e.path,
      });
    }
  });

  const readOnly = {};
  headingPairsAndParts.forEach((e) => {
    readOnly[e.name] = { value: e.value, tvalue: e.value };
  });

  return { editable, readOnly };
}

function getTextObjectsFromXML(xml: Document) {
  return getTextFromNodelist(xml.lastChild.childNodes);
}

//returns all textnodes as object{path:'',value:''} from node list
function getTextFromNodelist(
  nodes: NodeListOf<ChildNode>,
  name?: string,
  metaObjects?: any[],
) {
  if (typeof metaObjects === 'undefined') {
    metaObjects = [];
  }
  if (typeof name === 'undefined') {
    name = '';
  }
  Array.from(nodes).forEach(function (e) {
    if (e.childNodes.length == 1 && e.firstChild.nodeType === 3) {
      var metaObject = {
        path: (name + '/' + e.nodeName).slice(1),
        value: e.firstChild.textContent,
      };
      metaObjects.push(metaObject);
    } else if (e.childNodes.length > 0) {
      metaObjects = getTextFromNodelist(
        e.childNodes,
        name + '/' + e.nodeName,
        metaObjects,
      );
    } else {
      const metaObject = {
        path: (name + '/' + e.nodeName).slice(1),
        value: '',
      };
      if (
        metaObject.path == 'office:meta/meta:document-statistic' ||
        metaObject.path === 'office:meta/meta:template'
      ) {
        Array.from((e as any).attributes).forEach((attr: any) => {
          metaObjects.push({
            path: metaObject.path + '/@' + attr.name,
            value: attr.value,
          });
        });
      } else {
        metaObjects.push(metaObject);
      }
    }
  });
  return metaObjects;
}

function editXML(xml: any, metadata: any) {
  for (const key in metadata.editable) {
    const object = metadata.editable[key];
    if (object.xmlPath.includes('/@')) {
      var nodes = xml.getElementsByTagName(
        object.xmlPath.split('/').slice(-2, -1),
      );
      for (var i = 0; i < nodes.length; i++) {
        nodes[i].getAttributeNode(
          object.xmlPath.split('/').slice(-1)[0].replace('@', ''),
        ).value = object.value;
      }
    } else {
      var nodes = xml.getElementsByTagName(object.xmlPath.split('/').slice(-1));
      if (nodes.length > 0 && object.xmlPath != '') {
        for (var i = 0; i < nodes.length; i++) {
          var valueToInsert =
            object.value instanceof Array
              ? object.value[
                  i < object.value.length ? i : object.value.length - 1
                ]
              : object.value;
          if (
            nodes[i].childNodes.length > 0 &&
            nodes[i].firstChild.nodeType === 3
          ) {
            nodes[i].firstChild.data = valueToInsert;
          } else {
            nodes[i].appendChild(document.createTextNode(valueToInsert));
          }
        }
      }
    }
  }
  return xml;
}

async function getModifiedMetadataAsXml(
  officeFile: Buffer<ArrayBufferLike>,
  metadata: any,
) {
  const zip = await loadFile(officeFile);
  const xmls = await getMetadataAsXML(zip);
  return xmls.map((e) => editXML(e, metadata));
}

function getBlob(zip: JSZip): Promise<Buffer<ArrayBufferLike>> {
  return zip.generateAsync({ type: 'nodebuffer' });
}

function serializeXML(xml: any) {
  return new XMLSerializer().serializeToString(xml);
}

export async function editData(
  officeFile: Buffer<ArrayBufferLike>,
  metadata: any,
) {
  const newMetaFiles = await getModifiedMetadataAsXml(officeFile, metadata);
  const zip = await loadFile(officeFile);
  const OPformat = getFormat(zip);
  if (OPformat === 'office') {
    zip.remove('docProps/core.xml');
    zip.remove('docProps/app.xml');
    zip.file('docProps/core.xml', serializeXML(newMetaFiles[0]));
    zip.file('docProps/app.xml', serializeXML(newMetaFiles[1]));
  } else {
    zip.remove('meta.xml');
    zip.file('meta.xml', serializeXML(newMetaFiles[0]));
  }
  return getBlob(zip);
}

export async function removeData(officeFile: Buffer<ArrayBufferLike>) {
  const zip = await loadFile(officeFile);
  const OPformat = getFormat(zip);
  if (OPformat === 'office') {
    zip.remove('docProps/core.xml');
    zip.remove('docProps/app.xml');

    if (
      Object.prototype.hasOwnProperty.call(zip.files, 'docProps/custom.xml')
    ) {
      zip.remove('docProps/custom.xml');
    }
    const appXML =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>';
    const coreXML =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>';
    zip.file('docProps/core.xml', coreXML);
    zip.file('docProps/app.xml', appXML);
  } else if (OPformat === 'openoffice') {
    zip.remove('meta.xml');
    const metaXML =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.1"></office:document-meta>';
    zip.file('meta.xml', metaXML);
  } else {
    throw new Error('File not valid');
  }
  return getBlob(zip);
}

export async function getData(officeFile: Buffer<ArrayBufferLike>) {
  const zip = await loadFile(officeFile);
  const files = await getMetadataAsXML(zip);
  return translateMetadata(
    [].concat.apply(
      [],
      files.map((file) => getTextObjectsFromXML(file)),
    ),
    properties,
  );
}
