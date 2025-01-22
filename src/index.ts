import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from 'xmldom';
import { parse } from 'tinyduration';

type MetaObject = { path: string; value: string };
type HeadingPair = { name: string; length: number; value: string[] };

type Prop = {
  value: string;
  tvalue: any;
  xmlPath: string;
};

type Editable = Record<string, { value: string; tvalue: any; xmlPath: string }>;
type Readonly = Record<string, { value: string[]; tvalue: string[] }>;

type DOC_SECURITY_TYPE = '0' | '1' | '2' | '4' | '8';
const DOC_SECURITY: Record<DOC_SECURITY_TYPE, string> = {
  '0': 'None',
  '1': 'Document is password protected.',
  '2': 'Document is recommended to be opened as read-only.',
  '4': 'Document is enforced to be opened as read-only.',
  '8': 'Document is locked for annotation.',
};

function pluralizeMinute(num: number): string {
  return `${num.toString()} minute` + (num === 1 ? '' : 's');
}

type PROP_TYPE =
  | 'str'
  | 'int'
  | 'float'
  | 'Date'
  | 'enumDocSecurity'
  | 'bool'
  | 'ISO8601'
  | 'intMinutes';

const typeConverters: Record<PROP_TYPE, Function> = {
  str: (e: string) => e,
  int: (e: string) => e,
  float: (e: string) => e,
  Date: (e: string) => new Date(e).toString(),
  enumDocSecurity: (e: DOC_SECURITY_TYPE) => DOC_SECURITY[e] || 'Unknown',
  bool: (e: 'false' | 'true' | unknown) =>
    e === 'false' ? 'No' : e === 'true' ? 'Yes' : 'Unknown',
  ISO8601: (e: string) => {
    try {
      const duration = parse(e);

      const minutes =
        (duration.years || 0) * 525600 +
        (duration.months || 0) * 43200 +
        (duration.weeks || 0) * 10080 +
        (duration.days || 0) * 1440 +
        (duration.hours || 0) * 60 +
        (duration.minutes || 0) +
        Math.floor((duration.seconds || 0) / 60);
      return pluralizeMinute(minutes);
    } catch (err) {
      return '';
    }
  },
  intMinutes: (e: string) => pluralizeMinute(parseInt(e)),
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.extendedproperties(v=office.14).aspx
const PROPERTIES: Record<string, { name: string; type: PROP_TYPE }> = {
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

async function getMetadataAsXML(zip: JSZip): Promise<Document[]> {
  const OPformat = getFormat(zip);
  if (OPformat === 'office') {
    return [
      await getXmlFromZip(zip, 'docProps/core.xml'),
      await getXmlFromZip(zip, 'docProps/app.xml'),
    ];
  } else if (OPformat === 'openoffice') {
    return [await getXmlFromZip(zip, 'meta.xml')];
  }
  return [];
}

async function getXmlFromZip(zip: JSZip, fileName: string): Promise<Document> {
  const zipfile = zip.file(fileName);
  if (!zipfile) {
    throw new Error('Error: File not found');
  }
  const text = await zipfile.async('text');
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

function createPropertyOrArray(object: any, property: string, val: Prop) {
  if (object.hasOwnProperty(property)) {
    if (Array.isArray(object[property].value)) {
      object[property].value.push(val.value);
      // FIXME: is this a typo? should it be tvalue?
      object[property].rvalue.push(val.tvalue);
    } else {
      object[property].value = [object[property].value, val.value];
      object[property].tvalue = [object[property].tvalue, val.tvalue];
    }
  } else {
    object[property] = val;
  }
}

function translateMetadata(textObjects: MetaObject[]): {
  editable: Editable;
  readOnly: Readonly;
} {
  const headingPairsAndParts: HeadingPair[] = [];
  textObjects.map((element: MetaObject, idx: number, arr: MetaObject[]) => {
    if (element.path === 'HeadingPairs/vt:vector/vt:variant/vt:lpstr') {
      const name = !!PROPERTIES[element.value]
        ? PROPERTIES[element.value].name
        : element.value.replace(/ /g, '');
      const length = parseInt(arr[idx + 1].value);
      headingPairsAndParts.push({ name, length, value: [] });
    } else if (element.path === 'TitlesOfParts/vt:vector/vt:lpstr') {
      for (const pair of headingPairsAndParts) {
        if (pair.value.length < pair.length) {
          pair.value.push(element.value);
          break;
        }
      }
    }
  });

  const filteredTextObjects = textObjects.filter(
    (e) =>
      e.path !== 'HeadingPairs/vt:vector/vt:variant/vt:lpstr' &&
      e.path !== 'TitlesOfParts/vt:vector/vt:lpstr' &&
      e.path !== 'HeadingPairs/vt:vector/vt:variant/vt:i4',
  );

  const editable: Editable = {};
  for (const element of filteredTextObjects) {
    const prop = PROPERTIES[element.path];
    if (prop) {
      // known, defined property
      const tvalue = typeConverters[prop.type](element.value);
      createPropertyOrArray(editable, prop.name, {
        value: element.value,
        tvalue,
        xmlPath: element.path,
      });
    } else {
      createPropertyOrArray(editable, element.path, {
        value: element.value,
        tvalue: element.value,
        xmlPath: element.path,
      });
    }
  }

  const readOnly: Readonly = {};
  for (const e of headingPairsAndParts) {
    readOnly[e.name] = { value: e.value, tvalue: e.value };
  }

  return { editable, readOnly };
}

function getTextObjectsFromXML(xml: Document): MetaObject[] {
  if (!xml.lastChild) {
    return [];
  }
  return getTextFromNodelist(xml.lastChild.childNodes);
}

//returns all textnodes as object{path:'',value:''} from node list
function getTextFromNodelist(
  nodes: NodeListOf<ChildNode>,
  name: string = '',
  metaObjects: MetaObject[] = [],
): MetaObject[] {
  Array.from(nodes).forEach(function (e) {
    if (
      e.childNodes.length === 1 &&
      e.firstChild?.nodeType === Node.TEXT_NODE
    ) {
      const metaObject = {
        path: (name + '/' + e.nodeName).slice(1),
        value: e.firstChild.textContent!,
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
        metaObject.path === 'office:meta/meta:document-statistic' ||
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

function editXML(xml: Document, metadata: any): Document {
  for (const key in metadata.editable) {
    const object = metadata.editable[key];
    if (object.xmlPath.includes('/@')) {
      const tag = object.xmlPath.split('/').slice(-2, -1);
      const nodes = xml.getElementsByTagName(tag);
      for (let i = 0; i < nodes.length; i++) {
        nodes[i].getAttributeNode(
          object.xmlPath.split('/').slice(-1)[0].replace('@', ''),
        ).value = object.value;
      }
    } else {
      const nodes = xml.getElementsByTagName(
        object.xmlPath.split('/').slice(-1),
      );
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
): Promise<Document[]> {
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
  const payload = files.flatMap((file) => getTextObjectsFromXML(file));
  return translateMetadata(payload);
}
