/**
 * @vitest-environment jsdom
 */

import * as OP from '../src/index.js';
import { readFile } from 'node:fs/promises';
import { it, expect } from 'vitest';

const filesPath = './src/test/';

it('Should parse metadata correctly', async () => {
  expect.assertions(4);

  const file = await readFile(filesPath + '1testdoc.docx');
  const metadata = await OP.getData(file);
  expect(metadata.editable.title.value).toEqual('A title');
  expect(metadata.editable.lastModifiedBy.value).toEqual(
    'Torkel Helland Velure',
  );
  expect(metadata.editable.manager.value).toEqual('Torkel Manager');
  expect(metadata.editable.manager.value).not.toEqual('wrong manager');
});

it('Should translate dates correctly', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + '1testdoc.docx');
  const metadata = await OP.getData(file);

  const created = new Date(metadata.editable.created.tvalue);
  const modified = new Date(metadata.editable.modified.tvalue);

  expect(created.toISOString()).toEqual('2018-03-16T15:33:00.000Z');
  expect(modified.toISOString()).toEqual('2018-03-16T15:36:00.000Z');
});

it('Should translate totalTime correctly', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + '1testdoc.docx');
  const metadata = await OP.getData(file);
  expect(metadata.editable.totalTime.tvalue).toEqual('3 minutes');
  expect(metadata.editable.totalTime.value).toEqual('3');
});

it('1: Should translate ISO8601-time to minutes', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Doc1.odt');
  const metadata = await OP.getData(file);
  expect(metadata.editable.totalTime.tvalue).toEqual('1 minute');
  expect(metadata.editable.totalTime.value).toEqual('PT60S');
});

it('2: Should translate ISO8601-time to minutes', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'longertime.odt');
  const metadata = await OP.getData(file);
  expect(metadata.editable.totalTime.tvalue).toEqual('3 minutes');
  expect(metadata.editable.totalTime.value).toEqual('PT3M43S');
});

it('Should translate docsecurity to correct string', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + '1testdoc.docx');
  const metadata = await OP.getData(file);
  expect(metadata.editable.docSecurity.tvalue).toEqual('None');
  expect(metadata.editable.docSecurity.value).toEqual('0');
});

it('Should create array for duplicate properties', async () => {
  expect.assertions(3);

  const file = await readFile(filesPath + 'multiprops.docx');
  const metadata = await OP.getData(file);
  expect(metadata.editable.company.value).toBeInstanceOf(Array);
  expect(metadata.editable.company.value[0]).toEqual(
    'University of Manchester',
  );
  expect(metadata.editable.company.value[1]).toEqual('Kilburn');
});

it('Should edit metadata properties correctly', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + '1testdoc.docx');
  const metadata = await OP.getData(file);
  metadata.editable.title.value = 'something else';
  const blob = await OP.editData(file, metadata);
  const meta = await OP.getData(blob);
  expect(meta.editable.title.value).toBe('something else');
  expect(meta.editable.title.value).not.toBe('not this');
});

it('Should edit metadata properties correctly when duplicate elements are present', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'multiprops.docx');
  const metadata = await OP.getData(file);
  metadata.editable.company.value = ['Something else', 'ABCDE'];
  const blob = await OP.editData(file, metadata);
  const meta = await OP.getData(blob);
  expect(meta.editable.company.value[0]).toBe('Something else');
  expect(meta.editable.company.value[1]).toBe('ABCDE');
});

it('Should edit attributes correctly', async () => {
  expect.assertions(4);

  const file = await readFile(filesPath + 'Doc1.odt');
  const metadata = await OP.getData(file);
  expect(metadata.editable.paragraphs.value).toBe('0');
  expect(metadata.editable.pages.value).toBe('1');

  metadata.editable.pages.value = '27';
  metadata.editable.paragraphs.value = '10';
  const blob = await OP.editData(file, metadata);
  const meta = await OP.getData(blob);
  expect(meta.editable.paragraphs.value).toBe('10');
  expect(meta.editable.pages.value).toBe('27');
});

it('Should create new textnode for edited empty node', async () => {
  expect.assertions(3);

  const file = await readFile(filesPath + '2testdoc.docx');
  const metadata = await OP.getData(file);
  expect(metadata.editable.company.value).toBe('');

  metadata.editable.company.value = 'Google';
  const blob = await OP.editData(file, metadata);
  const meta = await OP.getData(blob);
  expect(meta.editable.company.value).toBe('Google');
  expect(meta.editable.company.value).not.toBe('not this');
});

it('Should remove all metadata', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + '1testdoc.docx');
  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
  expect(Object.keys(meta.readOnly).length).toBe(0);
});

it('Should parse headingPairsAndParts correctly', async () => {
  expect.assertions(1);
  var slideTitles =
    'Apache Performance Tuning,Agenda,Introduction,Redundancy in Hardware,Server Configuration,Scaling Vertically,Scaling Vertically,Scaling Horizontally,Scaling Horizontally,Load Balancing Schemes,DNS Round-Robin,Example Zone File,Peer-based: NLB,Peer-based: Wackamole,Load Balancing Device,Load Balancing,Linux Virtual Server,Example: mod_proxy_balancer,Apache Configuration,Example: Tomcat, mod_jk,Apache Configuration,Tomcat Configuration,Problem: Session State,Solutions: Session State,Tomcat Session Replication,Session Replication Config,Caching Content,mod_cache Configuration,Make Popular Pages Static,Static Page Substitution,Tuning the Database Tier,Putting it All Together,Monitoring the Farm,Monitoring Solutions,Monitoring Caveats,Conference Roadmap,Current Version,Thank You';
  const file = await readFile(filesPath + '1testppt.pptx');

  const metadata = await OP.getData(file);
  expect(metadata.readOnly.slideTitles.value.join(',')).toBe(slideTitles);
});

it('Should throw error on invalid file', async () => {
  expect.assertions(1);

  const file = await readFile(filesPath + 'invaliddoc.docx');
  try {
    await OP.getData(file);
  } catch (err) {
    expect(err.message).toBe('Error: File not valid');
  }
});

it('Should work with xlsb', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Book1.xlsb');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with xlsm', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Book1.xlsm');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with xlsx', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Book1.xlsx');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with docm', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Doc1.docm');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with docx', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Doc1.docx');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with dotm', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Doc1.dotm');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with dotx', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Doc1.dotx');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with ppsm', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'pp.ppsm');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with ppsx', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'pp.ppsx');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with pptm', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'pp.pptm');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with potm', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'pp.potm');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with potx', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'pp.potx');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with xltm', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Book1.xltm');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with xltx', async () => {
  expect.assertions(2);

  const file = await readFile(filesPath + 'Book1.xltx');
  const metadata = await OP.getData(file);
  expect(Object.keys(metadata.editable).length).toBeGreaterThan(0);

  const blob = await OP.removeData(file);
  const meta = await OP.getData(blob);
  expect(Object.keys(meta.editable).length).toBe(0);
});

it('Should work with odt', async () => {
  expect.assertions(3);

  const file = await readFile(filesPath + 'Doc1.odt');
  const metadata = await OP.getData(file);
  expect(metadata.editable.creator.value).toBe('Torkel Helland Velure');

  metadata.editable.creator.value = 'Something else';
  const blob = await OP.editData(file, metadata);
  const meta = await OP.getData(blob);
  expect(meta.editable.creator.value).toBe('Something else');

  const blob2 = await OP.removeData(blob);
  const meta2 = await OP.getData(blob2);
  expect(Object.keys(meta2.editable).length).toBe(0);
});

it('Should work with odp', async () => {
  expect.assertions(3);

  const file = await readFile(filesPath + 'pp.odp');
  const metadata = await OP.getData(file);
  expect(metadata.editable.creator.value).toBe('Torkel Velure');

  metadata.editable.creator.value = 'Something else';
  const blob = await OP.editData(file, metadata);
  const meta = await OP.getData(blob);
  expect(meta.editable.creator.value).toBe('Something else');

  const blob2 = await OP.removeData(blob);
  const meta2 = await OP.getData(blob2);
  expect(Object.keys(meta2.editable).length).toBe(0);
});

it('Should work with ods', async () => {
  expect.assertions(3);

  const file = await readFile(filesPath + 'Book1.ods');
  const metadata = await OP.getData(file);
  expect(metadata.editable.creator.value).toBe('Torkel Velure');

  metadata.editable.creator.value = 'Something else';
  const blob = await OP.editData(file, metadata);
  const meta = await OP.getData(blob);
  expect(meta.editable.creator.value).toBe('Something else');

  const blob2 = await OP.removeData(blob);
  const meta2 = await OP.getData(blob2);
  expect(Object.keys(meta2.editable).length).toBe(0);
});
