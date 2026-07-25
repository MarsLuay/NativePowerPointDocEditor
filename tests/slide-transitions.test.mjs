import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { test } from 'node:test';

import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const source = await bundleSource('src/powerpoint/slideTransitions.ts', 'slide-transitions.cjs');
const { readSlideTransition, writeSlideTransition } = require(source);

const slide = (body) => [
  '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  ' xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">',
  '<p:cSld><p:spTree/></p:cSld>',
  body,
  '</p:sld>',
].join('');

test('reads PowerPoint transition duration, advance options, and effect directions', () => {
  const xml = slide([
    '<p:transition spd="med" p14:dur="750" advClick="0" advTm="2000">',
    '<p:wipe dir="l"/>',
    '</p:transition>',
  ].join(''));

  assert.deepEqual(readSlideTransition(xml), {
    kind: 'wipe',
    direction: 'left',
    durationMs: 750,
    advanceAfterMs: 2000,
    advanceOnClick: false,
    speed: 'medium',
  });
});

test('writes a PowerPoint p14 duration and preserves timing and slide extension markup', () => {
  const xml = slide([
    '<p:transition spd="fast"><p:fade/><p:sndAc><p:stSnd loop="1"/></p:sndAc><p:extLst><p:ext uri="keep"/></p:extLst></p:transition>',
    '<p:timing><p:tnLst><p:par/></p:tnLst></p:timing>',
    '<p:extLst><p:ext uri="slide-keep"/></p:extLst>',
  ].join(''));
  const updated = writeSlideTransition(xml, {
    kind: 'push',
    direction: 'down',
    durationMs: 1250,
    advanceAfterMs: 5000,
    advanceOnClick: false,
  });

  assert.match(updated, /<p:transition\b[^>]*\bspd="slow"[^>]*\badvClick="0"[^>]*\badvTm="5000"[^>]*\bp14:dur="1250"/);
  assert.match(updated, /<p:push dir="d"\/>/);
  assert.match(updated, /<p:timing><p:tnLst><p:par\/><\/p:tnLst><\/p:timing>/);
  assert.match(updated, /<p:extLst><p:ext uri="slide-keep"\/><\/p:extLst>/);
  assert.match(updated, /<p:sndAc><p:stSnd loop="1"\/><\/p:sndAc>/);
  assert.match(updated, /<p:transition[\s\S]*?<p:extLst><p:ext uri="keep"\/><\/p:extLst><\/p:transition>/);
  assert.ok(updated.indexOf('<p:transition') < updated.indexOf('<p:timing'));
  assert.deepEqual(readSlideTransition(updated), {
    kind: 'push',
    direction: 'down',
    durationMs: 1250,
    advanceAfterMs: 5000,
    advanceOnClick: false,
    speed: 'slow',
  });
});

test('writes PowerPoint split directions and preserves no-transition slide data', () => {
  const xml = slide('<p:timing><p:tnLst><p:par/></p:tnLst></p:timing><p:extLst><p:ext uri="keep"/></p:extLst>');
  const withSplit = writeSlideTransition(xml, {
    kind: 'split',
    splitDirection: 'in',
    splitOrientation: 'vertical',
    throughBlack: undefined,
    advanceOnClick: true,
  });

  assert.match(withSplit, /<p:split dir="in" orient="vert"\/>/);
  assert.deepEqual(readSlideTransition(withSplit), {
    kind: 'split',
    splitDirection: 'in',
    splitOrientation: 'vertical',
    durationMs: null,
    advanceAfterMs: null,
    advanceOnClick: true,
    speed: 'fast',
  });

  const withoutTransition = writeSlideTransition(withSplit, { kind: 'none' });
  assert.deepEqual(readSlideTransition(withoutTransition), { kind: 'none' });
  assert.match(withoutTransition, /<p:timing><p:tnLst><p:par\/><\/p:tnLst><\/p:timing>/);
  assert.match(withoutTransition, /<p:extLst><p:ext uri="keep"\/><\/p:extLst>/);
});

test('replaces AlternateContent transition branches without leaving a conflicting fallback', () => {
  const xml = slide([
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">',
    '<mc:Choice Requires="p14"><p:transition><p:fade/></p:transition></mc:Choice>',
    '<mc:Fallback><p:transition><p:cut/></p:transition></mc:Fallback>',
    '</mc:AlternateContent>',
    '<p:timing/>',
  ].join(''));
  const updated = writeSlideTransition(xml, { kind: 'fade', throughBlack: true });

  assert.equal((updated.match(/<p:transition\b/g) ?? []).length, 1);
  assert.doesNotMatch(updated, /AlternateContent/);
  assert.match(updated, /<p:fade thruBlk="1"\/>/);
  assert.match(updated, /<p:timing\/>/);
});

test('rejects invalid PowerPoint transition options', () => {
  const xml = slide('');
  assert.throws(
    () => writeSlideTransition(xml, { kind: 'wipe' }),
    /wipe transitions require a direction/,
  );
  assert.throws(
    () => writeSlideTransition(xml, { kind: 'fade', direction: 'left' }),
    /fade transitions do not support a side direction/,
  );
  assert.throws(
    () => writeSlideTransition(xml, { kind: 'cut', durationMs: 1.5 }),
    /durationMs must be an integer/,
  );
  assert.throws(
    () => writeSlideTransition(xml, { kind: 'fade', advanceAfterMs: 2_147_483_648 }),
    /advanceAfterMs must be an integer/,
  );
});
