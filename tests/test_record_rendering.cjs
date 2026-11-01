// Offline only: NODE_PATH=<jsdom installation>/node_modules node --test tests/test_record_rendering.cjs
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');

const source = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
// Execute only the renderer, never the app initializer or its production fetch calls.
const renderer = source.slice(source.indexOf('    function renderRecords('), source.indexOf('    function openEditModal('));
function fixture() {
  const dom = new JSDOM('<div id="records"></div>');
  const edits = [], deletes = [];
  const render = new Function('document', 'openEditModal', 'openDeleteModal', renderer + '; return renderRecords;')(
    dom.window.document, (...args) => edits.push(args), id => deletes.push(id));
  return { document: dom.window.document, render, edits, deletes };
}
const record = { id: 7, record_date: '2026-11-01', location: '한글 숙소 · 2층', blanket: 2, mat: 0,
  pillow_cover: 3, towel: 4, body_towel: 1, pillow_fill: 0, cotton_blanket: 2 };

test('normal Unicode/date/counts and edit/delete arguments are preserved', () => {
  const f = fixture(); f.render([record]);
  assert.equal(f.document.querySelector('.record-location-text').textContent, record.location);
  assert.equal(f.document.querySelector('.record-date').textContent, record.record_date);
  assert.equal(f.document.querySelectorAll('.record-item').length, 5);
  f.document.querySelector('.edit').click(); f.document.querySelector('.delete').click();
  assert.deepEqual(f.edits, [[7, '2026-11-01', record.location, 2, 0, 3, 4, 1, 0, 2]]);
  assert.deepEqual(f.deletes, [7]);
});
test('HTML and JavaScript payloads remain inert text and literal event arguments', () => {
  const f = fixture();
  const payload = `'><img src=x onerror="alert(1)"><script>alert(2)</script>`;
  const r = {...record, id: '1);alert(3)//', location: payload, record_date: payload};
  f.render([r]);
  assert.equal(f.document.querySelector('.record-location-text').textContent, payload);
  assert.equal(f.document.querySelector('.record-date').textContent, payload);
  assert.equal(f.document.querySelectorAll('img, script, [onclick], [onerror]').length, 0);
  f.document.querySelector('.edit').click(); f.document.querySelector('.delete').click();
  assert.equal(f.edits[0][2], payload);
  assert.equal(f.deletes[0], r.id);
});
test('rerender clears old handlers and empty state remains functional', () => {
  const f = fixture(); f.render([record, {...record, id: 8}]); f.render([record]);
  assert.equal(f.document.querySelectorAll('.record-card').length, 1);
  f.document.querySelector('.edit').click(); assert.equal(f.edits.length, 1);
  f.render([]);
  assert.equal(f.document.querySelectorAll('.record-card').length, 0);
  assert.match(f.document.querySelector('.empty-title').textContent, /기록이 없습니다/);
});
