// Optional UI smoke test. Set JSDOM_MODULE to a path containing jsdom,
// or install jsdom in your usual development environment.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require(process.env.JSDOM_MODULE || 'jsdom');
const html = fs.readFileSync(path.join(__dirname, '../deliverables/index.html'), 'utf8');
const dom = new JSDOM(html, { runScripts: 'dangerously' });
const document = dom.window.document;
const q = document.getElementById('q');
const category = document.getElementById('category');
const visible = () => [...document.querySelectorAll('article')].filter(c => !c.hidden);
function search(value) {
  q.value = value;
  q.dispatchEvent(new dom.window.Event('input'));
}
assert.equal(document.getElementById('filters').hidden, false);
assert.equal(visible().length, 12);
search('نجار');
assert.equal(visible().length, 1);
assert.match(visible()[0].textContent, /Najjar/);
search('716496049');
assert.equal(visible().length, 0);
search('71496049');
assert.equal(visible().length, 1);
search('إلدورادو');
assert.equal(visible().length, 1);
search('الدورادو');
assert.equal(visible().length, 1);
search('');
category.value = 'صيدليات';
category.dispatchEvent(new dom.window.Event('change'));
assert.equal(visible().length, 3);
search('اللبان');
assert.equal(visible().length, 1);
search('no matching store');
assert.equal(visible().length, 0);
assert.equal(document.getElementById('count').textContent, 'عرض 0 من 12 إدراجاً');
assert.ok([...document.querySelectorAll('.category')].every(g => g.hidden));
search('');
category.value = '';
category.dispatchEvent(new dom.window.Event('change'));
assert.equal(visible().length, 12);
assert.equal(document.querySelectorAll('.lead').length, 4);
assert.equal(document.querySelectorAll('.sources li').length, 27);
assert.equal(document.querySelectorAll('script[src]').length, 0);
assert.equal(document.documentElement.dir, 'rtl');
dom.window.close();
console.log('PASS: all 12 records, 4 lead badges, 27 sources, Arabic/phone search, category and combined filters, empty state, and offline layout.');
