import fs from 'node:fs';

const files = ['js/app.js', 'index.html', 'js/core.js'];
let src = '';
files.forEach((f) => { src += fs.readFileSync(f, 'utf8') + '\n'; });
const set = new Set();
const re = /class=[\\"']?([^"'\\>]+)[\\"']/g;
let m;
while ((m = re.exec(src))) {
  String(m[1] || '').split(/[\s]+/).forEach((c) => {
    const name = c.trim();
    if (!name) return;
    if (!/^[a-zA-Z][a-zA-Z0-9_-]*$/.test(name)) return;
    set.add(name);
  });
}
const all = [...set].sort();
console.log(all.length + ' classes');
console.log(all.join(' '));
