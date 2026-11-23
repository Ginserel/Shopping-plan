#!/usr/bin/env node
/*
 * Builds two single-file versions of the app from src/:
 *   dist/shopping-plan.html  – everything inlined (works offline, e.g. from a shared drive)
 *   dist/artifact.html       – body-only page that loads SheetJS from cdnjs (for hosted previews)
 */
'use strict';
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const src = (f) => fs.readFileSync(path.join(root, 'src', f), 'utf8');
const html = src('index.html');
const css = src('style.css');
const js = (f) => fs.readFileSync(path.join(root, 'src', f), 'utf8').replace(/<\/script/gi, '<\\/script');
const vendor = fs.readFileSync(path.join(root, 'vendor', 'xlsx.full.min.js'), 'utf8').replace(/<\/script/gi, '<\\/script');
const CDN = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

function inline(page, { cdn }) {
  return page
    .replace('<link rel="stylesheet" href="style.css">', () => `<style>\n${css}\n</style>`)
    .replace('<script src="../vendor/xlsx.full.min.js"></script>', () => (cdn ? `<script src="${CDN}"></script>` : `<script>\n${vendor}\n</script>`))
    .replace(/<script src="([a-z]+\.js)"><\/script>/g, (_, f) => `<script>\n${js(f)}\n</script>`);
}

fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
fs.writeFileSync(path.join(root, 'dist', 'shopping-plan.html'), inline(html, { cdn: false }));

// Hosted page: the host supplies doctype/html/head/body.
const head = /<head>([\s\S]*?)<\/head>/.exec(html)[1].replace(/<meta charset[^>]*>\s*/, '').replace(/<meta name="viewport"[^>]*>\s*/, '');
const body = /<body>([\s\S]*?)<\/body>/.exec(html)[1];
fs.writeFileSync(path.join(root, 'dist', 'artifact.html'), inline(head.trim() + '\n' + body.trim() + '\n', { cdn: true }));
for (const f of ['shopping-plan.html', 'artifact.html']) {
  console.log(`dist/${f}  ${(fs.statSync(path.join(root, 'dist', f)).size / 1024).toFixed(0)} KB`);
}
