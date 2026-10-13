// Renders index.html frame-by-frame in headless Chromium, then muxes with the soundtrack.
// Usage: node render.mjs [fps=30] [width=1080] [height=1920]
import { chromium } from 'playwright';
import { createServer } from 'node:http';
import { readFile, mkdir, writeFile, rm } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const dir = path.dirname(fileURLToPath(import.meta.url));
const [fps = 30, W = 1080, H = 1920] = process.argv.slice(2).map(Number);
const DURATION = 10, frames = Math.round(DURATION * fps);
const types = { '.html': 'text/html', '.js': 'text/javascript', '.wav': 'audio/wav' };

const server = createServer(async (req, res) => {
  try {
    const p = path.join(dir, decodeURIComponent(new URL(req.url, 'http://x').pathname));
    const body = await readFile(p);
    res.writeHead(200, { 'content-type': types[path.extname(p)] || 'application/octet-stream' });
    res.end(body);
  } catch { res.writeHead(404); res.end(); }
}).listen(0);
const port = server.address().port;

const browser = await chromium.launch({ args: ['--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
const page = await browser.newPage({ viewport: { width: W, height: H } });
page.on('console', m => console.log('[page]', m.text()));
await page.goto(`http://localhost:${port}/index.html?render&w=${W}&h=${H}`);
await page.waitForFunction('window.ready === true');

const out = path.join(dir, 'frames');
await rm(out, { recursive: true, force: true });
await mkdir(out);
for (let i = 0; i < frames; i++) {
  const url = await page.evaluate(t => window.renderAt(t), i / fps);
  await writeFile(path.join(out, `f${String(i).padStart(4, '0')}.png`), Buffer.from(url.split(',')[1], 'base64'));
  if (i % 30 === 0) console.log(`frame ${i}/${frames}`);
}
await browser.close();
server.close();

execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-framerate', String(fps), '-i', path.join(out, 'f%04d.png'),
  '-i', path.join(dir, 'soundtrack.wav'), '-c:v', 'libx264', '-preset', 'slow', '-crf', '18', '-pix_fmt', 'yuv420p',
  '-c:a', 'aac', '-b:a', '192k', '-t', String(DURATION), '-movflags', '+faststart',
  path.join(dir, 'baby_dino_surprise_egg.mp4')], { stdio: 'inherit' });
console.log('wrote baby_dino_surprise_egg.mp4');
