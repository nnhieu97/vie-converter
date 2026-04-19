const fs = require('fs');
const https = require('https');
const path = require('path');
const devCerts = require('office-addin-dev-certs');
const { buildProject } = require('./build');

const rootDir = path.resolve(__dirname, '..');
const srcDir = path.join(rootDir, 'src');
const assetsDir = path.join(rootDir, 'assets');
const distDir = path.join(rootDir, 'dist');
const port = Number(process.env.PORT || 3000);

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
};

let buildInProgress = false;
let pendingBuild = false;

async function safeBuild() {
  if (buildInProgress) {
    pendingBuild = true;
    return;
  }

  buildInProgress = true;
  try {
    await buildProject({ production: false });
    console.log('[build] Refreshed dist/');
  } catch (error) {
    console.error('[build] Error:', error);
  } finally {
    buildInProgress = false;
  }

  if (pendingBuild) {
    pendingBuild = false;
    await safeBuild();
  }
}

function toSafeRelativePath(urlPath) {
  const normalized = path.normalize(decodeURIComponent(urlPath)).replace(/^([.][.][/\\])+/, '');
  if (normalized.includes('..')) {
    return null;
  }
  return normalized.replace(/^[/\\]+/, '');
}

function createStaticHandler() {
  return (req, res) => {
    const rawPath = req.url ? req.url.split('?')[0] : '/';
    const cleanPath = rawPath === '/' ? '/taskpane.html' : rawPath;
    const relativePath = toSafeRelativePath(cleanPath);

    if (!relativePath) {
      res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Bad request');
      return;
    }

    const filePath = path.join(distDir, relativePath);
    if (!filePath.startsWith(distDir)) {
      res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Forbidden');
      return;
    }

    fs.readFile(filePath, (error, content) => {
      if (error) {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Not found');
        return;
      }

      const ext = path.extname(filePath).toLowerCase();
      const contentType = MIME_TYPES[ext] || 'application/octet-stream';
      res.writeHead(200, { 'Content-Type': contentType, 'Cache-Control': 'no-store' });
      res.end(content);
    });
  };
}

function registerWatchers() {
  const watchTargets = [srcDir, assetsDir];
  for (const target of watchTargets) {
    if (!fs.existsSync(target)) {
      continue;
    }

    fs.watch(target, { recursive: true }, async (_eventType, filename) => {
      if (!filename) return;
      const ext = path.extname(filename).toLowerCase();
      if (!['.js', '.ts', '.html', '.css', '.png', '.svg'].includes(ext)) {
        return;
      }
      await safeBuild();
    });
  }
}

async function start() {
  await safeBuild();

  const httpsOptions = await devCerts.getHttpsServerOptions();
  const server = https.createServer(httpsOptions, createStaticHandler());

  registerWatchers();

  server.listen(port, 'localhost', () => {
    console.log(`[server] HTTPS serving at https://localhost:${port}`);
    console.log('[server] Sideload manifest.xml and open task pane in Word.');
  });
}

start().catch((error) => {
  console.error('[server] Failed to start:', error);
  process.exit(1);
});
