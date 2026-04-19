const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const distDir = path.join(rootDir, 'dist');

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function copyFileRelative(fromRelative, toRelative) {
  const fromPath = path.join(rootDir, fromRelative);
  const toPath = path.join(rootDir, toRelative);
  ensureDir(path.dirname(toPath));
  fs.copyFileSync(fromPath, toPath);
}

async function buildProject({ production = false } = {}) {
  ensureDir(distDir);

  await esbuild.build({
    entryPoints: [path.join(rootDir, 'src/taskpane/taskpane.js')],
    outfile: path.join(distDir, 'taskpane.js'),
    bundle: true,
    format: 'iife',
    platform: 'browser',
    target: ['es2020'],
    minify: production,
    sourcemap: production ? false : 'inline',
    define: {
      'process.env.NODE_ENV': JSON.stringify(production ? 'production' : 'development'),
    },
  });

  copyFileRelative('src/taskpane/taskpane.html', 'dist/taskpane.html');
  copyFileRelative('src/taskpane/taskpane.css', 'dist/taskpane.css');
  copyFileRelative('assets/icon-16.png', 'dist/assets/icon-16.png');
  copyFileRelative('assets/icon-32.png', 'dist/assets/icon-32.png');
  copyFileRelative('assets/icon-80.png', 'dist/assets/icon-80.png');
}

if (require.main === module) {
  const production = process.argv.includes('--production');
  buildProject({ production })
    .then(() => {
      console.log(`[build] Done (${production ? 'production' : 'development'})`);
    })
    .catch((error) => {
      console.error('[build] Failed:', error);
      process.exit(1);
    });
}

module.exports = { buildProject };
