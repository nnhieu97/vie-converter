const { spawn } = require('child_process');
const fs = require('fs');
const net = require('net');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const manifestPath = path.join(rootDir, 'manifest.local.xml');
const serverScriptPath = path.join(rootDir, 'scripts', 'dev-server.js');
const port = Number(process.env.PORT || 3000);

let serverProcess = null;
let sideloadProcess = null;
let sideloadStarted = false;
let shuttingDown = false;

function sideloadCommand() {
  if (process.platform === 'win32') {
    return { cmd: 'npx', args: ['office-addin-debugging', 'start', manifestPath, 'desktop'], useShell: true };
  }
  return { cmd: 'npx', args: ['office-addin-debugging', 'start', manifestPath, 'desktop'], useShell: false };
}

function canConnectToPortAtHost(portNumber, host) {
  return new Promise((resolve) => {
    const socket = new net.Socket();

    socket.setTimeout(800);
    socket.once('connect', () => {
      socket.destroy();
      resolve(true);
    });
    socket.once('timeout', () => {
      socket.destroy();
      resolve(false);
    });
    socket.once('error', () => {
      resolve(false);
    });

    socket.connect(portNumber, host);
  });
}

async function canConnectToPort(portNumber) {
  const hosts = ['127.0.0.1', '::1', 'localhost'];
  for (const host of hosts) {
    // eslint-disable-next-line no-await-in-loop
    if (await canConnectToPortAtHost(portNumber, host)) {
      return true;
    }
  }
  return false;
}

function startSideload() {
  if (sideloadStarted || shuttingDown) {
    return;
  }
  sideloadStarted = true;

  const { cmd, args, useShell } = sideloadCommand();
  console.log(`[desktop] Sideloading ${path.basename(manifestPath)} and opening Word...`);

  sideloadProcess = spawn(cmd, args, {
    cwd: rootDir,
    stdio: 'inherit',
    shell: useShell,
  });

  sideloadProcess.on('exit', (code) => {
    if (shuttingDown) {
      return;
    }

    if (code !== 0) {
      console.error(`[desktop] Sideload failed with exit code ${code}.`);
      shutdown(code || 1);
      return;
    }

    console.log('[desktop] Word launch command completed.');
  });

  sideloadProcess.on('error', (error) => {
    console.error('[desktop] Failed to start sideload command:', error);
    shutdown(1);
  });
}

function shutdown(exitCode = 0) {
  if (shuttingDown) {
    return;
  }
  shuttingDown = true;

  const finalize = () => process.exit(exitCode);

  if (!serverProcess || serverProcess.killed) {
    finalize();
    return;
  }

  serverProcess.once('exit', finalize);
  serverProcess.kill('SIGINT');

  setTimeout(() => {
    if (!serverProcess.killed) {
      serverProcess.kill('SIGTERM');
    }
  }, 1500);
}

function startDevServer() {
  serverProcess = spawn(process.execPath, [serverScriptPath], {
    cwd: rootDir,
    stdio: ['inherit', 'pipe', 'pipe'],
    shell: false,
  });

  serverProcess.stdout.on('data', (chunk) => {
    const text = chunk.toString();
    process.stdout.write(text);

    if (text.includes('[server] HTTPS serving at')) {
      startSideload();
    }
  });

  serverProcess.stderr.on('data', (chunk) => {
    const text = chunk.toString();
    process.stderr.write(text);
  });

  serverProcess.on('exit', (code) => {
    if (!shuttingDown) {
      console.error(`[desktop] Dev server exited with code ${code}.`);
      shutdown(code || 1);
    }
  });

  serverProcess.on('error', (error) => {
    console.error('[desktop] Failed to start dev server:', error);
    shutdown(1);
  });
}

async function main() {
  if (!fs.existsSync(manifestPath)) {
    throw new Error(`Manifest not found: ${manifestPath}`);
  }

  const inUse = await canConnectToPort(port);

  if (inUse) {
    console.log(`[desktop] Port ${port} already has a running server. Reusing it.`);
    startSideload();
    return;
  }

  startDevServer();
}

process.on('SIGINT', () => shutdown(0));
process.on('SIGTERM', () => shutdown(0));

main().catch((error) => {
  console.error('[desktop] Start script failed:', error);
  shutdown(1);
});
