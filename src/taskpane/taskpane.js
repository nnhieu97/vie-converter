import { buildConversionPlan } from '../lib/vietnamese-conversion';

const sourceModeEl = document.getElementById('source-mode');
const scopeModeEl = document.getElementById('scope-mode');
const previewBtn = document.getElementById('preview-btn');
const applyBtn = document.getElementById('apply-btn');
const statusEl = document.getElementById('status');
const metaEl = document.getElementById('meta');
const beforeEl = document.getElementById('before-text');
const afterEl = document.getElementById('after-text');

const state = {
  ready: false,
  lastPreview: null,
};

function setStatus(message, level = 'info') {
  statusEl.textContent = message;
  statusEl.classList.remove('ok', 'warn', 'error');
  if (level === 'ok' || level === 'warn' || level === 'error') {
    statusEl.classList.add(level);
  }
}

function setApplyEnabled(enabled) {
  applyBtn.disabled = !enabled;
}

function getSourceMode() {
  return sourceModeEl.value;
}

function getScopeMode() {
  return scopeModeEl.value;
}

function buildPlan(text, sourceMode, fontName = '') {
  return buildConversionPlan(text, { sourceMode, fontName });
}

function countCharDiff(before, after) {
  const maxLen = Math.max(before.length, after.length);
  let diff = Math.abs(before.length - after.length);
  const minLen = Math.min(before.length, after.length);
  for (let i = 0; i < minLen; i += 1) {
    if (before[i] !== after[i]) diff += 1;
  }
  return Math.min(diff, maxLen);
}

function renderPreview(plan) {
  beforeEl.textContent = plan.original || '';
  afterEl.textContent = plan.converted || '';

  const hintPart = plan.hint ? ` | ${plan.hint}` : '';
  const fontPart = plan.fontName ? ` | font: ${plan.fontName}` : '';

  if (!plan.changed) {
    const detectedText = plan.detected ? ` | detect: ${plan.detected}` : '';
    metaEl.textContent = `Không có thay đổi${detectedText}${fontPart}${plan.reason ? ` | ${plan.reason}` : ''}${hintPart}`;
    setApplyEnabled(false);
    return;
  }

  const diffCount = countCharDiff(plan.original, plan.converted);
  const detectedPart = plan.detected ? `detect: ${plan.detected}` : 'detect: manual';
  metaEl.textContent = `${detectedPart} | source: ${plan.effectiveSource}${fontPart} | thay đổi ký tự: ~${diffCount}${hintPart}`;
  setApplyEnabled(true);
}

async function getSelection(context) {
  const range = context.document.getSelection();
  range.load('text,font/name');
  await context.sync();
  return {
    range,
    text: range.text || '',
    fontName: (range.font && range.font.name) || '',
  };
}

async function previewSelection() {
  if (!state.ready) {
    setStatus('Word API chưa sẵn sàng.', 'warn');
    return;
  }

  if (getScopeMode() !== 'selection') {
    setStatus('Main body đang là hook cho phase tiếp theo. Hiện chỉ hỗ trợ vùng chọn.', 'warn');
    return;
  }

  setStatus('Đang tạo preview...', 'info');
  setApplyEnabled(false);

  try {
    await Word.run(async (context) => {
      const { text, fontName } = await getSelection(context);
      const plan = buildPlan(text, getSourceMode(), fontName);
      state.lastPreview = plan;
      renderPreview(plan);

      if (plan.changed) {
        setStatus('Preview sẵn sàng. Kiểm tra lại trước khi Apply.', 'ok');
      } else {
        setStatus(plan.reason || 'Không có thay đổi.', 'warn');
      }
    });
  } catch (error) {
    setStatus(`Không tạo được preview: ${error?.message || String(error)}`, 'error');
  }
}

async function applySelection() {
  if (!state.ready) {
    setStatus('Word API chưa sẵn sàng.', 'warn');
    return;
  }

  if (getScopeMode() !== 'selection') {
    setStatus('Main body đang là hook cho phase tiếp theo. Hiện chỉ hỗ trợ vùng chọn.', 'warn');
    return;
  }

  setStatus('Đang áp dụng chuyển mã...', 'info');
  setApplyEnabled(false);

  try {
    await Word.run(async (context) => {
      const { range, text, fontName } = await getSelection(context);
      const plan = buildPlan(text, getSourceMode(), fontName);

      if (!plan.changed) {
        renderPreview(plan);
        setStatus(plan.reason || 'Không có thay đổi để áp dụng.', 'warn');
        return;
      }

      range.insertText(plan.converted, Word.InsertLocation.replace);
      await context.sync();

      state.lastPreview = plan;
      renderPreview(plan);
      setStatus(`Đã chuyển mã thành công từ ${plan.effectiveSource} sang Unicode.`, 'ok');
    });
  } catch (error) {
    setStatus(`Không áp dụng được chuyển mã: ${error?.message || String(error)}`, 'error');
  }
}

function wireEvents() {
  previewBtn.addEventListener('click', previewSelection);
  applyBtn.addEventListener('click', applySelection);

  sourceModeEl.addEventListener('change', () => {
    setApplyEnabled(false);
    setStatus('Đã thay đổi cấu hình nguồn bảng mã. Hãy Preview lại.', 'warn');
  });

  scopeModeEl.addEventListener('change', () => {
    setApplyEnabled(false);
    if (scopeModeEl.value === 'selection') {
      setStatus('Đang ở chế độ vùng chọn.', 'info');
    } else {
      setStatus('Main body đang tắt trong MVP.', 'warn');
    }
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus('Add-in này chỉ chạy trong Microsoft Word.', 'error');
    return;
  }

  state.ready = true;
  wireEvents();
  setStatus('Sẵn sàng. Hãy chọn đoạn văn bản rồi bấm Preview.', 'ok');
});