import { buildChangePlanFromAtoms } from '../lib/vietnamese-conversion';

const sourceModeEl = document.getElementById('source-mode');
const scopeModeEl = document.getElementById('scope-mode');
const setTimesFontEl = document.getElementById('set-times-font');
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

function setApplyEnabled(_enabled) {
  applyBtn.disabled = false;
}

function getSourceMode() {
  return sourceModeEl.value;
}

function getScopeMode() {
  return scopeModeEl.value;
}

function shouldSetTimesNewRoman() {
  return Boolean(setTimesFontEl && setTimesFontEl.checked);
}

function pickFontValue(raw, key) {
  if (!raw || typeof raw !== 'object') return undefined;
  if (!Object.prototype.hasOwnProperty.call(raw, key)) return undefined;
  return raw[key];
}

function toFormatSnapshot(font) {
  const raw = font && typeof font.toJSON === 'function' ? font.toJSON() : {};
  return {
    fontName: pickFontValue(raw, 'name'),
    fontSize: pickFontValue(raw, 'size'),
    bold: pickFontValue(raw, 'bold'),
    italic: pickFontValue(raw, 'italic'),
    underline: pickFontValue(raw, 'underline'),
    fontColor: pickFontValue(raw, 'color'),
    highlightColor: pickFontValue(raw, 'highlightColor'),
    strikeThrough: pickFontValue(raw, 'strikeThrough'),
    doubleStrikeThrough: pickFontValue(raw, 'doubleStrikeThrough'),
    superscript: pickFontValue(raw, 'superscript'),
    subscript: pickFontValue(raw, 'subscript'),
    allCaps: pickFontValue(raw, 'allCaps'),
    smallCaps: pickFontValue(raw, 'smallCaps'),
    hidden: pickFontValue(raw, 'hidden'),
    spacing: pickFontValue(raw, 'spacing'),
    kerning: pickFontValue(raw, 'kerning'),
    scale: pickFontValue(raw, 'scaling'),
    position: pickFontValue(raw, 'position'),
  };
}

function normalizeWithoutSplitChars(text) {
  return String(text || '').replace(/[\r\t\v\n\f]/g, '');
}

function buildUnitPlans(units, sourceMode) {
  return units.map((unit, index) => {
    const atom = {
      id: `atom-${index + 1}`,
      text: unit.text || '',
      format: unit.format || null,
    };
    const plan = buildChangePlanFromAtoms([atom], { sourceMode });
    const item = (plan.items && plan.items[0]) || null;

    return {
      unit,
      plan,
      action: item ? item.action : 'noop',
      comment: item ? item.comment : null,
      beforeText: plan.inputText || unit.text || '',
      afterText: plan.outputText || unit.text || '',
      changed: Boolean(item && item.action === 'convert'),
    };
  });
}

function mergeUnitPlans(unitPlans, fallbackInputText = '') {
  const summary = unitPlans.reduce(
    (acc, entry) => {
      acc.totalBlocks += 1;
      if (entry.action === 'convert') acc.convertedCount += 1;
      else if (entry.action === 'skip') acc.skippedCount += 1;
      else acc.noopCount += 1;
      return acc;
    },
    { totalBlocks: 0, convertedCount: 0, skippedCount: 0, noopCount: 0 }
  );

  const inputText = fallbackInputText || unitPlans.map((entry) => entry.beforeText).join('\n');
  const outputText = unitPlans.map((entry) => entry.afterText).join('\n');
  const comments = unitPlans.filter((entry) => entry.action === 'skip' && entry.comment).map((entry) => entry.comment);

  return {
    changed: unitPlans.some((entry) => entry.action === 'convert'),
    inputText,
    outputText,
    summary,
    comments,
  };
}

function renderChangePlan(changePlan) {
  beforeEl.textContent = changePlan.inputText || '';
  afterEl.textContent = changePlan.outputText || '';

  const summary = changePlan.summary || { totalBlocks: 0, convertedCount: 0, skippedCount: 0, noopCount: 0 };
  const skipComment = changePlan.comments && changePlan.comments.length ? ` | ${changePlan.comments[0]}` : '';

  if (!changePlan.changed) {
    metaEl.textContent = `Không có thay đổi | block: ${summary.totalBlocks} | convert: ${summary.convertedCount} | skip: ${summary.skippedCount}${skipComment}`;
    setApplyEnabled(false);
    return;
  }

  metaEl.textContent = `Block: ${summary.totalBlocks} | convert: ${summary.convertedCount} | skip: ${summary.skippedCount} | giữ nguyên: ${summary.noopCount}${skipComment}`;
  setApplyEnabled(true);
}

async function getSelectionData(context) {
  const range = context.document.getSelection();
  range.load('text,font/name,isEmpty');
  await context.sync();

  if (range.isEmpty) {
    return {
      range,
      text: '',
      isEmpty: true,
      units: [],
    };
  }

  const textRanges = range.getTextRanges(['\r', '\t', '\v', '\n', '\f'], false);
  textRanges.load('items');
  await context.sync();

  let units = [];
  if (textRanges.items.length) {
    for (const item of textRanges.items) {
      item.load('text');
      item.font.load();
    }
    await context.sync();

    units = textRanges.items.map((item, index) => ({
      id: `unit-${index + 1}`,
      range: item,
      text: item.text || '',
      format: toFormatSnapshot(item.font),
    }));

    const selectedNormalized = normalizeWithoutSplitChars(range.text || '');
    const unitsNormalized = normalizeWithoutSplitChars(units.map((item) => item.text || '').join(''));
    if (!selectedNormalized || selectedNormalized !== unitsNormalized) {
      range.font.load();
      await context.sync();
      units = [
        {
          id: 'unit-1',
          range,
          text: range.text || '',
          format: toFormatSnapshot(range.font),
        },
      ];
    }
  } else {
    range.font.load();
    await context.sync();
    units = [
      {
        id: 'unit-1',
        range,
        text: range.text || '',
        format: toFormatSnapshot(range.font),
      },
    ];
  }

  return {
    range,
    text: range.text || '',
    isEmpty: false,
    units,
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
      const selection = await getSelectionData(context);
      if (selection.isEmpty) {
        state.lastPreview = null;
        beforeEl.textContent = '';
        afterEl.textContent = '';
        metaEl.textContent = 'Vùng chọn rỗng. Hãy bôi đen đoạn cần xử lý.';
        setStatus('Không có vùng chọn. Hãy bôi đen text trước khi Preview.', 'warn');
        return;
      }

      const unitPlans = buildUnitPlans(selection.units, getSourceMode());
      const previewPlan = mergeUnitPlans(unitPlans, selection.text);
      state.lastPreview = previewPlan;
      renderChangePlan(previewPlan);

      if (previewPlan.changed) {
        setStatus(
          `Preview sẵn sàng. Convert: ${previewPlan.summary.convertedCount}, skip: ${previewPlan.summary.skippedCount}.`,
          'ok'
        );
      } else if (previewPlan.summary.skippedCount > 0) {
        setStatus(`Không convert do mixed format. Skip ${previewPlan.summary.skippedCount} block.`, 'warn');
      } else {
        setStatus('Không có thay đổi.', 'warn');
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

  try {
    await Word.run(async (context) => {
      const selection = await getSelectionData(context);
      if (selection.isEmpty) {
        setStatus('Không có vùng chọn để Apply. Hãy bôi đen text trước.', 'warn');
        return;
      }

      const unitPlans = buildUnitPlans(selection.units, getSourceMode());
      const resultPlan = mergeUnitPlans(unitPlans, selection.text);

      const setTimes = shouldSetTimesNewRoman();
      let convertApplied = 0;
      let fontApplied = 0;
      let runtimeErrors = 0;

      for (const entry of unitPlans) {
        if (entry.action === 'skip') {
          continue;
        }

        try {
          if (entry.action === 'convert') {
            const replaced = entry.unit.range.insertText(entry.afterText, Word.InsertLocation.replace);
            if (setTimes) {
              replaced.font.name = 'Times New Roman';
              fontApplied += 1;
            }
            await context.sync();
            convertApplied += 1;
            continue;
          }

          if (setTimes) {
            entry.unit.range.font.name = 'Times New Roman';
            await context.sync();
            fontApplied += 1;
          }
        } catch (_error) {
          runtimeErrors += 1;
        }
      }

      state.lastPreview = resultPlan;
      renderChangePlan(resultPlan);

      const summary = resultPlan.summary;
      let message = `Đã xử lý theo từng block: convert ${convertApplied}/${summary.convertedCount}, skip ${summary.skippedCount}.`;
      if (setTimes) {
        message += ` Đặt Times New Roman cho ${fontApplied} block không bị skip.`;
      }
      if (runtimeErrors > 0) {
        message += ` Có ${runtimeErrors} block lỗi khi apply.`;
        setStatus(message, 'warn');
      } else {
        setStatus(message, 'ok');
      }
    });
  } catch (error) {
    setStatus(`Không áp dụng được chuyển mã: ${error?.message || String(error)}`, 'error');
  }
}

function wireEvents() {
  previewBtn.addEventListener('click', previewSelection);
  applyBtn.addEventListener('click', applySelection);

  sourceModeEl.addEventListener('change', () => {
    setStatus('Đã thay đổi cấu hình nguồn bảng mã. Hãy Preview lại.', 'warn');
  });

  scopeModeEl.addEventListener('change', () => {
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
  setApplyEnabled(true);
  setStatus('Sẵn sàng. Hãy chọn đoạn văn bản rồi bấm Preview.', 'ok');
});
