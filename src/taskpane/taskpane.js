import {
  buildUnitChangePlans,
  createFormatSnapshotFromRaw,
  mergeUnitChangePlans,
  normalizeTextWithoutSplitChars,
  sanitizeUnitText,
} from '../lib/vietnamese-conversion';

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

function canInsertWordComments() {
  try {
    return (
      typeof Office !== 'undefined' &&
      Office.context &&
      Office.context.requirements &&
      Office.context.requirements.isSetSupported('WordApi', '1.4')
    );
  } catch (_error) {
    return false;
  }
}

function toFormatSnapshot(font) {
  const raw = font && typeof font.toJSON === 'function' ? font.toJSON() : {};
  return createFormatSnapshotFromRaw(raw);
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

function buildSelectionResult(selection, sourceMode) {
  const unitPlans = buildUnitChangePlans(selection.units, { sourceMode });
  const mergedPlan = mergeUnitChangePlans(unitPlans, selection.text);
  return { unitPlans, mergedPlan };
}

async function getSelectionData(context) {
  const range = context.document.getSelection();
  range.load('text,isEmpty');
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
      text: sanitizeUnitText(item.text || ''),
      format: toFormatSnapshot(item.font),
    }));
    units = units.filter((item) => item.text.length > 0);

    const selectedNormalized = normalizeTextWithoutSplitChars(sanitizeUnitText(range.text || ''));
    const unitsNormalized = normalizeTextWithoutSplitChars(units.map((item) => item.text || '').join(''));

    if (!selectedNormalized || selectedNormalized !== unitsNormalized) {
      range.font.load();
      await context.sync();
      units = [
        {
          id: 'unit-1',
          range,
          text: sanitizeUnitText(range.text || ''),
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
        text: sanitizeUnitText(range.text || ''),
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

      const { mergedPlan } = buildSelectionResult(selection, getSourceMode());
      state.lastPreview = mergedPlan;
      renderChangePlan(mergedPlan);

      if (mergedPlan.changed) {
        setStatus(`Preview sẵn sàng. Convert: ${mergedPlan.summary.convertedCount}, skip: ${mergedPlan.summary.skippedCount}.`, 'ok');
      } else if (mergedPlan.summary.skippedCount > 0) {
        setStatus(`Không convert do mixed format. Skip ${mergedPlan.summary.skippedCount} block.`, 'warn');
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

      const { unitPlans, mergedPlan } = buildSelectionResult(selection, getSourceMode());

      const setTimes = shouldSetTimesNewRoman();
      const supportComments = canInsertWordComments();
      let convertApplied = 0;
      let fontApplied = 0;
      let commentAdded = 0;
      let runtimeErrors = 0;

      for (const entry of unitPlans) {
        if (entry.action === 'skip') {
          if (supportComments && entry.comment) {
            try {
              entry.range.insertComment(entry.comment);
              await context.sync();
              commentAdded += 1;
            } catch (_error) {
              runtimeErrors += 1;
            }
          }
          continue;
        }

        try {
          if (entry.action === 'convert') {
            const replaced = entry.range.insertText(entry.afterText, Word.InsertLocation.replace);
            if (setTimes) {
              replaced.font.name = 'Times New Roman';
              fontApplied += 1;
            }
            await context.sync();
            convertApplied += 1;
            continue;
          }

          if (setTimes) {
            entry.range.font.name = 'Times New Roman';
            await context.sync();
            fontApplied += 1;
          }
        } catch (_error) {
          runtimeErrors += 1;
        }
      }

      state.lastPreview = mergedPlan;
      renderChangePlan(mergedPlan);

      const summary = mergedPlan.summary;
      let message = `Đã xử lý theo từng block: convert ${convertApplied}/${summary.convertedCount}, skip ${summary.skippedCount}.`;
      if (setTimes) {
        message += ` Đặt Times New Roman cho ${fontApplied} block không bị skip.`;
      }
      if (supportComments && commentAdded > 0) {
        message += ` Đã thêm ${commentAdded} comment cho block skip.`;
      } else if (!supportComments && summary.skippedCount > 0) {
        message += ' Host hiện tại không hỗ trợ thêm comment tự động (cần WordApi 1.4+).';
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
