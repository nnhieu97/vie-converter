import {
  buildUnitChangePlans,
  createFormatSnapshotFromRaw,
  FORMAT_SNAPSHOT_MAP,
  mergeUnitChangePlans,
  normalizeTextWithoutSplitChars,
  sanitizeUnitText,
} from '../lib/vietnamese-conversion';

const sourceModeEl = document.getElementById('source-mode');
const scopeModeEl = document.getElementById('scope-mode');
const setTimesFontEl = document.getElementById('set-times-font');
const allowMixedFormatEl = document.getElementById('allow-mixed-format');
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

const TEXT_RANGE_DELIMITERS = Object.freeze(['\r', '\t', '\v', '\n', '\f']);
const FONT_LOAD_PROPERTIES = Object.freeze(Array.from(new Set(Object.values(FORMAT_SNAPSHOT_MAP))));

function setStatus(message, level = 'info') {
  statusEl.textContent = message;
  statusEl.classList.remove('ok', 'warn', 'error');
  if (level === 'ok' || level === 'warn' || level === 'error') {
    statusEl.classList.add(level);
  }
}

function setControlsBusy(isBusy) {
  const disabled = Boolean(isBusy);
  previewBtn.disabled = disabled;
  applyBtn.disabled = disabled;
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

function shouldAllowMixedFormat() {
  return Boolean(allowMixedFormatEl && allowMixedFormatEl.checked);
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

function formatElapsed(ms) {
  const totalSeconds = Math.floor(Math.max(0, ms) / 1000);
  const minutes = String(Math.floor(totalSeconds / 60)).padStart(2, '0');
  const seconds = String(totalSeconds % 60).padStart(2, '0');
  return `${minutes}:${seconds}`;
}

function toFormatSnapshot(font) {
  const raw = font && typeof font.toJSON === 'function' ? font.toJSON() : {};
  return createFormatSnapshotFromRaw(raw);
}

function loadFontForSnapshot(font) {
  if (font && typeof font.load === 'function') {
    font.load(FONT_LOAD_PROPERTIES);
  }
}

function renderChangePlan(changePlan) {
  beforeEl.textContent = changePlan.inputText || '';
  afterEl.textContent = changePlan.outputText || '';

  const summary = changePlan.summary || { totalBlocks: 0, convertedCount: 0, skippedCount: 0, noopCount: 0, vni: 0, tcvn3: 0, unicodeOrUnknown: 0 };
  const skipComment = changePlan.comments && changePlan.comments.length ? ` | ${changePlan.comments[0]}` : '';

  if (!changePlan.changed) {
    metaEl.textContent =
      `Không có thay đổi | block: ${summary.totalBlocks} | convert: ${summary.convertedCount} | skip: ${summary.skippedCount}` +
      ` | vni: ${summary.vni} | tcvn3: ${summary.tcvn3} | unicode/unknown: ${summary.unicodeOrUnknown}` +
      `${skipComment}`;
    return;
  }

  metaEl.textContent =
    `Block: ${summary.totalBlocks} | convert: ${summary.convertedCount} | skip: ${summary.skippedCount} | giữ nguyên: ${summary.noopCount}` +
    ` | vni: ${summary.vni} | tcvn3: ${summary.tcvn3} | unicode/unknown: ${summary.unicodeOrUnknown}` +
    `${skipComment}`;
}

function buildSelectionResult(selection, sourceMode, allowMixedFormat) {
  const unitPlans = buildUnitChangePlans(selection.units, { sourceMode, allowMixedFormat });
  const mergedPlan = mergeUnitChangePlans(unitPlans, selection.text);
  return { unitPlans, mergedPlan };
}

function createSelectionReadStats() {
  return {
    syncCount: 0,
    rangeCount: 0,
    unitCountBeforeFilter: 0,
    unitCountAfterFilter: 0,
    finalUnitCount: 0,
    fallbackReason: null,
    readMs: 0,
  };
}

async function syncSelectionRead(context, stats) {
  await context.sync();
  stats.syncCount += 1;
}

async function readSelectionBase(context, stats) {
  const range = context.document.getSelection();
  range.load('text,isEmpty');
  await syncSelectionRead(context, stats);
  return range;
}

async function readSelectionTextRanges(context, range, stats) {
  const textRanges = range.getTextRanges(TEXT_RANGE_DELIMITERS, false);
  textRanges.load('items');
  await syncSelectionRead(context, stats);
  stats.rangeCount = textRanges.items.length;
  return textRanges;
}

async function readTextRangeUnits(context, textRanges, stats) {
  for (const item of textRanges.items) {
    item.load('text');
    loadFontForSnapshot(item.font);
  }

  await syncSelectionRead(context, stats);
  stats.unitCountBeforeFilter = textRanges.items.length;

  const units = textRanges.items
    .map((item, index) => ({
      id: `unit-${index + 1}`,
      range: item,
      text: sanitizeUnitText(item.text || ''),
      format: toFormatSnapshot(item.font),
    }))
    .filter((item) => item.text.length > 0);

  stats.unitCountAfterFilter = units.length;
  return units;
}

async function readRangeFallbackUnit(context, range, stats, fallbackReason) {
  stats.fallbackReason = fallbackReason;
  loadFontForSnapshot(range.font);
  await syncSelectionRead(context, stats);

  const units = [
    {
      id: 'unit-1',
      range,
      text: sanitizeUnitText(range.text || ''),
      format: toFormatSnapshot(range.font),
    },
  ];

  if (!stats.unitCountBeforeFilter) {
    stats.unitCountBeforeFilter = 1;
    stats.unitCountAfterFilter = units.length;
  }

  return units;
}

function getRangeFallbackReason(range, units) {
  const selectedNormalized = normalizeTextWithoutSplitChars(sanitizeUnitText(range.text || ''));
  const unitsNormalized = normalizeTextWithoutSplitChars(units.map((item) => item.text || '').join(''));

  if (!selectedNormalized) {
    return 'empty-normalized-selection';
  }

  if (selectedNormalized !== unitsNormalized) {
    return 'text-range-mismatch';
  }

  return null;
}

function finalizeSelectionReadStats(stats, startedAt, units) {
  stats.finalUnitCount = units.length;
  stats.readMs = Date.now() - startedAt;
  return stats;
}

async function getSelectionData(context) {
  const startedAt = Date.now();
  const readStats = createSelectionReadStats();
  const range = await readSelectionBase(context, readStats);

  if (range.isEmpty) {
    const units = [];
    return {
      range,
      text: '',
      isEmpty: true,
      units,
      readStats: finalizeSelectionReadStats(readStats, startedAt, units),
    };
  }

  const textRanges = await readSelectionTextRanges(context, range, readStats);

  let units = [];
  if (textRanges.items.length) {
    units = await readTextRangeUnits(context, textRanges, readStats);

    const fallbackReason = getRangeFallbackReason(range, units);
    if (fallbackReason) {
      units = await readRangeFallbackUnit(context, range, readStats, fallbackReason);
    }
  } else {
    units = await readRangeFallbackUnit(context, range, readStats, 'no-text-ranges');
  }

  return {
    range,
    text: range.text || '',
    isEmpty: false,
    units,
    readStats: finalizeSelectionReadStats(readStats, startedAt, units),
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
  setControlsBusy(true);

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

      const { mergedPlan } = buildSelectionResult(selection, getSourceMode(), shouldAllowMixedFormat());
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
  } finally {
    setControlsBusy(false);
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

  setStatus('Đang phân tích vùng chọn...', 'info');
  setControlsBusy(true);

  try {
    await Word.run(async (context) => {
      const selection = await getSelectionData(context);
      if (selection.isEmpty) {
        setStatus('Không có vùng chọn để Apply. Hãy bôi đen text trước.', 'warn');
        return;
      }

      const { unitPlans, mergedPlan } = buildSelectionResult(selection, getSourceMode(), shouldAllowMixedFormat());

      const setTimes = shouldSetTimesNewRoman();
      const supportComments = canInsertWordComments();
      const totalBlocks = unitPlans.length;
      const startedAt = Date.now();
      let processedBlocks = 0;
      let convertApplied = 0;
      let fontApplied = 0;
      let runtimeErrors = 0;

      const progressPercent = totalBlocks > 0 ? 0 : 100;
      setStatus(`Đang áp dụng chuyển mã... 0/${totalBlocks} block (${progressPercent}%).`, 'info');

      const buildProgressMessage = () => {
        const percent = totalBlocks > 0 ? Math.round((processedBlocks * 100) / totalBlocks) : 100;
        const elapsed = Date.now() - startedAt;
        let message = `Đang áp dụng chuyển mã... ${processedBlocks}/${totalBlocks} block (${percent}%).`;
        if (elapsed > 2000) {
          message += ` Thời gian: ${formatElapsed(elapsed)}.`;
        }
        return message;
      };

      const progressTimer = setInterval(() => {
        if (Date.now() - startedAt > 2000) {
          setStatus(buildProgressMessage(), 'info');
        }
      }, 1000);

      try {
        for (const entry of unitPlans) {
          try {
            if (entry.action === 'skip') {
              if (supportComments && entry.comment) {
                entry.range.insertComment(entry.comment);
                await context.sync();
              }
            } else if (entry.action === 'convert') {
              const replaced = entry.range.insertText(entry.afterText, Word.InsertLocation.replace);
              if (setTimes) {
                replaced.font.name = 'Times New Roman';
                fontApplied += 1;
              }
              await context.sync();
              convertApplied += 1;
            } else if (setTimes) {
              entry.range.font.name = 'Times New Roman';
              await context.sync();
              fontApplied += 1;
            }
          } catch (_error) {
            runtimeErrors += 1;
          } finally {
            processedBlocks += 1;
          }
        }
      } finally {
        clearInterval(progressTimer);
      }

      state.lastPreview = mergedPlan;
      renderChangePlan(mergedPlan);

      const summary = mergedPlan.summary;
      const elapsed = Date.now() - startedAt;
      let message = `Đã xử lý theo từng block: convert ${convertApplied}/${summary.convertedCount}, skip ${summary.skippedCount}.`;
      if (setTimes) {
        message += ` Đặt Times New Roman cho ${fontApplied} block không bị skip.`;
      }
      if (elapsed > 2000) {
        message += ` Thời gian: ${formatElapsed(elapsed)}.`;
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
  } finally {
    setControlsBusy(false);
  }
}

function wireEvents() {
  previewBtn.addEventListener('click', previewSelection);
  applyBtn.addEventListener('click', applySelection);

  sourceModeEl.addEventListener('change', () => {
    setStatus('Đã thay đổi cấu hình nguồn bảng mã. Hãy Preview lại.', 'warn');
  });

  if (allowMixedFormatEl) {
    allowMixedFormatEl.addEventListener('change', () => {
      setStatus('Đã thay đổi tùy chọn mixed format. Hãy Preview lại trước khi Apply.', 'warn');
    });
  }

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
  setControlsBusy(false);
  setStatus('Sẵn sàng. Hãy chọn đoạn văn bản rồi bấm Preview.', 'ok');
});
