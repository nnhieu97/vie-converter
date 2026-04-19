import CHARSETS from './vietnamese-charsets.json';

const conversionCache = new Map();
const SUPPORTED_SOURCES_SET = new Set(['tcvn3', 'vni']);

const VNI_LEGACY_CHAR_REGEX = /[\u00F6\u00D6\u00FC\u00DC\u00E6\u00C6\u00E5\u00C5\u00F1\u00D1\u00EF\u00CF\u00E4\u00C4]/g;
const TCVN3_LEGACY_CHAR_REGEX = /[\u00B5\u00B8\u00A9\u00B7\u00A1\u00A7\u00CC\u00D0\u00DD\u00D3\u00A8\u00AE\u00DC\u00AC\u00AD\u00B9\u00B6\u00CA\u00BB\u00BC\u00BD\u00C6]/g;
const VI_UNICODE_CHAR_REGEX = /[\u00E0-\u00E3\u00E8-\u00EA\u00EC-\u00ED\u00F2-\u00F5\u00F9-\u00FA\u00FD\u0103\u0102\u0111\u0110\u0129\u0128\u0169\u0168\u01A1\u01A0\u01AF\u01AE\u1EA0-\u1EF9]/gu;

const VNI_LEGACY_REGEX = new RegExp(VNI_LEGACY_CHAR_REGEX.source);
const VNI_Y_ALIAS_REGEX = /[yY][\u00FB\u00F5\u00EF\u00DB\u00D5\u00CF]/;
const TCVN3_LEGACY_REGEX = new RegExp(TCVN3_LEGACY_CHAR_REGEX.source);
const VI_UNICODE_REGEX = new RegExp(VI_UNICODE_CHAR_REGEX.source, 'u');
const SOURCE_EVIDENCE_WEIGHT = 1;
// Some documents use y + tone aliases in VNI (e.g. yû, yõ, yï).
// Normalize them directly to Unicode to avoid ambiguous single-byte mappings.
const VNI_ALIAS_MAP = Object.freeze({
  'y\u00FB': '\u1EF7', // yû -> ỷ
  'y\u00F5': '\u1EF9', // yõ -> ỹ
  'y\u00EF': '\u1EF5', // yï -> ỵ
  'Y\u00DB': '\u1EF6', // YÛ -> Ỷ
  'Y\u00D5': '\u1EF8', // YÕ -> Ỹ
  'Y\u00CF': '\u1EF4', // YÏ -> Ỵ
});

function buildLegacyTokens(sourceKey) {
  const chars = CHARSETS[sourceKey] || [];
  const unique = new Set();

  for (const token of chars) {
    if (!token) continue;
    const isLegacyToken = /[^\x00-\x7F]/.test(token) || token.length > 1;
    if (!isLegacyToken) continue;
    unique.add(token);
  }

  return Array.from(unique).sort((a, b) => b.length - a.length);
}

const SOURCE_LEGACY_TOKENS = Object.freeze({
  vni: Object.freeze([...buildLegacyTokens('VNI'), ...Object.keys(VNI_ALIAS_MAP)]),
  tcvn3: buildLegacyTokens('TCVN3'),
});

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getCacheKey(text, fromCharset, toCharset) {
  return `${fromCharset}:${toCharset}:${text}`;
}

function convertWithLongestMatch(text, fromChars, toChars) {
  if (!text) return '';

  const sorted = fromChars
    .map((value, index) => ({ value, index }))
    .filter((item) => item.value)
    .sort((a, b) => b.value.length - a.value.length);

  let result = text;
  const placeholderPrefix = '__VIE_MAP__';

  for (const item of sorted) {
    const token = `${placeholderPrefix}${item.index}__`;
    const re = new RegExp(escapeRegExp(item.value), 'g');
    result = result.replace(re, token);
  }

  for (let i = 0; i < toChars.length; i += 1) {
    const token = `${placeholderPrefix}${i}__`;
    const re = new RegExp(escapeRegExp(token), 'g');
    result = result.replace(re, toChars[i]);
  }

  return result;
}

function normalizeVniAliases(text) {
  let normalized = text;
  for (const [alias, canonical] of Object.entries(VNI_ALIAS_MAP)) {
    normalized = normalized.replace(new RegExp(escapeRegExp(alias), 'g'), canonical);
  }
  return normalized;
}

export function convert(text, fromCharset, toCharset) {
  if (!text) return '';

  const fromKey = String(fromCharset || '').toUpperCase();
  const toKey = String(toCharset || '').toUpperCase();
  const fromChars = CHARSETS[fromKey];
  const toChars = CHARSETS[toKey];

  if (!fromChars || !toChars) {
    throw new Error('Charset is not valid');
  }

  if (fromKey === toKey) {
    return text;
  }

  const cacheKey = getCacheKey(text, fromKey, toKey);
  if (conversionCache.has(cacheKey)) {
    return conversionCache.get(cacheKey);
  }

  const result = convertWithLongestMatch(text, fromChars, toChars);

  if (conversionCache.size > 1000) {
    const firstKey = conversionCache.keys().next().value;
    if (firstKey) conversionCache.delete(firstKey);
  }
  conversionCache.set(cacheKey, result);

  return result;
}

export function toUnicode(text, currentCharset) {
  return convert(text, currentCharset, 'unicode');
}

export function toVNI(text, currentCharset) {
  return convert(text, currentCharset, 'vni');
}

export function toTCVN3(text, currentCharset) {
  return convert(text, currentCharset, 'tcvn3');
}

export function toVIQR(text, currentCharset) {
  return convert(text, currentCharset, 'viqr');
}

export function createConverter(targetCharset) {
  return (text, currentCharset) => convert(text, currentCharset, targetCharset);
}

// Based on vietnamese-conversion detectCharset logic.
export function detectCharset(text) {
  if (!text) return null;

  const vniAsciiMatches = text.match(/[aAeEoOuUyY][1-9]|[dD]9/g) || [];
  const hasVNIAscii = vniAsciiMatches.length >= 2;
  const hasVNILegacy = VNI_LEGACY_REGEX.test(text) || VNI_Y_ALIAS_REGEX.test(text);
  const hasVNI = hasVNILegacy || hasVNIAscii;
  const hasVIQR = /[\^+\(\)\\'`?~.]/.test(text);
  const hasUnicode = /[\u00E0-\u00E3\u00E8-\u00EA\u00EC-\u00ED\u00F2-\u00F5\u00F9-\u00FA\u00FD\u0103\u0102\u0111\u0110\u0129\u0128\u0169\u0168\u01A1\u01A0\u01AF\u01AE\u1EA0-\u1EF9]/u.test(text.toLowerCase());
  const hasTCVN3 = /[\u00B5\u00B8\u00A9\u00B7\u00A1\u00A7\u00CC\u00D0\u00DD\u00D3\u00A8\u00AE\u00DC\u00AC\u00AD\u00B9\u00B6\u00CA\u00BB\u00BC\u00BD\u00C6]/.test(text) && !hasUnicode && !hasVIQR && !hasVNI;

  if (hasUnicode) return 'unicode';
  if (hasVIQR && !hasVNI) return 'viqr';
  if (hasVNI && !hasVIQR) return 'vni';
  if (hasTCVN3) return 'tcvn3';
  return null;
}

export function inferCharsetFromFont(fontName) {
  const normalized = (fontName || '').toLowerCase();
  if (!normalized) return null;

  // Legacy .Vn* fonts are typically TCVN3 (ABC), not VNI.
  if (normalized.includes('.vn') || normalized.includes('vntime') || normalized.includes('vnarial') || normalized.includes('tcvn')) {
    return 'tcvn3';
  }

  if (normalized.includes('vni')) {
    return 'vni';
  }

  return null;
}

export function shouldUppercaseForTCVNFont(fontName, effectiveSource) {
  if (effectiveSource !== 'tcvn3') {
    return false;
  }

  const normalized = (fontName || '').trim().toLowerCase();
  if (!normalized) {
    return false;
  }

  // TCVN convention: font names ending with "H" are uppercase variants.
  return normalized.endsWith('h');
}

function countRegexMatches(text, regex) {
  if (!text) return 0;
  const matches = text.match(regex);
  return matches ? matches.length : 0;
}

export function scoreSourceEvidence(text, source) {
  if (!text) return 0;
  const tokens = SOURCE_LEGACY_TOKENS[source] || [];
  let score = 0;

  for (const token of tokens) {
    const matches = text.match(new RegExp(escapeRegExp(token), 'g'));
    if (matches) {
      score += matches.length * token.length;
    }
  }

  return score;
}

export function scoreUnicodeReadability(original, converted) {
  const unicodeCount = countRegexMatches(converted, VI_UNICODE_CHAR_REGEX);
  const originalLegacyCount =
    countRegexMatches(original, VNI_LEGACY_CHAR_REGEX) +
    countRegexMatches(original, TCVN3_LEGACY_CHAR_REGEX);
  const convertedLegacyCount =
    countRegexMatches(converted, VNI_LEGACY_CHAR_REGEX) +
    countRegexMatches(converted, TCVN3_LEGACY_CHAR_REGEX);
  const replacementCount = countRegexMatches(converted, /\uFFFD/g);

  let score = 0;
  score += unicodeCount * 3;
  score += Math.max(0, originalLegacyCount - convertedLegacyCount) * 2;
  score -= convertedLegacyCount * 3;
  score -= replacementCount * 20;
  if (converted === original) score -= 8;

  return score;
}

export function convertToUnicodeSafe(text, sourceCharset) {
  const sourceKey = (sourceCharset || '').toUpperCase();
  const fromChars = CHARSETS[sourceKey];
  const toChars = CHARSETS.UNICODE;
  const preparedText = sourceKey === 'VNI' ? normalizeVniAliases(text) : text;

  if (!fromChars || !toChars || sourceKey === 'UNICODE') {
    return toUnicode(preparedText, sourceCharset);
  }

  if (sourceKey === 'VNI' || sourceKey === 'TCVN3') {
    return convertWithLongestMatch(preparedText, fromChars, toChars);
  }

  return toUnicode(preparedText, sourceCharset);
}

export function detectSourceByContent(text, options = {}) {
  const { fontName = '' } = options;
  if (!text) return { source: null, hint: null };

  const originalScore = scoreUnicodeReadability(text, text);
  const scored = [];
  let best = {
    source: null,
    score: Number.NEGATIVE_INFINITY,
  };

  for (const source of SUPPORTED_SOURCES_SET) {
    try {
      const converted = convertToUnicodeSafe(text, source);
      const readabilityScore = scoreUnicodeReadability(text, converted);
      const evidenceScore = scoreSourceEvidence(text, source);
      const score = readabilityScore + evidenceScore * SOURCE_EVIDENCE_WEIGHT;
      scored.push({ source, score });
      if (score > best.score) {
        best = { source, score };
      }
    } catch (_error) {
      // Ignore invalid candidate conversion.
    }
  }

  if (!best.source) {
    return { source: null, hint: null };
  }

  // Require a clear gain over original to avoid false positives.
  if (best.score >= originalScore + 3) {
    const sorted = scored.slice().sort((a, b) => b.score - a.score);
    const second = sorted[1] || null;
    const fontHint = inferCharsetFromFont(fontName);

    // If scores are nearly tied, use font as a tie-break hint.
    if (fontHint && second && Math.abs(best.score - second.score) <= 1) {
      const hinted = scored.find((item) => item.source === fontHint);
      if (hinted && hinted.score >= originalScore + 3) {
        return {
          source: fontHint,
          hint: `Auto fallback: điểm ${best.source}/${second.source} gần nhau, ưu tiên theo font '${fontName}' => ${fontHint}.`,
        };
      }
    }

    return {
      source: best.source,
      hint: `Auto fallback: chọn ${best.source} dựa trên phân tích ký tự.`,
    };
  }

  return { source: null, hint: null };
}

export function detectLegacyFallbackSource(text, fontName = '') {
  const contentDecision = detectSourceByContent(text, { fontName });
  if (contentDecision.source) {
    return contentDecision;
  }

  const fontHint = inferCharsetFromFont(fontName);
  const hasVniPattern = VNI_LEGACY_REGEX.test(text) || VNI_Y_ALIAS_REGEX.test(text);
  const hasTcvnPattern = TCVN3_LEGACY_REGEX.test(text);
  const hasUnicodeVietnamese = VI_UNICODE_REGEX.test(text);

  if (fontHint === 'vni' && hasVniPattern) {
    return {
      source: 'vni',
      hint: `Auto fallback: detect unicode nhưng font '${fontName}' cho thấy VNI.`,
    };
  }

  if (fontHint === 'tcvn3' && hasTcvnPattern) {
    return {
      source: 'tcvn3',
      hint: `Auto fallback: detect unicode nhưng font '${fontName}' cho thấy TCVN3.`,
    };
  }

  if (!hasUnicodeVietnamese && hasVniPattern) {
    return {
      source: 'vni',
      hint: 'Auto fallback: chuỗi có mẫu ký tự legacy giống VNI.',
    };
  }

  if (!hasUnicodeVietnamese && hasTcvnPattern) {
    return {
      source: 'tcvn3',
      hint: 'Auto fallback: chuỗi có mẫu ký tự legacy giống TCVN3.',
    };
  }

  return { source: null, hint: null };
}

export function detectSourceForUnicodeConversion(text, options = {}) {
  const { fontName = '', sourceMode = 'auto' } = options;
  const normalizedMode = String(sourceMode || 'auto').toLowerCase();

  if (normalizedMode !== 'auto') {
    if (SUPPORTED_SOURCES_SET.has(normalizedMode)) {
      return {
        detected: null,
        effectiveSource: normalizedMode,
        hint: null,
        reason: null,
      };
    }

    return {
      detected: null,
      effectiveSource: null,
      hint: null,
      reason: `Nguồn bảng mã '${sourceMode}' không được hỗ trợ.`,
    };
  }

  const detected = detectCharset(text);
  if (detected && SUPPORTED_SOURCES_SET.has(detected)) {
    return {
      detected,
      effectiveSource: detected,
      hint: null,
      reason: null,
    };
  }

  if (!detected || detected === 'unicode') {
    const fallback = detectLegacyFallbackSource(text, fontName);
    if (fallback.source) {
      return {
        detected,
        effectiveSource: fallback.source,
        hint: fallback.hint,
        reason: null,
      };
    }

    if (!detected) {
      return {
        detected,
        effectiveSource: null,
        hint: null,
        reason: 'Không nhận diện được bảng mã. Bạn có thể thử chọn thủ công TCVN3 hoặc VNI.',
      };
    }

    return {
      detected,
      effectiveSource: null,
      hint: null,
      reason: 'Đoạn đang là Unicode, không cần chuyển.',
    };
  }

  return {
    detected,
    effectiveSource: null,
    hint: null,
    reason: `Phát hiện '${detected}', hiện chỉ hỗ trợ chuyển từ TCVN3/VNI sang Unicode.`,
  };
}

export function buildConversionPlan(text, options = {}) {
  const { sourceMode = 'auto', fontName = '' } = options;
  const original = text || '';

  if (!original.trim()) {
    return {
      changed: false,
      reason: 'Vùng chọn rỗng. Hãy bôi đen đoạn cần chuyển mã trước.',
      original,
      converted: original,
      detected: null,
      effectiveSource: null,
      hint: null,
      fontName,
    };
  }

  const decision = detectSourceForUnicodeConversion(original, { sourceMode, fontName });
  if (!decision.effectiveSource) {
    return {
      changed: false,
      reason: decision.reason || 'Không nhận diện được bảng mã.',
      original,
      converted: original,
      detected: decision.detected,
      effectiveSource: null,
      hint: decision.hint,
      fontName,
    };
  }

  try {
    let converted = convertToUnicodeSafe(original, decision.effectiveSource);
    let hint = decision.hint;

    if (shouldUppercaseForTCVNFont(fontName, decision.effectiveSource)) {
      converted = converted.toLocaleUpperCase('vi-VN');
      hint = hint
        ? `${hint} | Áp dụng chữ hoa do font TCVN kết thúc bằng H.`
        : 'Áp dụng chữ hoa do font TCVN kết thúc bằng H.';
    }

    const changed = converted !== original;

    return {
      changed,
      reason: changed ? null : 'Không có thay đổi sau khi chuyển mã.',
      original,
      converted,
      detected: decision.detected,
      effectiveSource: decision.effectiveSource,
      hint,
      fontName,
    };
  } catch (error) {
    return {
      changed: false,
      reason: `Lỗi chuyển mã: ${error?.message || String(error)}`,
      original,
      converted: original,
      detected: decision.detected,
      effectiveSource: decision.effectiveSource,
      hint: decision.hint,
      fontName,
    };
  }
}

export const SUPPORTED_SOURCES = Object.freeze(Array.from(SUPPORTED_SOURCES_SET));
export { CHARSETS };
