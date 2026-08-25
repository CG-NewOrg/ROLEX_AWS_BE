'use strict';

const XSS_PATTERNS = [
  // ---- Script tags ----------------------------------------------------------
  // { label: 'script tag',            pattern: /<\s*script[\s\S]*?>/i },
  // { label: 'closing script tag',    pattern: /<\s*\/\s*script\s*>/i },

  // ---- Dangerous HTML tags --------------------------------------------------
  // { label: 'iframe tag',            pattern: /<\s*i\s*frame[\s\S]*?>/i },
  // { label: 'object tag',            pattern: /<\s*object[\s\S]*?>/i },
  { label: 'embed tag',             pattern: /<\s*embed[\s\S]*?>/i },
  { label: 'link tag',              pattern: /<\s*link[\s\S]*?>/i },
  // { label: 'meta tag',              pattern: /<\s*meta[\s\S]*?>/i },
  { label: 'base tag',              pattern: /<\s*base[\s\S]*?>/i },
  // { label: 'form tag',              pattern: /<\s*form[\s\S]*?>/i },
  { label: 'img tag with event',    pattern: /<\s*img[^>]+(?:onerror|onload|src\s*=\s*["']?\s*javascript)[^>]*>/i },
  { label: 'svg tag',               pattern: /<\s*svg[\s\S]*?>/i },
  { label: 'math tag',              pattern: /<\s*math[\s\S]*?>/i },

  // ---- Inline event handlers ------------------------------------------------
  { label: 'inline event handler',
    pattern: /\bon\w+\s*=\s*["']?[^"'>]*(?:alert|confirm|prompt|eval|fetch|import|document|window|location|cookie|exec)/i },
  { label: 'generic event handler attribute',
    pattern: /\s(?:onclick|ondblclick|onmousedown|onmouseup|onmouseover|onmouseout|onmousemove|onkeydown|onkeyup|onkeypress|onfocus|onblur|onchange|onsubmit|onreset|onselect|onload|onunload|onabort|onerror|onresize|onscroll|oncontextmenu|oncopy|oncut|onpaste|oninput|oninvalid|onsearch|ontoggle|onwheel|ondrag|ondrop|onanimationstart|onanimationend|ontransitionend)\s*=/i },

  // ---- Dangerous URI schemes ------------------------------------------------
  // { label: 'javascript: URI',       pattern: /javascript\s*:/i },
  // { label: 'vbscript: URI',         pattern: /vbscript\s*:/i },
  // { label: 'data: URI (html/js)',   pattern: /data\s*:\s*(?:text\/html|application\/javascript|text\/javascript|application\/x-javascript)/i },

  // ---- Dangerous JavaScript functions / globals ----------------------------
  { label: 'alert() call',          pattern: /\balert\s*\(/i },
  { label: 'confirm() call',        pattern: /\bconfirm\s*\(/i },
  { label: 'prompt() call',         pattern: /\bprompt\s*\(/i },
  // { label: 'eval() call',           pattern: /\beval\s*\(/i },
  // { label: 'setTimeout() call',     pattern: /\bsetTimeout\s*\(/i },
  // { label: 'setInterval() call',    pattern: /\bsetInterval\s*\(/i },
  // { label: 'Function() constructor',pattern: /\bFunction\s*\(/i },
  // { label: 'execScript() call',     pattern: /\bexecScript\s*\(/i },
  // { label: 'importScripts() call',  pattern: /\bimportScripts\s*\(/i },

  // ---- DOM manipulation sinks -----------------------------------------------
  // { label: 'innerHTML assignment',  pattern: /\.innerHTML\s*=/i },
  { label: 'outerHTML assignment',  pattern: /\.outerHTML\s*=/i },
  { label: 'document.write()',      pattern: /document\s*\.\s*write\s*\(/i },
  { label: 'document.writeln()',    pattern: /document\s*\.\s*writeln\s*\(/i },
  { label: 'document.cookie',       pattern: /document\s*\.\s*cookie/i },
  { label: 'document.domain',       pattern: /document\s*\.\s*domain/i },
  // { label: 'window.location',       pattern: /window\s*\.\s*location/i },
  // { label: 'location.href',         pattern: /location\s*\.\s*href/i },
  { label: 'document.createElement',pattern: /document\s*\.\s*createElement\s*\(/i },

  // ---- Network / async fetch -----------------------------------------------
  // { label: 'fetch() call',          pattern: /\bfetch\s*\(/i },
  // { label: 'XMLHttpRequest',        pattern: /\bXMLHttpRequest\b/i },
  // { label: 'import() call',         pattern: /\bimport\s*\(/i },

  // // ---- CSS expression -------------------------------------------------------
  // { label: 'CSS expression()',      pattern: /expression\s*\(/i },

  // ---- Template literals with script-like content ---------------------------
  { label: 'backtick template with call',
    pattern: /`[^`]*\$\{[^}]*(?:alert|eval|fetch|document|window)[^}]*\}[^`]*`/i }
];


function decodeValue(value) {
  let decoded = value;

  // Pass 1 – URL-decode (%xx and %uXXXX)
  try { decoded = decodeURIComponent(decoded); } catch (_) { /* malformed – keep as-is */ }

  // Pass 2 – HTML entity decode (numeric &#NN; &#xNN; and named &lt; etc.)
  decoded = decoded
    .replace(/&#x([0-9a-f]+);?/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&#(\d+);?/gi, (_, dec) => String.fromCharCode(Number(dec)))
    .replace(/&lt;/gi,   '<')
    .replace(/&gt;/gi,   '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#x27;/gi, "'")
    .replace(/&amp;/gi,  '&');

  // Pass 3 – second URL-decode (handles double-encoded payloads)
  try { decoded = decodeURIComponent(decoded); } catch (_) { /* ignore */ }

  // Pass 4 – collapse Unicode look-alike whitespace between keywords and ( )
  decoded = decoded.replace(/[\u00a0\u2000-\u200b\u2028\u2029\ufeff]/g, ' ');

  return decoded;
}

function detectXSS(value, fieldPath = '') {
  if (typeof value === 'string') {
    const decoded = decodeValue(value);
    for (const { label, pattern } of XSS_PATTERNS) {
      if (pattern.test(decoded)) {
        return { label, field: fieldPath || 'input' };
      }
    }
    return null;
  }

  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      const hit = detectXSS(value[i], `${fieldPath}[${i}]`);
      if (hit) return hit;
    }
    return null;
  }

  if (value !== null && typeof value === 'object') {
    for (const key of Object.keys(value)) {
      const hit = detectXSS(value[key], fieldPath ? `${fieldPath}.${key}` : key);
      if (hit) return hit;
    }
    return null;
  }

  return null;
}

function xssProtectionMiddleware(req, res, next) {
  const targets = [
    { source: req.body,   label: 'body' },
    { source: req.query,  label: 'query' },
    { source: req.params, label: 'params' }
  ];

  for (const { source, label } of targets) {
    if (!source) continue;
    const hit = detectXSS(source, label);
    if (hit) {
      console.warn(
        `[XSS] Blocked request – pattern "${hit.label}" detected in ${hit.field} | ` +
        `method=${req.method} path=${req.path} ip=${req.ip}`
      );
      return res.status(400).json({
        error: 'Bad Request',
        message: `Potentially unsafe content detected in field "${hit.field}". ` +
                 `Blocked pattern: ${hit.label}.`
      });
    }
  }

  next();
}

module.exports = xssProtectionMiddleware;
module.exports.detectXSS   = detectXSS;
module.exports.decodeValue  = decodeValue;
module.exports.XSS_PATTERNS = XSS_PATTERNS;