/* ==========================================================================
   DocBrisk KPI Engine
   Pure functions, no Office.js dependency — testable in isolation and
   reusable from the web dashboard exporter, the standalone HTML export,
   and (later) a server-side version if this ever needs to run in a
   backend job instead of inside Excel.

   Nothing here is a machine-learning model. It is the same style of
   approach DocBrisk's résumé parser already uses: transparent, inspectable
   heuristics — column-name keyword matching plus statistical shape of the
   data — so every suggestion can be explained in one sentence, and wrong
   guesses are cheap to fix by hand.
   ========================================================================== */

const KpiEngine = (() => {

  /* ---------------------------------------------------------------------
     1. Column profiling
     --------------------------------------------------------------------- */

  const CURRENCY_RE = /[₹$€£]/;
  const PERCENT_RE = /%\s*$/;
  const NUMERIC_CLEAN_RE = /[₹$€£,\s]/g;

  const DATE_FORMATS = [
    /^\d{4}-\d{1,2}-\d{1,2}$/,                 // 2024-01-31
    /^\d{1,2}\/\d{1,2}\/\d{2,4}$/,              // 31/01/2024 or 1/5/24
    /^\d{1,2}-\d{1,2}-\d{2,4}$/,                 // 31-01-2024
    /^\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}$/,       // 31 Jan 2024
    /^[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{2,4}$/      // Jan 31, 2024
  ];

  function looksLikeDate(raw) {
    const t = String(raw).trim();
    if (!t) return false;
    if (!DATE_FORMATS.some((re) => re.test(t))) return false;
    const d = new Date(t);
    return !isNaN(d.getTime()) && d.getFullYear() > 1900 && d.getFullYear() < 2100;
  }

  function looksLikeNumber(raw) {
    const t = String(raw).trim();
    if (!t) return false;
    const cleaned = t.replace(NUMERIC_CLEAN_RE, '').replace(PERCENT_RE, '');
    if (!cleaned || !/^-?\d+(\.\d+)?$/.test(cleaned)) return false;
    return true;
  }

  function toNumber(raw) {
    const t = String(raw).trim().replace(NUMERIC_CLEAN_RE, '').replace(/%\s*$/, '');
    const n = parseFloat(t);
    return isNaN(n) ? null : n;
  }

  /* Column-name keyword map used to guess what a column MEANS, not just
     what shape it is. Ordered so more specific tags are checked first
     (e.g. "profit" before the generic "amount" bucket that "revenue"
     also matches). */
  const SEMANTIC_RULES = [
    ['profit', /\b(profit|margin|net\s?income)\b/i],
    ['revenue', /\b(revenue|sales|turnover|income|amount|price|total)\b/i],
    ['cost', /\b(cost|expense|spend|cogs|expenditure)\b/i],
    ['quantity', /\b(qty|quantity|units?|volume|count)\b/i],
    ['date', /\b(date|month|year|period|day|week|quarter)\b/i],
    ['geo', /\b(region|state|country|city|territory|zone)\b/i],
    ['entity', /\b(customer|client|user|account|employee|vendor|supplier)\b/i],
    ['category', /\b(category|segment|type|product|sku|department|channel)\b/i],
    ['status', /\b(status|stage|outcome)\b/i]
  ];

  function semanticTag(colName) {
    for (const [tag, re] of SEMANTIC_RULES) if (re.test(colName)) return tag;
    return null;
  }

  /**
   * Profile every column: infer its type from the actual cell values
   * (never from the header text alone — a column called "Date" full of
   * free text is still text), compute summary stats, and attach a
   * semantic guess used later to rank KPI suggestions.
   *
   * @param {string[]} headers
   * @param {Array<Array<string|number>>} rows  — excludes the header row
   */
  function profileColumns(headers, rows) {
    return headers.map((name, i) => {
      const values = rows.map((r) => r[i]).filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
      const sample = values.slice(0, 200); // enough to type a column reliably without scanning millions of rows
      const nullCount = rows.length - values.length;

      const numericHits = sample.filter(looksLikeNumber).length;
      const dateHits = sample.filter(looksLikeDate).length;
      const isNumeric = sample.length > 0 && numericHits / sample.length >= 0.9;
      const isDate = !isNumeric && sample.length > 0 && dateHits / sample.length >= 0.9;

      const distinct = new Set(values.map((v) => String(v).trim().toLowerCase()));
      // Categorical if it has few enough distinct values outright (works for
      // small datasets), OR the distinct values are a small share of all
      // rows (works for large datasets with many rows but still few
      // categories, e.g. 40 regions across 50,000 orders).
      const isCategorical = !isNumeric && !isDate && distinct.size >= 2 &&
        (distinct.size <= 20 || distinct.size / Math.max(1, values.length) <= 0.5);

      let type = 'text';
      if (isNumeric) type = 'numeric';
      else if (isDate) type = 'date';
      else if (isCategorical) type = 'categorical';

      const profile = {
        index: i, name, type,
        nullCount, rowCount: values.length,
        distinctCount: distinct.size,
        semantic: semanticTag(name)
      };

      if (type === 'numeric') {
        const nums = values.map(toNumber).filter((n) => n !== null);
        const sum = nums.reduce((a, b) => a + b, 0);
        profile.isCurrency = sample.some((v) => CURRENCY_RE.test(String(v)));
        profile.isPercent = sample.some((v) => PERCENT_RE.test(String(v)));
        profile.stats = {
          sum, count: nums.length,
          avg: nums.length ? sum / nums.length : 0,
          min: nums.length ? Math.min(...nums) : 0,
          max: nums.length ? Math.max(...nums) : 0
        };
      }

      if (type === 'date') {
        const dates = values.map((v) => new Date(v)).filter((d) => !isNaN(d.getTime())).sort((a, b) => a - b);
        profile.stats = {
          min: dates[0] || null,
          max: dates[dates.length - 1] || null,
          span_days: dates.length > 1 ? Math.round((dates[dates.length - 1] - dates[0]) / 86400000) : 0
        };
      }

      if (type === 'categorical') {
        const freq = new Map();
        values.forEach((v) => { const k = String(v).trim(); freq.set(k, (freq.get(k) || 0) + 1); });
        profile.categories = Array.from(freq.entries()).sort((a, b) => b[1] - a[1]);
      }

      return profile;
    });
  }

  /* ---------------------------------------------------------------------
     2. KPI suggestion
     --------------------------------------------------------------------- */

  const fmt = (n) => Math.abs(n) >= 1000
    ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
    : n.toLocaleString(undefined, { maximumFractionDigits: 2 });

  function numericColumns(profiles) { return profiles.filter((p) => p.type === 'numeric'); }
  function categoricalColumns(profiles) { return profiles.filter((p) => p.type === 'categorical'); }
  function dateColumns(profiles) { return profiles.filter((p) => p.type === 'date'); }

  let uidCounter = 0;
  const uid = () => 'kpi_' + (++uidCounter) + '_' + Date.now().toString(36);

  /**
   * Build a ranked list of candidate KPIs from the column profiles. Every
   * suggestion carries a `confidence` (0–100, used only to sort — never
   * shown as a false precision number to the end user) and an `explain`
   * string a human can sanity-check in one read.
   */
  function suggestKpis(profiles, rows) {
    const out = [];
    const nums = numericColumns(profiles);
    const cats = categoricalColumns(profiles);
    const dates = dateColumns(profiles);

    const bump = (tag) => (tag === 'revenue' || tag === 'cost' || tag === 'profit' || tag === 'quantity') ? 20 : 0;

    // Total / average of every numeric column.
    nums.forEach((col) => {
      out.push({
        id: uid(), kind: 'sum', columns: [col.name], label: `Total ${col.name}`,
        value: col.stats.sum, format: col.isCurrency ? 'currency' : col.isPercent ? 'percent' : 'number',
        chart: 'scorecard', confidence: 55 + bump(col.semantic),
        explain: `Sum of every value in "${col.name}".`
      });
      out.push({
        id: uid(), kind: 'avg', columns: [col.name], label: `Average ${col.name}`,
        value: col.stats.avg, format: col.isCurrency ? 'currency' : col.isPercent ? 'percent' : 'number',
        chart: 'scorecard', confidence: 40 + bump(col.semantic),
        explain: `Average of every value in "${col.name}".`
      });
    });

    // Revenue − Cost → profit + margin, only when both are unambiguous.
    const revenueCol = nums.find((c) => c.semantic === 'revenue');
    const costCol = nums.find((c) => c.semantic === 'cost');
    if (revenueCol && costCol) {
      const profit = revenueCol.stats.sum - costCol.stats.sum;
      const currencyLike = revenueCol.isCurrency || costCol.isCurrency;
      out.push({
        id: uid(), kind: 'formula', columns: [revenueCol.name, costCol.name], label: 'Total Profit',
        value: profit, format: currencyLike ? 'currency' : 'number', chart: 'scorecard', confidence: 90,
        explain: `"${revenueCol.name}" minus "${costCol.name}", summed across all rows.`
      });
      if (revenueCol.stats.sum) {
        out.push({
          id: uid(), kind: 'formula', columns: [revenueCol.name, costCol.name], label: 'Profit Margin %',
          value: (profit / revenueCol.stats.sum) * 100, format: 'percent', chart: 'scorecard', confidence: 88,
          explain: `Profit as a percentage of "${revenueCol.name}".`
        });
      }
    }

    // Time trend: a date column paired with a numeric column.
    if (dates.length && nums.length) {
      const dateCol = dates[0];
      nums.forEach((col) => {
        out.push({
          id: uid(), kind: 'trend', columns: [dateCol.name, col.name], label: `${col.name} Over Time`,
          chart: 'line', confidence: 60 + bump(col.semantic),
          explain: `"${col.name}" plotted against "${dateCol.name}", grouped by month.`,
          buildSeries: () => buildTimeSeries(rows, profiles, dateCol, col)
        });
      });
    }

    // Category breakdowns: top-N by a numeric measure, and a plain count.
    cats.forEach((cat) => {
      if (cat.distinctCount >= 2 && cat.distinctCount <= 20) {
        out.push({
          id: uid(), kind: 'countBy', columns: [cat.name], label: `Records by ${cat.name}`,
          chart: 'bar', confidence: 35 + bump(cat.semantic),
          explain: `Row count grouped by "${cat.name}".`,
          buildSeries: () => cat.categories.slice(0, 12).map(([k, v]) => ({ label: k, value: v }))
        });
        nums.forEach((col) => {
          out.push({
            id: uid(), kind: 'topN', columns: [cat.name, col.name], label: `Top ${cat.name} by ${col.name}`,
            chart: 'bar', confidence: 45 + bump(cat.semantic) + bump(col.semantic),
            explain: `"${col.name}" summed within each "${cat.name}", top 10 shown.`,
            buildSeries: () => buildTopN(rows, profiles, cat, col)
          });
        });
      }
    });

    // Always-useful baseline KPIs.
    out.push({
      id: uid(), kind: 'count', columns: [], label: 'Total Records',
      value: rows.length, format: 'number', chart: 'scorecard', confidence: 30,
      explain: 'Number of rows in the dataset.'
    });
    const entityCol = profiles.find((p) => p.semantic === 'entity');
    if (entityCol) {
      out.push({
        id: uid(), kind: 'distinct', columns: [entityCol.name], label: `Distinct ${entityCol.name}`,
        value: entityCol.distinctCount, format: 'number', chart: 'scorecard', confidence: 50,
        explain: `Count of unique values in "${entityCol.name}".`
      });
    }

    return capByKind(out.sort((a, b) => b.confidence - a.confidence)).slice(0, 16);
  }

  /* A dataset with several categorical columns generates a "Top X by Y"
     candidate for every category×measure pair, which can otherwise fill
     the entire suggestion list on its own and push out low-flash but
     broadly useful KPIs like Total Records or Distinct Customer. Cap each
     kind's contribution, keeping its highest-confidence members (the list
     is already confidence-sorted when this runs). */
  const KIND_CAP = { topN: 4, countBy: 3, trend: 3 };
  function capByKind(sortedList) {
    const seen = {};
    return sortedList.filter((k) => {
      const cap = KIND_CAP[k.kind];
      if (!cap) return true;
      seen[k.kind] = (seen[k.kind] || 0) + 1;
      return seen[k.kind] <= cap;
    });
  }

  function monthKey(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); }

  function buildTimeSeries(rows, profiles, dateCol, numCol) {
    const buckets = new Map();
    rows.forEach((r) => {
      const d = new Date(r[dateCol.index]);
      const n = toNumber(r[numCol.index]);
      if (isNaN(d.getTime()) || n === null) return;
      const k = monthKey(d);
      buckets.set(k, (buckets.get(k) || 0) + n);
    });
    return Array.from(buckets.entries()).sort(([a], [b]) => a.localeCompare(b))
      .map(([label, value]) => ({ label, value }));
  }

  function buildTopN(rows, profiles, catCol, numCol) {
    const buckets = new Map();
    rows.forEach((r) => {
      const k = String(r[catCol.index] ?? '').trim();
      const n = toNumber(r[numCol.index]);
      if (!k || n === null) return;
      buckets.set(k, (buckets.get(k) || 0) + n);
    });
    return Array.from(buckets.entries()).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([label, value]) => ({ label, value }));
  }

  /* ---------------------------------------------------------------------
     3. Custom KPI formula parser — no eval(), no Function() constructor.
     Grammar: expr := term (('+'|'-') term)*
              term := factor (('*'|'/') factor)*
              factor := NUMBER | AGGFN '(' COLUMN ')' | '(' expr ')'
              AGGFN := SUM | AVG | COUNT | MIN | MAX | DISTINCT
     A formula operates on already-aggregated column values, because a
     KPI is a single summary number — not a per-row calculated column.
     --------------------------------------------------------------------- */

  class FormulaError extends Error {}

  const AGG_FNS = new Set(['SUM', 'AVG', 'COUNT', 'MIN', 'MAX', 'DISTINCT']);

  /* Pass 1: find every "FUNC(column name)" call — quoted or not, spaces
     and all — and replace it with its computed numeric value. Doing this
     as a dedicated pass (rather than trying to tokenize column names as
     generic identifiers) is what lets "SUM(Order Date)" and
     SUM(Revenue) - SUM(Cost)" both work without requiring quotes around
     every column name. */
  function substituteAggregates(src, profiles, rows) {
    const byName = new Map(profiles.map((p) => [p.name.trim().toLowerCase(), p]));
    const callRe = /([A-Za-z_]+)\s*\(\s*(?:"([^"]+)"|'([^']+)'|([^()]+?))\s*\)/g;
    let sawCall = false;
    const replaced = src.replace(callRe, (whole, fn, q1, q2, bare) => {
      sawCall = true;
      const fnUp = fn.trim().toUpperCase();
      if (!AGG_FNS.has(fnUp)) {
        throw new FormulaError(`Unknown function "${fn}". Use SUM, AVG, COUNT, MIN, MAX or DISTINCT.`);
      }
      const colName = (q1 ?? q2 ?? bare ?? '').trim();
      const col = byName.get(colName.toLowerCase());
      if (!col) throw new FormulaError(`No column named "${colName}".`);
      return String(aggregate(fnUp, col, rows));
    });
    if (!sawCall) throw new FormulaError('A KPI formula needs at least one function, e.g. SUM(Revenue).');
    return replaced;
  }

  /* Pass 2: what's left after substitution should be plain arithmetic —
     numbers, + - * / ( ) — nothing else. A stray letter at this point
     means something wasn't recognised as a function call. */
  function evaluateArithmetic(src) {
    if (/[A-Za-z]/.test(src)) {
      throw new FormulaError('Formula contains text outside of a function call, e.g. SUM(...), AVG(...).');
    }
    const tokens = [];
    const re = /\s*(\d+(?:\.\d+)?|[+\-*/()])/g;
    let m, last = 0;
    while ((m = re.exec(src))) {
      if (m.index !== last) throw new FormulaError(`Unexpected character near "${src.slice(last, m.index)}"`);
      tokens.push(/^[0-9.]/.test(m[1]) ? { t: 'NUM', v: parseFloat(m[1]) } : { t: m[1] });
      last = re.lastIndex;
    }
    if (last !== src.length) throw new FormulaError(`Unexpected character near "${src.slice(last)}"`);
    if (!tokens.length) throw new FormulaError('Empty formula.');

    let pos = 0;
    const peek = () => tokens[pos];
    const next = () => tokens[pos++];
    function parseExpr() {
      let v = parseTerm();
      while (peek() && (peek().t === '+' || peek().t === '-')) {
        const op = next().t; const rhs = parseTerm();
        v = op === '+' ? v + rhs : v - rhs;
      }
      return v;
    }
    function parseTerm() {
      let v = parseFactor();
      while (peek() && (peek().t === '*' || peek().t === '/')) {
        const op = next().t; const rhs = parseFactor();
        if (op === '/' && rhs === 0) throw new FormulaError('Division by zero.');
        v = op === '*' ? v * rhs : v / rhs;
      }
      return v;
    }
    function parseFactor() {
      const tok = peek();
      if (!tok) throw new FormulaError('Unexpected end of formula.');
      if (tok.t === 'NUM') { next(); return tok.v; }
      if (tok.t === '(') { next(); const v = parseExpr(); expectParen(); return v; }
      if (tok.t === '-') { next(); return -parseFactor(); }
      throw new FormulaError('Formula must start with a number or "(".');
    }
    function expectParen() {
      const tok = next();
      if (!tok || tok.t !== ')') throw new FormulaError('Missing closing ")".');
    }

    const result = parseExpr();
    if (pos !== tokens.length) throw new FormulaError('Unexpected text after the end of the formula.');
    return result;
  }

  function evaluateFormula(src, profiles, rows) {
    const substituted = substituteAggregates(String(src || '').trim(), profiles, rows);
    return evaluateArithmetic(substituted);
  }

  function aggregate(fn, col, rows) {
    const values = rows.map((r) => r[col.index]).filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
    if (fn === 'COUNT') return values.length;
    if (fn === 'DISTINCT') return new Set(values.map((v) => String(v).trim().toLowerCase())).size;
    const nums = values.map(toNumber).filter((n) => n !== null);
    if (!nums.length) return 0;
    if (fn === 'SUM') return nums.reduce((a, b) => a + b, 0);
    if (fn === 'AVG') return nums.reduce((a, b) => a + b, 0) / nums.length;
    if (fn === 'MIN') return Math.min(...nums);
    if (fn === 'MAX') return Math.max(...nums);
    throw new FormulaError(`Unknown function "${fn}".`);
  }

  function buildCustomKpi(label, formula, profiles, rows, format) {
    const value = evaluateFormula(formula, profiles, rows);
    return {
      id: uid(), kind: 'custom', columns: [], label, value, format: format || 'number',
      chart: 'scorecard', confidence: 100, formula,
      explain: `Custom formula: ${formula}`
    };
  }

  return {
    profileColumns, suggestKpis, buildCustomKpi, evaluateFormula, FormulaError,
    formatNumber: fmt, toNumber, looksLikeNumber, looksLikeDate
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = KpiEngine;
