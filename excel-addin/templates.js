/* ==========================================================================
   DocBrisk KPI Dashboard Templates
   Each template is a grid of "slots". A slot has a preferred KPI kind
   (used to auto-fill it from the suggestions list) and a size. The user
   can always drag a different selected KPI into any slot afterwards —
   these are starting points, not a fixed layout.

   `lookerReportId` is intentionally blank: it's a one-time setup step for
   whoever owns this add-in (see README "Looker Studio setup"), not
   something that can be generated automatically. See looker-studio.js.
   ========================================================================== */

const KpiTemplates = (() => {
  const TEMPLATES = [
    {
      id: 'sales-overview',
      name: 'Sales Overview',
      description: 'Revenue, profit and trend for a typical sales export.',
      lookerReportId: '',
      slots: [
        { id: 's1', size: 'small', prefer: ['sum'], keyword: /revenue|sales/i },
        { id: 's2', size: 'small', prefer: ['sum'], keyword: /cost/i },
        { id: 's3', size: 'small', prefer: ['formula'], keyword: /profit/i },
        { id: 's4', size: 'small', prefer: ['count'], keyword: /records/i },
        { id: 's5', size: 'wide', prefer: ['trend'], keyword: /over time/i },
        { id: 's6', size: 'wide', prefer: ['topN'], keyword: /top .* by/i }
      ]
    },
    {
      id: 'finance-summary',
      name: 'Finance Summary',
      description: 'Margins and cost breakdowns for a P&L-style export.',
      lookerReportId: '',
      slots: [
        { id: 's1', size: 'small', prefer: ['formula'], keyword: /margin/i },
        { id: 's2', size: 'small', prefer: ['sum'], keyword: /cost/i },
        { id: 's3', size: 'small', prefer: ['avg'], keyword: /average/i },
        { id: 's4', size: 'wide', prefer: ['trend'], keyword: /over time/i },
        { id: 's5', size: 'wide', prefer: ['countBy'], keyword: /records by/i }
      ]
    },
    {
      id: 'operations',
      name: 'Operations',
      description: 'Volumes and distribution across categories or regions.',
      lookerReportId: '',
      slots: [
        { id: 's1', size: 'small', prefer: ['count'], keyword: /records/i },
        { id: 's2', size: 'small', prefer: ['distinct'], keyword: /distinct/i },
        { id: 's3', size: 'wide', prefer: ['countBy'], keyword: /records by/i },
        { id: 's4', size: 'wide', prefer: ['topN'], keyword: /top .* by/i }
      ]
    },
    {
      id: 'blank',
      name: 'Blank canvas',
      description: 'Start empty and place any four KPIs yourself.',
      lookerReportId: '',
      slots: [
        { id: 's1', size: 'small', prefer: [], keyword: null },
        { id: 's2', size: 'small', prefer: [], keyword: null },
        { id: 's3', size: 'small', prefer: [], keyword: null },
        { id: 's4', size: 'wide', prefer: [], keyword: null }
      ]
    }
  ];

  function list() { return TEMPLATES; }
  function get(id) { return TEMPLATES.find((t) => t.id === id) || TEMPLATES[TEMPLATES.length - 1]; }

  /** Auto-fill a template's slots from the ranked KPI suggestions. Every
   *  slot gets the best still-unused match; leftover slots stay empty for
   *  the user to fill by hand. Never mutates the input KPI list. */
  function autoFill(template, rankedKpis) {
    const used = new Set();
    return template.slots.map((slot) => {
      let pick = null;
      if (slot.keyword || slot.prefer.length) {
        pick = rankedKpis.find((k) =>
          !used.has(k.id) &&
          (slot.prefer.length === 0 || slot.prefer.includes(k.kind)) &&
          (!slot.keyword || slot.keyword.test(k.label)));
      }
      if (!pick) pick = rankedKpis.find((k) => !used.has(k.id));
      if (pick) used.add(pick.id);
      return { slotId: slot.id, size: slot.size, kpi: pick || null };
    });
  }

  return { list, get, autoFill };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = KpiTemplates;
