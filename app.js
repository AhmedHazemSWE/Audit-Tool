(function() {
  const SECTION_KEYS = ["customerOnly", "invoiceOnly", "unreconciled", "reconciled"];
  const SECTION_TITLES = {
    customerOnly: "Customer Only",
    invoiceOnly: "Invoice Only",
    unreconciled: "Unreconciled",
    reconciled: "Reconciled"
  };

  let xlsxLoaderPromise = null;

  function getProjectNameValue() {
    const input = qs('#project-name-input');
    if (!input) return '';
    return input.value.trim();
  }

  function formatProjectNameForFilename(name) {
    if (!name) return '';
    const cleaned = name.replace(/[\\/:*?"<>|]/g, '').trim();
    return cleaned.replace(/\s+/g, '-');
  }

  function ensureXLSXLoaded() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxLoaderPromise) {
      xlsxLoaderPromise = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        script.async = true;
        script.onload = () => {
          if (window.XLSX) resolve(window.XLSX);
          else reject(new Error('Failed to load XLSX library.'));
        };
        script.onerror = () => reject(new Error('Unable to load XLSX library.'));
        document.head.appendChild(script);
      });
    }
    return xlsxLoaderPromise;
  }

  const qs = (s, root = document) => root.querySelector(s);
  const qsa = (s, root = document) => Array.from(root.querySelectorAll(s));

  function createElement(tag, attrs = {}, children = []) {
    const el = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
      if (k === 'class') el.className = v; else if (k === 'text') el.textContent = v; else el.setAttribute(k, v);
    });
    children.forEach(c => el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c));
    return el;
  }

  function normalizeName(name) {
    if (!name) return "";
    return name
      .trim()
      .replace(/[^a-zA-Z0-9\s]/g, '')
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  function parseNames(multilineValue) {
    return multilineValue
      .split(/\r?\n/)
      .map(s => s.trim())
      .filter(Boolean);
  }

  function toSetCaseInsensitive(names) {
    const set = new Set();
    for (const n of names) {
      const norm = normalizeName(n);
      if (norm) set.add(norm);
    }
    return set;
  }

  function compareNameSets(leftNames, rightNames) {
    const leftSet = toSetCaseInsensitive(leftNames);
    const rightSet = toSetCaseInsensitive(rightNames);

    const inBoth = [];
    const onlyLeft = [];
    const onlyRight = [];

    const originalLeftMap = new Map();
    for (const n of leftNames) {
      const key = normalizeName(n);
      if (key && !originalLeftMap.has(key)) originalLeftMap.set(key, n);
    }
    const originalRightMap = new Map();
    for (const n of rightNames) {
      const key = normalizeName(n);
      if (key && !originalRightMap.has(key)) originalRightMap.set(key, n);
    }

    for (const key of leftSet) {
      if (rightSet.has(key)) inBoth.push(originalLeftMap.get(key) || key);
      else onlyLeft.push(originalLeftMap.get(key) || key);
    }
    for (const key of rightSet) {
      if (!leftSet.has(key)) onlyRight.push(originalRightMap.get(key) || key);
    }

    inBoth.sort((a,b) => a.localeCompare(b));
    onlyLeft.sort((a,b) => a.localeCompare(b));
    onlyRight.sort((a,b) => a.localeCompare(b));

    const exactSame = leftSet.size === rightSet.size && inBoth.length === leftSet.size;

    return { inBoth, onlyLeft, onlyRight, exactSame };
  }

  function collectSideData(sideRoot) {
    const plans = [];
    const cards = qsa('.plan-card', sideRoot);
    for (const card of cards) {
      const planName = qs('.plan-name-input', card).value.trim();
      if (!planName) continue;
      const sections = {};
      for (const key of SECTION_KEYS) {
        const ta = qs(`textarea[data-section="${key}"]`, card);
        sections[key] = parseNames(ta.value);
      }
      plans.push({ planName, sections });
    }
    return plans;
  }

  function buildResultUI(results) {
    const container = createElement('div');
    const title = createElement('h3', { class: 'result-header', text: 'Comparison Results' });
    container.appendChild(title);

    if (results.length === 0) {
      container.appendChild(createElement('p', { class: 'empty', text: 'No plans to compare. Add plans and click Compare.' }));
      return container;
    }

    for (const plan of results) {
      const wrap = createElement('div', { class: 'result-plan' });
      const header = createElement('div', { class: 'result-header' }, [
        createElement('h3', { text: plan.planName }),
        createElement('span', { class: 'badge', text: `Sections: ${SECTION_KEYS.length}` })
      ]);
      wrap.appendChild(header);

      const grid = createElement('div', { class: 'result-grid' });
      for (const key of SECTION_KEYS) {
        const sec = plan.sections[key];
        const sectionBox = createElement('div', { class: 'result-section' });
        sectionBox.appendChild(createElement('h4', { text: SECTION_TITLES[key] }));

        const badgeClass = sec.exactSame ? 'ok' : (sec.inBoth.length ? 'warn' : (sec.onlyLeft.length || sec.onlyRight.length ? 'danger' : ''));
        const stateText = sec.exactSame ? 'Exact match' : (
          sec.inBoth.length ? 'Partial overlap' : (sec.onlyLeft.length || sec.onlyRight.length ? 'No overlap' : 'Empty on both')
        );
        sectionBox.appendChild(createElement('div', { class: `match-state badge ${badgeClass}`, text: stateText }));

        const list = createElement('div', { class: 'result-list' });
        if (sec.inBoth.length === 0 && sec.onlyLeft.length === 0 && sec.onlyRight.length === 0) {
          list.appendChild(createElement('div', { class: 'empty', text: 'No names' }));
        } else {
          let hasUniqueEntries = false;
          if (sec.onlyLeft.length) {
            const left = createElement('div');
            left.appendChild(createElement('div', { class: 'pill only-onekonnect', text: 'Only OneKonnect' }));
            left.appendChild(createElement('ol', { class: 'names list' }, sec.onlyLeft.map(name =>
              createElement('li', { text: name })
            )));
            list.appendChild(left);
            hasUniqueEntries = true;
          }
          if (sec.onlyRight.length) {
            const right = createElement('div');
            right.appendChild(createElement('div', { class: 'pill only-puzzle', text: 'Only Puzzle' }));
            right.appendChild(createElement('ol', { class: 'names list' }, sec.onlyRight.map(name =>
              createElement('li', { text: name })
            )));
            list.appendChild(right);
            hasUniqueEntries = true;
          }
          if (!hasUniqueEntries && sec.inBoth.length) {
            list.appendChild(createElement('div', { class: 'empty', text: 'All names match in both sources.' }));
          }
        }
        sectionBox.appendChild(list);
        grid.appendChild(sectionBox);
      }
      wrap.appendChild(grid);
      container.appendChild(wrap);
    }
    return container;
  }

  function compareSides(onekonnectPlans, puzzlePlans) {
    const allPlanNames = new Set([
      ...onekonnectPlans.map(p => p.planName.trim()),
      ...puzzlePlans.map(p => p.planName.trim())
    ]);

    const byName = (arr) => {
      const m = new Map();
      for (const p of arr) m.set(p.planName.trim(), p);
      return m;
    };
    const leftMap = byName(onekonnectPlans);
    const rightMap = byName(puzzlePlans);

    const results = [];
    for (const name of Array.from(allPlanNames).sort((a,b)=>a.localeCompare(b))) {
      const left = leftMap.get(name);
      const right = rightMap.get(name);
      const sections = {};
      for (const key of SECTION_KEYS) {
        const leftNames = left ? left.sections[key] : [];
        const rightNames = right ? right.sections[key] : [];
        sections[key] = compareNameSets(leftNames, rightNames);
      }
      results.push({ planName: name, sections });
    }
    return results;
  }

  function formatList(values) {
    return values && values.length ? values.join(', ') : 'None';
  }

  function buildTextReport(results) {
    const lines = [];
    lines.push('Plan Comparator Report');
    const ts = new Date();
    lines.push(`Generated: ${ts.toISOString()}`);
    lines.push('');

    if (!results || results.length === 0) {
      lines.push('No plans to compare.');
      return lines.join('\n');
    }

    for (const plan of results) {
      lines.push(`=== Plan: ${plan.planName} ===`);
      for (const key of SECTION_KEYS) {
        const sec = plan.sections[key];
        const title = SECTION_TITLES[key];
        const state = sec.exactSame ? 'Exact match' : (sec.inBoth.length ? 'Partial overlap' : ((sec.onlyLeft.length || sec.onlyRight.length) ? 'No overlap' : 'Empty on both'));
        lines.push(`- ${title}: ${state}`);
        lines.push(`  In Both: ${formatList(sec.inBoth)}`);
        lines.push(`  Only OneKonnect: ${formatList(sec.onlyLeft)}`);
        lines.push(`  Only Puzzle: ${formatList(sec.onlyRight)}`);
      }
      lines.push('');
    }
    return lines.join('\n');
  }

  function makeSheetName(planName, existingNames) {
    const cleaned = (planName || 'Plan')
      .replace(/[\\/?*\[\]:]/g, ' ')
      .trim()
      .slice(0, 31) || 'Plan';
    let candidate = cleaned;
    let counter = 1;
    while (existingNames.has(candidate)) {
      const suffix = ` (${counter})`;
      const base = cleaned.slice(0, Math.max(0, 31 - suffix.length));
      candidate = `${base}${suffix}`;
      counter += 1;
    }
    existingNames.add(candidate);
    return candidate;
  }

  function buildExcelWorkbook(results) {
    const XLSXLib = window.XLSX;
    const wb = XLSXLib.utils.book_new();

    if (!results || results.length === 0) {
      const sheet = XLSXLib.utils.aoa_to_sheet([
        ['Plan Comparator Report'],
        [`Generated: ${new Date().toISOString()}`],
        [],
        ['No plans to compare.']
      ]);
      XLSXLib.utils.book_append_sheet(wb, sheet, 'Summary');
      return wb;
    }

    const usedNames = new Set();
    for (const plan of results) {
      const rows = [
        ['Reconciled', 'Result', 'Unreconciled', 'Result', 'Customer Only', 'Result', 'Invoice Only', 'Result']
      ];

      const sectionOrder = [
        { key: 'reconciled', title: 'Reconciled' },
        { key: 'unreconciled', title: 'Unreconciled' },
        { key: 'customerOnly', title: 'Customer Only' },
        { key: 'invoiceOnly', title: 'Invoice Only' }
      ];

      const columnIndexes = {
        reconciled: { nameCol: 0, resultCol: 1 },
        unreconciled: { nameCol: 2, resultCol: 3 },
        customerOnly: { nameCol: 4, resultCol: 5 },
        invoiceOnly: { nameCol: 6, resultCol: 7 }
      };

      const perSectionRows = sectionOrder.map(section => {
        const sec = plan.sections[section.key];
        const entries = [];
        if (sec) {
          sec.inBoth.forEach(name => entries.push({ name, result: 'In Both' }));
          sec.onlyLeft.forEach(name => entries.push({ name, result: 'Only OneKonnect' }));
          sec.onlyRight.forEach(name => entries.push({ name, result: 'Only Puzzle' }));
        }
        return {
          key: section.key,
          entries
        };
      });

      const maxRows = Math.max(1, ...perSectionRows.map(s => s.entries.length || 0));
      for (let i = 0; i < maxRows; i += 1) {
        const row = new Array(8).fill('');
        perSectionRows.forEach(section => {
          const { nameCol, resultCol } = columnIndexes[section.key];
          const entry = section.entries[i];
          if (entry) {
            row[nameCol] = entry.name;
            row[resultCol] = entry.result;
          }
        });
        rows.push(row);
      }

      const sheet = XLSXLib.utils.aoa_to_sheet(rows);
      const name = makeSheetName(plan.planName, usedNames);
      XLSXLib.utils.book_append_sheet(wb, sheet, name);
    }
    return wb;
  }

  function downloadBlob(filename, blob) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  }

  let lastComparisonResults = [];

  function addPlanCard(sideRoot) {
    const card = createElement('div', { class: 'plan-card' });

    const header = createElement('div', { class: 'plan-name' });
    const planInput = createElement('input', { class: 'plan-name-input', placeholder: 'Plan name e.g., Dental, Vision' });
    const removeBtn = createElement('button', { class: 'btn', title: 'Remove plan' }, [document.createTextNode('Remove')]);
    removeBtn.addEventListener('click', () => { card.remove(); });
    header.appendChild(planInput);
    header.appendChild(removeBtn);
    card.appendChild(header);

    const row = createElement('div', { class: 'plan-row' });
    for (const key of SECTION_KEYS) {
      const section = createElement('div', { class: 'section' });
      const label = createElement('label', { for: `${key}` }, [document.createTextNode(SECTION_TITLES[key])]);
      const ta = createElement('textarea', { "data-section": key, placeholder: 'One name per line' });
      section.appendChild(label);
      section.appendChild(ta);
      row.appendChild(section);
    }
    card.appendChild(row);
    qs('.plans', sideRoot).appendChild(card);
  }

  const PlanUI = {
    addPlan: function(sideId) {
      addPlanCard(qs(`#${sideId}`));
    }
  };

  window.PlanUI = PlanUI;

  function wire() {
    qs('#add-plan-onekonnect').addEventListener('click', () => PlanUI.addPlan('onekonnect-side'));
    qs('#add-plan-puzzle').addEventListener('click', () => PlanUI.addPlan('puzzle-side'));

    const downloadBtn = qs('#download-btn');

    qs('#clear-all-btn').addEventListener('click', () => {
      qsa('.plan-card').forEach(el => el.remove());
      qs('#results').innerHTML = '';
    });

    qs('#compare-btn').addEventListener('click', () => {
      const oneSide = collectSideData(qs('#onekonnect-side'));
      const otherSide = collectSideData(qs('#puzzle-side'));
      const comparison = compareSides(oneSide, otherSide);
      lastComparisonResults = comparison;
      const resultsRoot = qs('#results');
      resultsRoot.innerHTML = '';
      resultsRoot.appendChild(buildResultUI(comparison));
      resultsRoot.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });

    downloadBtn.addEventListener('click', async () => {
      if (downloadBtn.dataset.busy === 'true') return;

      const previousLabel = downloadBtn.textContent;
      const setLabel = (label) => { downloadBtn.textContent = label; };

      downloadBtn.dataset.busy = 'true';
      downloadBtn.disabled = true;
      setLabel('Preparing…');

      try {
        const comparison = (lastComparisonResults && lastComparisonResults.length)
          ? lastComparisonResults
          : compareSides(collectSideData(qs('#onekonnect-side')), collectSideData(qs('#puzzle-side')));

        await ensureXLSXLoaded();
        const workbook = buildExcelWorkbook(comparison);
        const wbout = window.XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const stamp = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        const projectNameRaw = getProjectNameValue();
        const projectNameSafe = formatProjectNameForFilename(projectNameRaw) || 'Project';
        const dateStr = `${stamp.getFullYear()}-${pad(stamp.getMonth() + 1)}-${pad(stamp.getDate())}`;
        const fname = `${projectNameSafe}_Audit_${dateStr}.xlsx`;
        downloadBlob(fname, blob);
        setLabel('Download Ready');
        setTimeout(() => {
          if (downloadBtn.dataset.busy === 'true') return;
          setLabel(previousLabel);
        }, 1500);
      } catch (err) {
        console.error(err);
        alert('Unable to download Excel report. Please check your connection and try again.');
        setLabel(previousLabel);
      } finally {
        downloadBtn.disabled = false;
        delete downloadBtn.dataset.busy;
        if (downloadBtn.textContent === 'Preparing…') {
          setLabel(previousLabel);
        }
      }
    });
  }

  document.addEventListener('DOMContentLoaded', wire);
})();


