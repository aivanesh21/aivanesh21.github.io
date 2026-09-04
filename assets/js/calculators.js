/* Sawant & Kadrekar LLP — calculators.
   All computation happens in the browser; nothing is transmitted.
   Every rate, slab and threshold is read from the config block emitted by the
   build from content/tax-rates.json — no rule is hard-coded in this file. */
(function () {
  'use strict';

  var form = document.querySelector('[data-calc]');
  var cfgEl = document.getElementById('calc-config');
  if (!form || !cfgEl) return;

  var CFG;
  try { CFG = JSON.parse(cfgEl.textContent); } catch (e) { return; }

  /* --- Formatting -------------------------------------------------------- */
  var inr0 = new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 });
  var inr2 = new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', minimumFractionDigits: 2, maximumFractionDigits: 2 });
  var money  = function (n) { return isFinite(n) ? inr0.format(Math.round(n)) : '—'; };
  var money2 = function (n) { return isFinite(n) ? inr2.format(n) : '—'; };
  var pct    = function (n) { return isFinite(n) ? n.toFixed(2) + '%' : '—'; };

  function val(name) {
    var el = form.querySelector('[name="' + name + '"]');
    if (!el) return 0;
    var n = parseFloat(el.value);
    return isFinite(n) && n >= 0 ? n : 0;
  }
  function pick(name) {
    var checked = form.querySelector('[name="' + name + '"]:checked');
    if (checked) return checked.value;
    var el = form.querySelector('[name="' + name + '"]');
    return el ? el.value : '';
  }
  function dateVal(name) {
    var el = form.querySelector('[name="' + name + '"]');
    if (!el || !el.value) return null;
    var d = new Date(el.value + 'T00:00:00Z');
    return isNaN(d.getTime()) ? null : d;
  }
  function set(key, text) {
    form.parentElement.querySelectorAll('[data-out="' + key + '"]').forEach(function (el) {
      el.textContent = text;
    });
  }
  function setHTML(key, html) {
    form.parentElement.querySelectorAll('[data-out="' + key + '"]').forEach(function (el) {
      el.innerHTML = html;
    });
  }
  function showRow(key, show) {
    var dd = form.parentElement.querySelector('[data-out="' + key + '"]');
    if (dd && dd.parentElement) dd.parentElement.hidden = !show;
  }

  /* --- Income tax engine -------------------------------------------------
     Shared by the income-tax and regime-comparison tools. */

  function slabsFor(regime, ageGroup) {
    if (regime.slabs) return regime.slabs;
    return (regime.slabsByAgeGroup || {})[ageGroup] || (regime.slabsByAgeGroup || {}).below60 || [];
  }

  function taxOnSlabs(taxable, slabs) {
    var tax = 0, lower = 0, rows = [];
    for (var i = 0; i < slabs.length; i++) {
      var upper = slabs[i].upTo == null ? Infinity : slabs[i].upTo;
      if (taxable > lower) {
        var amount = Math.min(taxable, upper) - lower;
        var t = amount * slabs[i].rate / 100;
        tax += t;
        rows.push({ from: lower, to: upper, rate: slabs[i].rate, amount: amount, tax: t });
      }
      lower = upper;
      if (lower >= taxable) break;
    }
    return { tax: tax, rows: rows };
  }

  /** Surcharge, with marginal relief measured against the band threshold. */
  function surchargeOn(taxable, baseTax, bands, cap, slabs) {
    var rate = 0, threshold = 0, prevRate = 0;
    for (var i = 0; i < bands.length; i++) {
      if (bands[i].rate > cap) continue;
      if (taxable > bands[i].above) { prevRate = rate; rate = bands[i].rate; threshold = bands[i].above; }
    }
    if (!rate) return { amount: 0, rate: 0, relief: 0 };

    var raw = baseTax * rate / 100;
    var taxAtThreshold = taxOnSlabs(threshold, slabs).tax;
    var totalAtThreshold = taxAtThreshold * (1 + prevRate / 100);
    var excessIncome = taxable - threshold;
    var totalHere = baseTax + raw;
    var relief = 0;

    if (totalHere - totalAtThreshold > excessIncome) {
      relief = totalHere - totalAtThreshold - excessIncome;
      relief = Math.min(relief, raw);
    }
    return { amount: raw - relief, rate: rate, relief: relief };
  }

  /** Full computation for one regime. */
  function computeRegime(regime, opts) {
    var it = CFG.incomeTax;
    var slabs = slabsFor(regime, opts.ageGroup);

    var gross = opts.salary + opts.otherIncome;

    var std = 0;
    if (opts.salary > 0) std = Math.min(regime.standardDeductionSalaried || 0, opts.salary);

    var deductions = 0;
    if (regime.allowsDeductions) {
      it.deductionFields.items.forEach(function (d) {
        var v = opts.deductions[d.name] || 0;
        deductions += d.cap ? Math.min(v, d.cap) : v;
      });
    }

    var taxable = Math.max(0, gross - std - deductions);
    var slab = taxOnSlabs(taxable, slabs);
    var baseTax = slab.tax;

    // Rebate
    var rebate = 0;
    var r = regime.rebate87A;
    if (r) {
      if (taxable <= r.incomeThreshold) {
        rebate = Math.min(baseTax, r.maxRebate);
      } else if (r.marginalRelief) {
        // Tax cannot exceed the amount by which income crosses the threshold.
        var excess = taxable - r.incomeThreshold;
        if (baseTax > excess) rebate = baseTax - excess;
      }
    }
    var afterRebate = Math.max(0, baseTax - rebate);

    var sur = surchargeOn(taxable, afterRebate, it.surcharge, regime.surchargeCapPercent, slabs);
    var cess = (afterRebate + sur.amount) * it.cess.rate / 100;
    var total = afterRebate + sur.amount + cess;

    return {
      gross: gross, std: std, deductions: deductions, taxable: taxable,
      slabRows: slab.rows, baseTax: baseTax, rebate: rebate,
      surcharge: sur.amount, surchargeRate: sur.rate, surchargeRelief: sur.relief,
      cess: cess, total: total,
      effective: gross > 0 ? (total / gross) * 100 : 0,
    };
  }

  function readTaxInputs() {
    var it = CFG.incomeTax;
    var deductions = {};
    it.deductionFields.items.forEach(function (d) { deductions[d.name] = val(d.name); });
    return {
      salary: val('salary'),
      otherIncome: val('otherIncome'),
      ageGroup: pick('ageGroup') || 'below60',
      deductions: deductions,
    };
  }

  /* --- income-tax -------------------------------------------------------- */

  function renderIncomeTax() {
    var it = CFG.incomeTax;
    var id = pick('regime') || 'new';
    var regime = it.regimes.filter(function (x) { return x.id === id; })[0] || it.regimes[0];
    var opts = readTaxInputs();
    var res = computeRegime(regime, opts);

    // Deductions only apply under a regime that allows them.
    var group = form.querySelector('[data-group="deductions"]');
    if (group) group.hidden = !regime.allowsDeductions;
    // Age group only changes the old-regime exemption limit.
    var ageField = form.querySelector('#f-ageGroup');
    if (ageField && ageField.closest('.field')) ageField.closest('.field').hidden = !regime.slabsByAgeGroup;

    set('gross', money(res.gross));
    set('stdDeduction', res.std ? '− ' + money(res.std) : '—');
    set('deductions', res.deductions ? '− ' + money(res.deductions) : '—');
    set('taxable', money(res.taxable));
    set('slabTax', money(res.baseTax));
    set('rebate', res.rebate ? '− ' + money(res.rebate) : '—');
    set('surcharge', res.surcharge ? money(res.surcharge) + ' (' + res.surchargeRate + '%)' : '—');
    set('cess', money(res.cess));
    set('total', money(res.total));
    set('effective', pct(res.effective));
    set('monthly', money(res.total / 12));

    showRow('deductions', regime.allowsDeductions);
    showRow('surcharge', res.surcharge > 0);

    var rows = res.slabRows.map(function (r) {
      var label = r.to === Infinity
        ? 'Above ' + money(r.from)
        : money(r.from) + ' – ' + money(r.to);
      return '<tr><th scope="row" style="font-weight:400">' + label + '</th>' +
             '<td class="num">' + r.rate + '%</td>' +
             '<td class="num">' + money(r.amount) + '</td>' +
             '<td class="num">' + money(r.tax) + '</td></tr>';
    });
    if (res.rebate) rows.push('<tr><th scope="row" style="font-weight:400">Less: rebate</th><td class="num">—</td><td class="num">—</td><td class="num">− ' + money(res.rebate) + '</td></tr>');
    if (res.surchargeRelief > 0) rows.push('<tr><th scope="row" style="font-weight:400">Marginal relief on surcharge</th><td class="num">—</td><td class="num">—</td><td class="num">− ' + money(res.surchargeRelief) + '</td></tr>');
    if (!rows.length) rows.push('<tr><td colspan="4" class="muted">No tax arises on this income.</td></tr>');
    setHTML('slabTable', rows.join(''));
  }

  /* --- regime-comparison ------------------------------------------------- */

  function renderComparison() {
    var it = CFG.incomeTax;
    var opts = readTaxInputs();
    var newR = it.regimes.filter(function (r) { return r.id === 'new'; })[0];
    var oldR = it.regimes.filter(function (r) { return r.id === 'old'; })[0];

    var n = computeRegime(newR, opts);
    var o = computeRegime(oldR, opts);

    set('newTotal', money(n.total));
    set('newTaxable', money(n.taxable));
    set('newEffective', pct(n.effective));
    set('oldTotal', money(o.total));
    set('oldTaxable', money(o.taxable));
    set('oldEffective', pct(o.effective));

    var diff = Math.abs(n.total - o.total);
    var better = n.total < o.total ? newR : (o.total < n.total ? oldR : null);

    set('verdictAmount', money(Math.min(n.total, o.total)));
    set('saving', diff < 1 ? 'No difference' : money(diff));
    set('verdictLabel', better
      ? 'The ' + better.label.toLowerCase() + ' costs you less on these figures.'
      : 'Both regimes produce the same liability on these figures.');

    // Break-even: the deduction total at which the two regimes match.
    var probe = { salary: opts.salary, otherIncome: opts.otherIncome, ageGroup: opts.ageGroup, deductions: {} };
    var newFixed = computeRegime(newR, probe).total;
    var lo = 0, hi = Math.max(0, probe.salary + probe.otherIncome), be = null;
    var testAt = function (d) {
      var p = { salary: probe.salary, otherIncome: probe.otherIncome, ageGroup: probe.ageGroup,
                deductions: { otherDeductions: d } };
      return computeRegime(oldR, p).total;
    };
    if (hi > 0 && testAt(hi) < newFixed && testAt(0) > newFixed) {
      for (var i = 0; i < 48; i++) {
        var mid = (lo + hi) / 2;
        if (testAt(mid) > newFixed) lo = mid; else hi = mid;
      }
      be = Math.round((lo + hi) / 2);
    }
    set('breakeven', be != null
      ? 'On this income the two regimes break even at roughly ' + money(be) +
        ' of total deductions. Claim more than that and the old regime wins; less, and the new regime does.'
      : 'On this income the new regime is lower whatever deductions are claimed, so the break-even does not arise.');

    var savingRow = form.parentElement.querySelector('[data-out="saving"]');
    if (savingRow && savingRow.parentElement) {
      savingRow.parentElement.classList.toggle('calc__row--good', diff >= 1);
    }
  }

  /* --- capital-gains ----------------------------------------------------- */

  function monthsBetween(a, b) {
    var m = (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + (b.getUTCMonth() - a.getUTCMonth());
    if (b.getUTCDate() < a.getUTCDate()) m -= 1;
    return m;
  }

  function renderCapitalGains() {
    var cg = CFG.capitalGains;
    var asset = cg.assets.filter(function (a) { return a.id === pick('assetType'); })[0] || cg.assets[0];
    var buy = dateVal('purchaseDate'), sell = dateVal('saleDate');

    var cost = val('purchaseCost'), sale = val('saleValue');
    var improvement = val('improvement'), expenses = val('expenses');
    var used = val('otherLtcgClaimed');

    var months = (buy && sell && sell > buy) ? monthsBetween(buy, sell) : null;
    var isLong = null;
    if (asset.longTermAfterMonths == null) isLong = false;              // always short-term by rule
    else if (months != null) isLong = months >= asset.longTermAfterMonths;

    set('holding', months == null ? 'Enter both dates'
      : Math.floor(months / 12) + 'y ' + (months % 12) + 'm');

    var net = Math.max(0, sale - expenses);
    var gain = net - cost - improvement;

    set('netSale', money(net));
    set('cost', money(cost));
    set('improvementOut', improvement ? money(improvement) : '—');
    showRow('improvementOut', improvement > 0);

    // Indexation is offered only when the index has actually been populated.
    var ciiValues = (cg.costInflationIndex && cg.costInflationIndex.values) || {};
    var haveCII = Object.keys(ciiValues).length > 0;
    showRow('indexedCost', false);
    set('indexedCost', '—');

    if (months == null) {
      set('nature', 'Enter both dates');
      set('gain', money(gain));
      set('exemption', '—');
      set('taxableGain', '—');
      set('rate', '—');
      set('tax', '—');
      return;
    }

    set('nature', isLong ? 'Long term' : 'Short term');
    set('gain', money(gain));

    var exemption = 0, taxableGain = Math.max(0, gain), rate = null, rateNote = '';

    if (isLong) {
      if (asset.longTermExemption) {
        exemption = Math.max(0, Math.min(asset.longTermExemption - used, Math.max(0, gain)));
        taxableGain = Math.max(0, gain - exemption);
      }
      rate = asset.longTermRate;
    } else {
      rate = asset.shortTermRate;
      rateNote = asset.shortTermRateNote || '';
    }

    set('exemption', exemption ? '− ' + money(exemption) : '—');
    showRow('exemption', exemption > 0);
    set('taxableGain', money(taxableGain));

    if (gain <= 0) {
      set('rate', '—');
      set('tax', 'No gain arises');
      return;
    }

    if (rate == null) {
      set('rate', 'Slab rate');
      set('tax', 'Taxed at your slab rate');
      var t = form.parentElement.querySelector('[data-out="tax"]');
      if (t) t.style.fontSize = 'var(--step-1)';
      return;
    }

    set('rate', rate + '%' + (haveCII || !asset.indexationAvailable ? '' : ' (no indexation)'));
    set('tax', money(taxableGain * rate / 100));
  }

  /* --- EMI --------------------------------------------------------------- */

  function renderEMI() {
    var p = val('principal'), annual = val('rate'), years = val('tenure');
    var n = Math.round(years * 12), r = annual / 12 / 100;

    if (!p || !n) {
      ['emi', 'principalOut', 'totalInterest', 'totalPayment', 'interestShare'].forEach(function (k) { set(k, '—'); });
      setHTML('amortTable', '<tr><td colspan="4" class="muted">Enter a loan amount and tenure.</td></tr>');
      return;
    }

    var emi = r === 0 ? p / n : (p * r * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1);
    var totalPayment = emi * n;
    var totalInterest = totalPayment - p;

    set('emi', money(emi));
    set('principalOut', money(p));
    set('totalInterest', money(totalInterest));
    set('totalPayment', money(totalPayment));
    set('interestShare', pct(totalPayment > 0 ? (totalInterest / totalPayment) * 100 : 0));

    var balance = p, rows = [], yearP = 0, yearI = 0;
    for (var m = 1; m <= n; m++) {
      var interest = balance * r;
      var principalPart = Math.min(emi - interest, balance);
      balance = Math.max(0, balance - principalPart);
      yearP += principalPart; yearI += interest;
      if (m % 12 === 0 || m === n) {
        rows.push('<tr><th scope="row" style="font-weight:400">Year ' + Math.ceil(m / 12) + '</th>' +
          '<td class="num">' + money(yearP) + '</td>' +
          '<td class="num">' + money(yearI) + '</td>' +
          '<td class="num">' + money(balance) + '</td></tr>');
        yearP = 0; yearI = 0;
      }
    }
    setHTML('amortTable', rows.join(''));
  }

  /* --- SIP --------------------------------------------------------------- */

  function renderSIP() {
    var monthly = val('monthly'), annual = val('rate'), years = val('years'), step = val('stepUp');
    var i = annual / 12 / 100;

    if (!monthly || !years) {
      ['future', 'invested', 'returns', 'multiple'].forEach(function (k) { set(k, '—'); });
      setHTML('sipTable', '<tr><td colspan="3" class="muted">Enter an amount and a period.</td></tr>');
      return;
    }

    // Month-by-month so the annual step-up is handled exactly.
    var balance = 0, invested = 0, contribution = monthly, rows = [];
    var months = Math.round(years * 12);
    for (var m = 1; m <= months; m++) {
      if (m > 1 && (m - 1) % 12 === 0 && step > 0) contribution *= (1 + step / 100);
      balance = (balance + contribution) * (1 + i);   // contribution at the start of the month
      invested += contribution;
      if (m % 12 === 0 || m === months) {
        rows.push('<tr><th scope="row" style="font-weight:400">Year ' + Math.ceil(m / 12) + '</th>' +
          '<td class="num">' + money(invested) + '</td>' +
          '<td class="num">' + money(balance) + '</td></tr>');
      }
    }

    set('future', money(balance));
    set('invested', money(invested));
    set('returns', money(balance - invested));
    set('multiple', invested > 0 ? (balance / invested).toFixed(2) + '×' : '—');
    setHTML('sipTable', rows.join(''));
  }

  /* --- GST --------------------------------------------------------------- */

  function renderGST() {
    var amount = val('amount');
    var rate = parseFloat(pick('rate')) || 0;
    var adding = (pick('direction') || 'add') === 'add';
    var intra = (pick('supply') || 'intra') === 'intra';

    var base = adding ? amount : amount / (1 + rate / 100);
    var gst = adding ? base * rate / 100 : amount - base;
    var total = base + gst;

    set('base', money2(base));
    set('gstAmount', money2(gst));
    set('cgst', intra ? money2(gst / 2) : '—');
    set('sgst', intra ? money2(gst / 2) : '—');
    set('igst', intra ? '—' : money2(gst));
    set('total', money2(total));

    showRow('cgst', intra);
    showRow('sgst', intra);
    showRow('igst', !intra);
  }

  /* --- HRA --------------------------------------------------------------- */

  function renderHRA() {
    var h = CFG.hra;
    var salary = val('basic'), received = val('hraReceived'), rent = val('rentPaid');
    var metro = (pick('city') || 'metro') === 'metro';
    var sharePct = metro ? h.metroPercent : h.nonMetroPercent;

    var l1 = received;
    var l2 = Math.max(0, rent - salary * (h.salaryPercentForRent / 100));
    var l3 = salary * sharePct / 100;

    var exempt = Math.max(0, Math.min(l1, l2, l3));
    var taxable = Math.max(0, received - exempt);

    set('limit1', money(l1));
    set('limit2', money(l2));
    set('limit3', money(l3) + ' (' + sharePct + '%)');
    set('exempt', money(exempt));
    set('taxable', money(taxable));

    var which = 'Actual HRA received';
    if (l2 <= l1 && l2 <= l3) which = 'Rent less 10% of salary';
    else if (l3 <= l1 && l3 <= l2) which = sharePct + '% of salary';
    set('binding', rent === 0 && received === 0 ? '—' : which);
  }

  /* --- Wire up ----------------------------------------------------------- */
  var RENDER = {
    'income-tax': renderIncomeTax,
    'regime-comparison': renderComparison,
    'capital-gains': renderCapitalGains,
    'emi': renderEMI,
    'sip': renderSIP,
    'gst': renderGST,
    'hra': renderHRA,
  };

  var render = RENDER[form.dataset.calc];
  if (!render) return;

  var raf = null;
  var schedule = function () {
    if (raf) cancelAnimationFrame(raf);
    raf = requestAnimationFrame(function () { raf = null; try { render(); } catch (e) { /* keep the page usable */ } });
  };

  form.addEventListener('input', schedule);
  form.addEventListener('change', schedule);
  form.addEventListener('submit', function (e) { e.preventDefault(); schedule(); });
  render();
})();
