/** Shared date field: always DD/MM/YYYY display + consistent calendar UI (locale-independent). */

const DASH_WEEKDAYS = ['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su'];
const DASH_MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const DASH_CAL_ICON = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>';

let dashActivePopover = null;

function dashPad2(n) {
  return String(n).padStart(2, '0');
}

function dashEscAttr(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;');
}

function dashToIsoDate_(d) {
  return d.getFullYear() + '-' + dashPad2(d.getMonth() + 1) + '-' + dashPad2(d.getDate());
}

export function dashNormalizeToIso(raw) {
  if (raw === undefined || raw === null || raw === '') return '';
  if (raw instanceof Date && !isNaN(raw.getTime())) return dashToIsoDate_(raw);
  const s = String(raw).trim();
  if (!s || s === '—') return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

  const monMatch = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (monMatch) {
    const months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    const mon = months[String(monMatch[2]).toLowerCase().slice(0, 3)];
    if (mon !== undefined) {
      let yr = parseInt(monMatch[3], 10);
      if (yr < 100) yr += 2000;
      const d = new Date(yr, mon, parseInt(monMatch[1], 10));
      if (!isNaN(d.getTime())) return dashToIsoDate_(d);
    }
  }

  const slash = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (slash) return dashDisplayToIso(slash[1] + '/' + slash[2] + '/' + slash[3]);

  const n = parseFloat(String(s).replace(',', '.'));
  if (!isNaN(n) && n > 20000 && n < 100000 && !/\//.test(s) && !/-/.test(s)) {
    const d = new Date((n - 25569) * 86400000);
    if (!isNaN(d.getTime())) return dashToIsoDate_(d);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return dashToIsoDate_(d);
  return '';
}

export function dashIsoToDisplay(iso) {
  if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return '';
  const parts = iso.split('-');
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

export function dashDisplayToIso(text) {
  const m = String(text || '').trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (!m) return '';
  const dd = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  let yy = parseInt(m[3], 10);
  if (m[3].length === 2) yy = 2000 + yy;
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return '';
  const d = new Date(yy, mm - 1, dd);
  if (d.getFullYear() !== yy || d.getMonth() !== mm - 1 || d.getDate() !== dd) return '';
  return yy + '-' + dashPad2(mm) + '-' + dashPad2(dd);
}

/** Compact date field for table cells (no label). */
export function dashDateTableCellHtml(rawValue, opts) {
  opts = opts || {};
  const iso = dashNormalizeToIso(rawValue);
  const display = dashIsoToDisplay(iso);
  const uid = opts.id || ('dash-date-' + Math.random().toString(36).slice(2, 9));
  const aria = opts.label != null ? opts.label : 'Date';

  return '<div class="dash-date-field dash-date-field--table">'
    + '<div class="dash-date-wrap">'
    + '<input type="text" id="' + uid + '-text" class="dash-date-text" placeholder="DD/MM/YYYY" value="' + dashEscAttr(display) + '" inputmode="numeric" autocomplete="off" aria-label="' + dashEscAttr(aria) + '">'
    + '<input type="hidden" class="dash-date-value" value="' + dashEscAttr(iso) + '">'
    + '<button type="button" class="dash-date-trigger" aria-label="Open calendar" title="Choose date">' + DASH_CAL_ICON + '</button>'
    + '<div class="dash-date-popover" hidden></div>'
    + '</div>'
    + '</div>';
}

export function dashDateFieldHtml(fieldName, rawValue, opts) {
  opts = opts || {};
  const iso = dashNormalizeToIso(rawValue);
  const display = dashIsoToDisplay(iso);
  const uid = opts.id || ('dash-date-' + String(fieldName).replace(/[^a-z0-9]+/gi, '-').toLowerCase());
  const label = opts.label != null ? opts.label : fieldName;
  const wrapClass = opts.wrapClass || 'form-field';
  const dataFieldAttr = opts.dataField === false ? '' : ' data-field="' + dashEscAttr(fieldName) + '"';

  return '<div class="' + wrapClass + ' dash-date-field">'
    + '<label for="' + uid + '-text">' + dashEscAttr(label) + '</label>'
    + '<div class="dash-date-wrap">'
    + '<input type="text" id="' + uid + '-text" class="dash-date-text" placeholder="DD/MM/YYYY" value="' + dashEscAttr(display) + '" inputmode="numeric" autocomplete="off" aria-label="' + dashEscAttr(label) + '">'
    + '<input type="hidden" class="dash-date-value"' + dataFieldAttr + ' value="' + dashEscAttr(iso) + '">'
    + '<button type="button" class="dash-date-trigger" aria-label="Open calendar" title="Choose date">' + DASH_CAL_ICON + '</button>'
    + '<div class="dash-date-popover" hidden></div>'
    + '</div>'
    + '</div>';
}

function dashClosePopover() {
  if (dashActivePopover) {
    dashActivePopover.hidden = true;
    dashActivePopover = null;
  }
}

function dashTodayIso_() {
  return dashToIsoDate_(new Date());
}

function dashYearBounds_(centerYear) {
  const now = new Date().getFullYear();
  return {
    min: Math.min(1980, centerYear - 15),
    max: Math.max(now + 5, centerYear + 15),
  };
}

function dashMonthSelectHtml_(selectedMonth) {
  return DASH_MONTHS.map(function(name, idx) {
    return '<option value="' + idx + '"' + (idx === selectedMonth ? ' selected' : '') + '>' + name + '</option>';
  }).join('');
}

function dashYearSelectHtml_(selectedYear) {
  const bounds = dashYearBounds_(selectedYear);
  let html = '';
  for (let yr = bounds.max; yr >= bounds.min; yr--) {
    html += '<option value="' + yr + '"' + (yr === selectedYear ? ' selected' : '') + '>' + yr + '</option>';
  }
  return html;
}

function dashRenderCalendar_(popover, viewDate, selectedIso, onSelect) {
  const y = viewDate.getFullYear();
  const m = viewDate.getMonth();
  const first = new Date(y, m, 1);
  const startOffset = (first.getDay() + 6) % 7;
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const todayIso = dashTodayIso_();

  let html = '<div class="dash-cal-head">'
    + '<div class="dash-cal-pickers">'
    + '<select class="dash-cal-month" aria-label="Month">' + dashMonthSelectHtml_(m) + '</select>'
    + '<select class="dash-cal-year" aria-label="Year">' + dashYearSelectHtml_(y) + '</select>'
    + '</div>'
    + '<div class="dash-cal-nav-row">'
    + '<button type="button" class="dash-cal-nav" data-nav="-1" aria-label="Previous month">&#8249;</button>'
    + '<button type="button" class="dash-cal-today">Today</button>'
    + '<button type="button" class="dash-cal-nav" data-nav="1" aria-label="Next month">&#8250;</button>'
    + '</div>'
    + '</div>'
    + '<div class="dash-cal-weekdays">' + DASH_WEEKDAYS.map(function(w) { return '<span>' + w + '</span>'; }).join('') + '</div>'
    + '<div class="dash-cal-grid">';

  for (let i = 0; i < startOffset; i++) html += '<span class="dash-cal-day empty"></span>';
  for (let day = 1; day <= daysInMonth; day++) {
    const iso = y + '-' + dashPad2(m + 1) + '-' + dashPad2(day);
    let cls = 'dash-cal-day';
    if (iso === selectedIso) cls += ' selected';
    if (iso === todayIso) cls += ' today';
    html += '<button type="button" class="' + cls + '" data-iso="' + iso + '">' + day + '</button>';
  }
  html += '</div>';
  popover.innerHTML = html;

  const monthSel = popover.querySelector('.dash-cal-month');
  const yearSel = popover.querySelector('.dash-cal-year');
  function rerenderFromPickers() {
    const newY = parseInt(yearSel.value, 10);
    const newM = parseInt(monthSel.value, 10);
    dashRenderCalendar_(popover, new Date(newY, newM, 1), selectedIso, onSelect);
  }
  if (monthSel) {
    monthSel.addEventListener('change', function(e) {
      e.stopPropagation();
      rerenderFromPickers();
    });
    monthSel.addEventListener('click', function(e) { e.stopPropagation(); });
  }
  if (yearSel) {
    yearSel.addEventListener('change', function(e) {
      e.stopPropagation();
      rerenderFromPickers();
    });
    yearSel.addEventListener('click', function(e) { e.stopPropagation(); });
  }

  popover.querySelectorAll('[data-nav]').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      dashRenderCalendar_(popover, new Date(y, m + parseInt(btn.dataset.nav, 10), 1), selectedIso, onSelect);
    });
  });
  const todayBtn = popover.querySelector('.dash-cal-today');
  if (todayBtn) {
    todayBtn.addEventListener('click', function(e) {
      e.stopPropagation();
      const today = new Date();
      dashRenderCalendar_(popover, new Date(today.getFullYear(), today.getMonth(), 1), selectedIso, onSelect);
    });
  }
  popover.querySelectorAll('.dash-cal-day[data-iso]').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      onSelect(btn.dataset.iso);
    });
  });
}

function dashFormatDigits_(digits) {
  if (digits.length <= 2) return digits;
  if (digits.length <= 4) return digits.slice(0, 2) + '/' + digits.slice(2);
  return digits.slice(0, 2) + '/' + digits.slice(2, 4) + '/' + digits.slice(4);
}

export function dashDateSyncField(fieldEl) {
  const text = fieldEl.querySelector('.dash-date-text');
  const hidden = fieldEl.querySelector('.dash-date-value');
  if (!text || !hidden) return;
  const trimmed = text.value.trim();
  if (!trimmed) {
    hidden.value = '';
    text.classList.remove('dash-date-invalid');
    return;
  }
  const iso = dashDisplayToIso(trimmed);
  if (iso) {
    hidden.value = iso;
    text.value = dashIsoToDisplay(iso);
    text.classList.remove('dash-date-invalid');
  } else {
    text.classList.add('dash-date-invalid');
  }
}

export function dashDateCollectValues(root) {
  const scope = root || document;
  scope.querySelectorAll('.dash-date-field').forEach(dashDateSyncField);
}

export function dashDateReadIso(root, fieldName) {
  if (!root) return '';
  const fields = root.querySelectorAll('.dash-date-value');
  for (let i = 0; i < fields.length; i++) {
    if (fields[i].dataset.field === fieldName) return fields[i].value || '';
  }
  return '';
}

export function initDashDateFields(root) {
  const scope = root || document;
  scope.querySelectorAll('.dash-date-field').forEach(function(fieldEl) {
    if (fieldEl.dataset.dashDateBound) return;
    fieldEl.dataset.dashDateBound = '1';

    const text = fieldEl.querySelector('.dash-date-text');
    const hidden = fieldEl.querySelector('.dash-date-value');
    const trigger = fieldEl.querySelector('.dash-date-trigger');
    const popover = fieldEl.querySelector('.dash-date-popover');
    if (!text || !hidden || !trigger || !popover) return;

    text.addEventListener('blur', function() { dashDateSyncField(fieldEl); });
    text.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        dashDateSyncField(fieldEl);
        dashClosePopover();
      }
    });
    text.addEventListener('input', function() {
      const digits = text.value.replace(/\D/g, '').slice(0, 8);
      const formatted = dashFormatDigits_(digits);
      if (formatted !== text.value) text.value = formatted;
    });

    trigger.addEventListener('click', function(e) {
      e.stopPropagation();
      if (!popover.hidden && dashActivePopover === popover) {
        dashClosePopover();
        return;
      }
      dashClosePopover();
      dashDateSyncField(fieldEl);
      const iso = hidden.value || dashDisplayToIso(text.value);
      const view = iso ? new Date(iso + 'T12:00:00') : new Date();
      dashRenderCalendar_(popover, view, iso, function(selectedIso) {
        hidden.value = selectedIso;
        text.value = dashIsoToDisplay(selectedIso);
        text.classList.remove('dash-date-invalid');
        dashClosePopover();
      });
      popover.hidden = false;
      dashActivePopover = popover;
    });

    popover.addEventListener('click', function(e) { e.stopPropagation(); });
  });

  if (!initDashDateFields._docBound) {
    initDashDateFields._docBound = true;
    document.addEventListener('click', dashClosePopover);
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') dashClosePopover();
    });
  }
}
