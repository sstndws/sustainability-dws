/**
 * Shared dashboard loading indicator markup.
 * Keep label short — the motion carries the loading state.
 */

export function dashLoadingHtml_(label, opts) {
  const text = String(label == null || label === '' ? 'Loading…' : label).trim() || 'Loading…';
  const inline = !!(opts && opts.inline);
  return ''
    + '<div class="dash-loading' + (inline ? ' dash-loading--inline' : '') + '" role="status" aria-live="polite">'
    + '<div class="dash-loading__visual" aria-hidden="true">'
    + '<span class="dash-loading__ring"></span>'
    + '<span class="dash-loading__dot"></span>'
    + '</div>'
    + '<p class="dash-loading__label">' + escapeLoadingLabel_(text) + '</p>'
    + '</div>';
}

function escapeLoadingLabel_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Fill an existing loading host element with the shared template. */
export function dashMountLoading_(el, label, opts) {
  if (!el) return;
  el.classList.add('dash-loading-host');
  el.innerHTML = dashLoadingHtml_(label, opts);
}

/** Inline busy state for buttons (spinner + label). */
export function dashBtnBusyHtml_(label) {
  const text = escapeLoadingLabel_(label == null || label === '' ? 'Loading…' : label);
  return ''
    + '<span class="dash-btn-busy" role="status" aria-live="polite">'
    + '<span class="dash-btn-busy__spin" aria-hidden="true"></span>'
    + '<span class="dash-btn-busy__text">' + text + '</span>'
    + '</span>';
}

export function dashSetButtonBusy_(btn, label, opts) {
  if (!btn) return;
  if (btn.dataset.dashBtnBusyPrev == null) {
    btn.dataset.dashBtnBusyPrev = btn.innerHTML;
  }
  btn.disabled = true;
  btn.classList.add('is-dash-busy');
  const prefix = opts && opts.prefixHtml ? String(opts.prefixHtml) : '';
  btn.innerHTML = prefix + dashBtnBusyHtml_(label);
}

export function dashClearButtonBusy_(btn, fallbackHtml) {
  if (!btn) return;
  btn.disabled = false;
  btn.classList.remove('is-dash-busy');
  const prev = btn.dataset.dashBtnBusyPrev;
  delete btn.dataset.dashBtnBusyPrev;
  if (prev != null) btn.innerHTML = prev;
  else if (fallbackHtml != null) btn.innerHTML = fallbackHtml;
}

/** Toast / banner while export or save is in progress. */
export function dashProgressToastHtml_(message) {
  const text = escapeLoadingLabel_(message);
  return ''
    + '<span class="dash-toast-progress" role="status" aria-live="polite">'
    + '<span class="dash-toast-progress__spin" aria-hidden="true"></span>'
    + '<span class="dash-toast-progress__text">' + text + '</span>'
    + '</span>';
}

export function dashMessageLooksInProgress_(message) {
  const s = String(message || '').trim();
  if (!s) return false;
  return /\b(generating|exporting|preparing|loading|building|saving|please wait|memperbarui|menyimpan)\b/i.test(s);
}
