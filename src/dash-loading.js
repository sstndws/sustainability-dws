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
