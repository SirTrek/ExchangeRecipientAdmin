// Type-to-confirm guard for destructive actions.
//
// Any form carrying data-expect is wired up: its [data-confirm-input] must be typed
// to match that value before the [data-confirm-button] enables, and submitting raises
// a final confirm() dialog.
//
// The expected value is read from a data- attribute rather than interpolated into this
// script. A quote in the value would otherwise be a syntax error that kills the whole
// block - which previously left the button permanently disabled - and HTML entities
// are not decoded inside <script>, so an encoded value would never compare equal.
//
// This is a convenience only. Every server handler re-checks the confirmation text
// independently, because anything in the browser can be bypassed.
(function () {
  function wire(form) {
    var expected = (form.dataset.expect || '').trim().toLowerCase();
    var input = form.querySelector('[data-confirm-input]');
    var button = form.querySelector('[data-confirm-button]');
    if (!input || !button) { return; }

    var message = form.dataset.confirmMessage || 'This cannot be undone. Continue?';

    function matches() {
      return expected.length > 0 && input.value.trim().toLowerCase() === expected;
    }

    input.addEventListener('input', function () {
      button.disabled = !matches();
    });

    form.addEventListener('submit', function (e) {
      if (!matches() || !confirm(message)) {
        e.preventDefault();
      }
    });
  }

  function init() {
    var forms = document.querySelectorAll('form[data-expect]');
    Array.prototype.forEach.call(forms, wire);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
