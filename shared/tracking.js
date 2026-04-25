/* ==========================================================================
   Rebuilt Training — page view tracking
   Sends a single 'PV' event per page load to the Apps Script endpoint,
   keyed off the email captured at the gate (localStorage 'rebuilt_email').
   No-ops when no email is stored. Failures are swallowed silently so
   tracking never blocks the page.
   ========================================================================== */

(function () {
  'use strict';

  var ENDPOINT = 'https://script.google.com/macros/s/AKfycbyRSuUKBjwlPZiY8E7IZZUEQ8u-uwwuTXZnQ0nTPqtfk-bu3zyAa6RKZYw7687arJVs/exec';
  var sent = false;

  function trackPageView() {
    if (sent) return;
    var email = (localStorage.getItem('rebuilt_email') || '').toLowerCase().trim();
    if (!email) return;
    sent = true;

    var payload = JSON.stringify({
      type:     'PV',
      email:    email,
      role:     localStorage.getItem('rebuilt_role') || '',
      page:     (location.pathname.split('/').pop() || location.pathname || '').toLowerCase(),
      title:    document.title || '',
      referrer: document.referrer || '',
      ts:       Date.now()
    });

    try {
      fetch(ENDPOINT, {
        method:    'POST',
        mode:      'no-cors',
        headers:   { 'Content-Type': 'application/json' },
        body:      payload,
        keepalive: true
      }).catch(function () { /* swallow */ });
    } catch (e) { /* swallow */ }
  }

  // Public hook so the gate page can fire tracking right after login.
  window.rebuiltTrackPageView = trackPageView;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', trackPageView);
  } else {
    trackPageView();
  }
})();
