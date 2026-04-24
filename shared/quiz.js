/* ==========================================================================
   Rebuilt Knowledge Check — shared quiz engine
   Used by all 10 knowledge-check pages. Reads config from window.QUIZ_CONFIG.
   ========================================================================== */

(function () {
  'use strict';

  const CONFIG       = window.QUIZ_CONFIG || {};
  const HUB_SCRIPT_URL = CONFIG.scriptUrl;
  const MATERIAL_URL = CONFIG.materialUrl || '';
  const MODULE_NAME  = CONFIG.moduleName || '';
  const MODULE_ID    = CONFIG.moduleId   || '';
  const PASS_PCT     = CONFIG.passThreshold != null ? CONFIG.passThreshold : 0.70;
  const QUESTIONS    = CONFIG.questions || [];
  const RESULT_COPY  = CONFIG.resultCopy || {
    pass: "Your score has been recorded.",
    fail: "You scored below the passing threshold. Review the material and give it another shot."
  };

  let userName = '', userEmail = '', currentQ = 0, score = 0, answered = false, startTime = null;

  function $(id) { return document.getElementById(id); }

  function validateEmail(email) {
    return /^[^\s@]+@rebuilt\.com$/i.test(email.trim());
  }

  function showScreen(id) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    $(id).classList.add('active');
    window.scrollTo(0, 0);
  }

  function prefillEmail() {
    const saved = localStorage.getItem('rebuilt_email');
    if (saved && $('inputEmail')) $('inputEmail').value = saved;
  }

  function wireMaterialLinks() {
    if (!MATERIAL_URL) return;
    ['materialLink', 'footerMaterialLink'].forEach(id => {
      const el = $(id);
      if (el) { el.href = MATERIAL_URL; el.classList.remove('no-link'); }
    });
  }

  function startQuiz() {
    userEmail = $('inputEmail').value.trim();
    userName  = userEmail.split('@')[0].replace(/[._]/g, ' ').replace(/(^\w|\s\w)/g, m => m.toUpperCase());
    $('errorEmail').style.display = 'none';
    if (!validateEmail(userEmail)) {
      $('errorEmail').style.display = 'block';
      return;
    }
    startTime = Date.now();
    localStorage.setItem('rebuilt_email', userEmail.toLowerCase());
    showScreen('screenQuiz');
    renderQuestion();
  }

  function renderQuestion() {
    const q = QUESTIONS[currentQ];
    answered = false;
    $('qNumber').textContent = 'Question ' + (currentQ + 1);
    $('qText').textContent = q.text;
    $('feedbackBox').style.display = 'none';
    $('btnNext').style.display = 'none';
    const pct = Math.round((currentQ / QUESTIONS.length) * 100);
    $('progressLabel').textContent = 'Question ' + (currentQ + 1) + ' of ' + QUESTIONS.length;
    $('progressPct').textContent = pct + '%';
    $('progressFill').style.width = pct + '%';
    const wrap = $('optionsWrap');
    wrap.innerHTML = '';
    const letters = ['A', 'B', 'C', 'D', 'E'];
    q.options.forEach((opt, i) => {
      const div = document.createElement('div');
      div.className = 'option';
      const letter = document.createElement('div');
      letter.className = 'option-letter';
      letter.textContent = letters[i];
      const text = document.createElement('div');
      text.className = 'option-text';
      text.textContent = opt;
      div.appendChild(letter);
      div.appendChild(text);
      div.addEventListener('click', () => selectOption(i));
      wrap.appendChild(div);
    });
  }

  function selectOption(idx) {
    if (answered) return;
    answered = true;
    const q = QUESTIONS[currentQ];
    const opts = document.querySelectorAll('.option');
    opts.forEach(o => o.classList.add('locked'));
    const fb = $('feedbackBox');
    const btn = $('btnNext');
    if (idx === q.correct) {
      score++;
      opts[idx].classList.add('correct');
      fb.className = 'feedback-box correct-fb';
      fb.innerHTML = '<span class="feedback-icon">✓</span> Correct! ' + q.explanation;
    } else {
      opts[idx].classList.add('incorrect');
      opts[q.correct].classList.add('reveal-correct');
      fb.className = 'feedback-box incorrect-fb';
      fb.innerHTML = '<span class="feedback-icon">✗</span> Not quite. ' + q.explanation;
    }
    fb.style.display = 'block';
    btn.style.display = 'block';
    btn.textContent = (currentQ < QUESTIONS.length - 1) ? 'Next Question →' : 'See Results →';
  }

  function nextQuestion() {
    currentQ++;
    if (currentQ < QUESTIONS.length) renderQuestion();
    else showResults();
  }

  function showResults() {
    const pct     = score / QUESTIONS.length;
    const pass    = pct >= PASS_PCT;
    const elapsed = Math.round((Date.now() - startTime) / 1000);
    $('progressFill').style.width = '100%';
    $('progressPct').textContent  = '100%';
    $('progressLabel').textContent = 'Complete';
    showScreen('screenResults');
    $('resultBadge').innerHTML  = pass ? '🏆' : '📝';
    $('resultBadge').className  = 'result-badge ' + (pass ? 'pass' : 'fail');
    $('resultStatus').textContent = pass ? 'Knowledge Check Passed!' : 'Not Passing — Review & Retry';
    $('resultStatus').className = 'result-status ' + (pass ? 'pass' : 'fail');
    $('resultScore').textContent = Math.round(pct * 100) + '%';
    $('resultScoreLabel').textContent = score + ' of ' + QUESTIONS.length + ' correct';
    $('resultDetail').innerHTML = 'Great work, <span class="result-name">' + userName + '</span>! ' +
      (pass ? RESULT_COPY.pass : RESULT_COPY.fail);
    saveScore(elapsed);
  }

  function saveScore(/* elapsed */) {
    const saveEl  = $('saveStatus');
    const overlay = $('savingOverlay');
    const scorePct = Math.round((score / QUESTIONS.length) * 100);
    const email = (localStorage.getItem('rebuilt_email') || userEmail || '').toLowerCase().trim();
    const role  = localStorage.getItem('rebuilt_role') || '';

    saveEl.textContent = 'Saving your score…';
    saveEl.className = 'save-status saving';
    overlay.classList.add('show');

    const payload = JSON.stringify({
      email:       email,
      module_id:   MODULE_ID,
      module_name: MODULE_NAME,
      type:        'KC',
      score_pct:   scorePct,
      role:        role
    });

    const timeout = setTimeout(() => {
      overlay.classList.remove('show');
      saveEl.textContent = '⚠ Save timed out — check your connection.';
      saveEl.className = 'save-status save-error';
    }, 10000);

    fetch(HUB_SCRIPT_URL, {
      method:  'POST',
      mode:    'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body:    payload
    }).then(() => {
      clearTimeout(timeout);
      overlay.classList.remove('show');
      saveEl.textContent = '✓ Score saved to your training record.';
      saveEl.className = 'save-status saved';
    }).catch(() => {
      clearTimeout(timeout);
      overlay.classList.remove('show');
      saveEl.textContent = '✓ Score submitted to your training record.';
      saveEl.className = 'save-status saved';
    });
  }

  // Expose handlers used by inline onclick in HTML
  window.startQuiz    = startQuiz;
  window.nextQuestion = nextQuestion;

  document.addEventListener('DOMContentLoaded', function () {
    prefillEmail();
    wireMaterialLinks();
  });
})();
