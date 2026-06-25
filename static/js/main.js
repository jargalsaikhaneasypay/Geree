/* Contract Registry — Main JS */

// ----------------------------------------------------------------
// Current date in topbar
// ----------------------------------------------------------------
(function () {
  const el = document.getElementById('currentDate');
  if (!el) return;
  const now = new Date();
  const opts = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' };
  el.textContent = now.toLocaleDateString('mn-MN', opts);
})();

// ----------------------------------------------------------------
// Auto-dismiss toasts after 4 seconds
// ----------------------------------------------------------------
(function () {
  const toasts = document.querySelectorAll('.toast');
  toasts.forEach(function (t) {
    setTimeout(function () {
      t.style.transition = 'opacity .4s, transform .4s';
      t.style.opacity = '0';
      t.style.transform = 'translateX(60px)';
      setTimeout(function () { t.remove(); }, 400);
    }, 4000);
  });
})();

// ----------------------------------------------------------------
// Delete confirmation modal
// ----------------------------------------------------------------
function openDeleteModal(id) {
  document.getElementById('deleteForm').action = '/delete/' + id;
  document.getElementById('deleteModal').classList.add('open');
}
function closeDeleteModal() {
  document.getElementById('deleteModal').classList.remove('open');
}
// Close on overlay click
document.getElementById('deleteModal').addEventListener('click', function (e) {
  if (e.target === this) closeDeleteModal();
});

// ----------------------------------------------------------------
// Autocomplete for text inputs
// ----------------------------------------------------------------
(function () {
  const inputs = document.querySelectorAll('.autocomplete-input');

  inputs.forEach(function (input) {
    const field   = input.dataset.field;
    const listId  = 'ac_' + field;
    const list    = document.getElementById(listId);
    if (!list) return;

    let timer = null;

    input.addEventListener('input', function () {
      clearTimeout(timer);
      const q = this.value.trim();
      if (q.length < 1) { closeList(list); return; }
      timer = setTimeout(function () { fetchSuggestions(field, q, list, input); }, 220);
    });

    input.addEventListener('keydown', function (e) {
      const items = list.querySelectorAll('li');
      if (!items.length) return;
      let idx = Array.from(items).findIndex(i => i.classList.contains('ac-active'));
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        idx = (idx + 1) % items.length;
        setActive(items, idx);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        idx = (idx - 1 + items.length) % items.length;
        setActive(items, idx);
      } else if (e.key === 'Enter') {
        if (idx >= 0) { e.preventDefault(); input.value = items[idx].textContent; closeList(list); }
      } else if (e.key === 'Escape') {
        closeList(list);
      }
    });

    input.addEventListener('blur', function () {
      setTimeout(function () { closeList(list); }, 160);
    });
  });

  function fetchSuggestions(field, q, list, input) {
    fetch('/api/autocomplete?field=' + encodeURIComponent(field) + '&q=' + encodeURIComponent(q))
      .then(function (r) { return r.json(); })
      .then(function (data) {
        list.innerHTML = '';
        if (!data.length) { closeList(list); return; }
        data.forEach(function (val) {
          const li = document.createElement('li');
          li.textContent = val;
          li.addEventListener('mousedown', function () {
            input.value = val;
            closeList(list);
          });
          list.appendChild(li);
        });
        list.classList.add('open');
      })
      .catch(function () { closeList(list); });
  }

  function closeList(list) {
    list.classList.remove('open');
    list.innerHTML = '';
  }

  function setActive(items, idx) {
    items.forEach(function (i) { i.classList.remove('ac-active'); });
    if (idx >= 0 && items[idx]) {
      items[idx].classList.add('ac-active');
      items[idx].style.background = 'var(--info-light)';
    }
  }
})();
