(() => {
  'use strict';
  const pop = document.getElementById('kpiTooltip');
  if (!pop) return;
  const triggers = [...document.querySelectorAll('.info-trigger')];
  let current = null;
  function close() {
    pop.hidden = true;
    if (current) current.setAttribute('aria-expanded', 'false');
    current = null;
  }
  function position(btn) {
    const r = btn.getBoundingClientRect();
    const w = pop.offsetWidth, h = pop.offsetHeight;
    const margin = 12, gap = 16;
    // Priorizar el lado derecho o izquierdo para no tapar el icono ni el puntero.
    let x;
    if (r.right + gap + w + margin <= innerWidth) x = r.right + gap;
    else if (r.left - gap - w >= margin) x = r.left - gap - w;
    else x = Math.max(margin, Math.min(r.left, innerWidth - w - margin));
    let y = Math.max(margin, Math.min(r.top, innerHeight - h - margin));
    // En pantallas estrechas colocar debajo o encima del disparador.
    if (x < r.right + gap && x + w > r.left - gap) {
      if (r.bottom + gap + h + margin <= innerHeight) y = r.bottom + gap;
      else if (r.top - gap - h >= margin) y = r.top - gap - h;
    }
    pop.style.left = `${x}px`;
    pop.style.top = `${y}px`;
  }
  function open(btn) {
    if (current && current !== btn) current.setAttribute('aria-expanded', 'false');
    current = btn;
    btn.setAttribute('aria-expanded', 'true');
    pop.textContent = btn.dataset.tooltip || '';
    pop.hidden = false;
    position(btn);
  }
  triggers.forEach(btn => {
    btn.addEventListener('mouseenter', () => {
      if (matchMedia('(hover: hover)').matches) open(btn);
    });
    // No cerrar al salir del icono: el usuario puede leer sin mantener el mouse encima.
    btn.addEventListener('click', event => {
      event.stopPropagation();
      if (current === btn && !pop.hidden) close();
      else open(btn);
    });
    btn.addEventListener('focus', () => open(btn));
  });
  document.addEventListener('pointerdown', event => {
    if (!event.target.closest('.info-trigger')) close();
  });
  document.addEventListener('keydown', event => { if (event.key === 'Escape') close(); });
  window.addEventListener('scroll', close, true);
  window.addEventListener('resize', () => { if (current) position(current); });
})();
