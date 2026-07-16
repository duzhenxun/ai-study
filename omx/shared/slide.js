
(() => {
  const fit = () => {
    const scale = Math.min(innerWidth / 1280, innerHeight / 720);
    document.documentElement.style.setProperty('--scale', String(scale));
  };
  fit();
  addEventListener('resize', fit);
  const actions = { ArrowRight: 'next', ArrowDown: 'next', ' ': 'next', PageDown: 'next', ArrowLeft: 'prev', ArrowUp: 'prev', PageUp: 'prev', Home: 'first', End: 'last', n: 'notes', N: 'notes', f: 'fullscreen', F: 'fullscreen' };
  addEventListener('keydown', (event) => {
    if (!actions[event.key]) return;
    event.preventDefault();
    parent.postMessage({ deckAction: actions[event.key] }, '*');
  });
  let startX = null;
  addEventListener('pointerdown', (event) => { startX = event.clientX; });
  addEventListener('pointerup', (event) => {
    if (startX === null) return;
    const delta = event.clientX - startX;
    startX = null;
    if (Math.abs(delta) > 45) parent.postMessage({ deckAction: delta < 0 ? 'next' : 'prev' }, '*');
  });
})();
