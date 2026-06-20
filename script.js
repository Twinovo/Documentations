const slides = [...document.querySelectorAll('.slide')];
const currentEl = document.getElementById('currentSlide');
const totalEl = document.getElementById('totalSlides');
const progress = document.getElementById('progressBar');
const label = document.getElementById('sectionLabel');
const prevBtn = document.getElementById('prevBtn');
const nextBtn = document.getElementById('nextBtn');
const overview = document.getElementById('overview');
const overviewGrid = document.getElementById('overviewGrid');
let index = 0;
let locked = false;

totalEl.textContent = String(slides.length).padStart(2, '0');

slides.forEach((slide, i) => {
  const button = document.createElement('button');
  button.className = 'overview-item';
  button.innerHTML = `<b>${String(i + 1).padStart(2, '0')}</b><span>${slide.dataset.section}</span>`;
  button.addEventListener('click', () => { overview.classList.remove('open'); showSlide(i); });
  overviewGrid.appendChild(button);
});

function updateOverview() {
  [...overviewGrid.children].forEach((item, i) => item.classList.toggle('current', i === index));
}

function showSlide(next, direction = null) {
  if (locked || next === index || next < 0 || next >= slides.length) return;
  locked = true;
  const old = slides[index];
  old.classList.remove('active');
  if ((direction || next > index) === true) old.classList.add('exit-left');
  index = next;
  slides[index].classList.add('active');
  currentEl.textContent = String(index + 1).padStart(2, '0');
  progress.style.width = `${((index + 1) / slides.length) * 100}%`;
  label.textContent = slides[index].dataset.section.toUpperCase();
  history.replaceState(null, '', `#${index + 1}`);
  updateOverview();
  setTimeout(() => { old.classList.remove('exit-left'); locked = false; }, 550);
}

function next() { showSlide(index + 1, true); }
function prev() { showSlide(index - 1, false); }
nextBtn.addEventListener('click', next);
prevBtn.addEventListener('click', prev);

document.addEventListener('keydown', e => {
  if (overview.classList.contains('open')) {
    if (e.key === 'Escape') overview.classList.remove('open');
    return;
  }
  if (['ArrowRight', 'ArrowDown', 'PageDown', ' '].includes(e.key)) { e.preventDefault(); next(); }
  if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(e.key)) { e.preventDefault(); prev(); }
  if (e.key === 'Home') showSlide(0);
  if (e.key === 'End') showSlide(slides.length - 1);
  if (e.key.toLowerCase() === 'o') overview.classList.add('open');
});

let touchX = 0;
document.addEventListener('touchstart', e => touchX = e.changedTouches[0].screenX, { passive: true });
document.addEventListener('touchend', e => {
  const delta = e.changedTouches[0].screenX - touchX;
  if (Math.abs(delta) > 55) delta < 0 ? next() : prev();
}, { passive: true });

document.getElementById('overviewBtn').addEventListener('click', () => { overview.classList.add('open'); updateOverview(); });
document.getElementById('closeOverview').addEventListener('click', () => overview.classList.remove('open'));
document.getElementById('restartBtn').addEventListener('click', () => showSlide(0));
document.getElementById('presentBtn').addEventListener('click', async () => {
  try {
    if (!document.fullscreenElement) await document.documentElement.requestFullscreen();
    else await document.exitFullscreen();
  } catch (_) {}
});

document.querySelectorAll('.journey-step').forEach(step => {
  step.addEventListener('click', () => {
    document.querySelectorAll('.journey-step').forEach(s => s.classList.remove('active'));
    step.classList.add('active');
    document.getElementById('journeyNumber').textContent = step.querySelector('b').textContent;
    document.getElementById('journeyText').textContent = step.dataset.detail;
  });
});

const mapStage = document.getElementById('mapStage');
const analysisMaps = [...document.querySelectorAll('.analysis-map')];
const mapCaptions = {
  Energy: 'Consumption intensity by zone',
  Lighting: 'Illumination coverage and fixture reach',
  Comfort: 'Thermal balance and occupancy comfort',
  Traffic: 'Movement routes and high-use transitions'
};
document.querySelectorAll('.viz-toggle').forEach(toggle => {
  toggle.addEventListener('click', () => {
    document.querySelectorAll('.viz-toggle').forEach(t => t.classList.remove('active'));
    toggle.classList.add('active');
    const mode = toggle.dataset.mode;
    document.getElementById('overlayMode').textContent = mode.toUpperCase();
    document.getElementById('mapCaption').textContent = mapCaptions[mode];
    analysisMaps.forEach(map => map.classList.toggle('active', map.dataset.map === mode));
    mapStage.dataset.mode = mode.toLowerCase();
    mapStage.classList.remove('mode-change');
    void mapStage.offsetWidth;
    mapStage.classList.add('mode-change');
  });
});


document.querySelectorAll('.map-hotspot').forEach(hotspot => {
  hotspot.addEventListener('click', () => hotspot.classList.toggle('open'));
});

const hashIndex = Number(location.hash.replace('#', '')) - 1;
if (hashIndex > 0 && hashIndex < slides.length) {
  slides[0].classList.remove('active');
  index = hashIndex;
  slides[index].classList.add('active');
  currentEl.textContent = String(index + 1).padStart(2, '0');
  progress.style.width = `${((index + 1) / slides.length) * 100}%`;
  label.textContent = slides[index].dataset.section.toUpperCase();
}
updateOverview();
