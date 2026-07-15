// Reveal.js interop module for ExploreTheWorld
// Loaded lazily by RevealJs__Interop via ESM import with cache-busting.

let _dotNetRef = null;
let _reveal = null;
let _autoPlayTimer = null;
let _resizeObserver = null;
let _notesPosition = 'hidden'; // 'hidden' | 'right' | 'bottom'

// Container elements — resolved from the reveal element during initialization
let _containerEl = null;  // .etw-slide-container — sized by resizeLayout()
let _wrapperJsEl = null;  // .etw-slides-wrapper  — measured for available space

let Reveal = null;
let RevealZoom = null;
let RevealNotes = null;
let RevealSearch = null;
let RevealHighlight = null;
let _revealLoaded = false;
let _revealLoadPromise = null;

// Resolve reveal.js assets relative to this module so the path works in every
// host (web root, BlazorWebView, Oqtane) regardless of the document base href.
const REVEAL_BASE = new URL('../revealjs/', import.meta.url).href.replace(/\/$/, '');

const THEME_LINK_ID = 'revealjs-theme';
const CSS_MARKER = 'data-etw-reveal-css';

const SLIDE_W = 1280;
const SLIDE_H = 720;

// Compute pixel dimensions for the slide container to fill the wrapper while
// preserving the 1280×720 aspect ratio, then let reveal.js layout within it.
function resizeLayout() {
    if (!_containerEl || !_wrapperJsEl) return;
    const aw = _wrapperJsEl.offsetWidth;
    const ah = _wrapperJsEl.offsetHeight;
    if (!aw || !ah) return;

    let w, h;
    if (aw / ah > SLIDE_W / SLIDE_H) {
        h = ah;
        w = Math.floor(h * SLIDE_W / SLIDE_H);
    } else {
        w = aw;
        h = Math.floor(w * SLIDE_H / SLIDE_W);
    }

    _containerEl.style.width  = `${w}px`;
    _containerEl.style.height = `${h}px`;
    _reveal?.layout();
    logDebug(`resizeLayout: ${w}×${h} (available ${aw}×${ah})`);
}

// ── Logging ────────────────────────────────────────────────────────────────
let _logLevel = 3; // 0=none 1=error 2=warn 3=info 4=debug

function logError(msg, ...args) { if (_logLevel >= 1) console.error(`[CountrySlides] ${msg}`, ...args); }
function logInfo(msg, ...args)  { if (_logLevel >= 3) console.info(`[CountrySlides] ${msg}`, ...args); }
function logDebug(msg, ...args) { if (_logLevel >= 4) console.debug(`[CountrySlides] ${msg}`, ...args); }

export function setLogLevel(level) { _logLevel = level; }

// ── Reveal.js loader ───────────────────────────────────────────────────────
async function loadRevealJs() {
    if (_revealLoaded) return true;
    if (_revealLoadPromise) return _revealLoadPromise;

    logInfo('Loading Reveal.js and plugins…');
    _revealLoadPromise = (async () => {
        try {
            const [rm, zm, nm, sm, hm] = await Promise.all([
                import(`${REVEAL_BASE}/reveal.esm.js`),
                import(`${REVEAL_BASE}/plugin/zoom/zoom.esm.js`),
                import(`${REVEAL_BASE}/plugin/notes/notes.esm.js`),
                import(`${REVEAL_BASE}/plugin/search/search.esm.js`),
                import(`${REVEAL_BASE}/plugin/highlight/highlight.esm.js`)
            ]);
            Reveal = rm.default;
            RevealZoom = zm.default;
            RevealNotes = nm.default;
            RevealSearch = sm.default;
            RevealHighlight = hm.default;
            _revealLoaded = true;
            logInfo(`Reveal.js loaded from ${REVEAL_BASE}`);
            return true;
        } catch (e) {
            logError(`Failed to load Reveal.js from ${REVEAL_BASE}`, e);
            _revealLoadPromise = null;
            return false;
        }
    })();
    return _revealLoadPromise;
}

// ── Stylesheet injection ───────────────────────────────────────────────────
// The reveal.js stylesheets are injected here (not via Blazor <HeadContent>)
// so the slides work in hosts without a HeadOutlet (BlazorWebView, Oqtane).
function _appendCssLink(href, id) {
    const link = document.createElement('link');
    if (id) link.id = id;
    link.rel = 'stylesheet';
    link.href = href;
    link.setAttribute(CSS_MARKER, '');
    const loaded = new Promise(resolve => {
        link.onload = () => resolve();
        link.onerror = () => { logError(`Failed to load stylesheet ${href}`); resolve(); };
        setTimeout(resolve, 3000); // never block initialization on a hung request
    });
    document.head.appendChild(link);
    return loaded;
}

function themeCssUrl(theme) {
    return `${REVEAL_BASE}/theme/${theme}.css`;
}

async function ensureRevealCss(theme) {
    const pending = [];
    if (!document.head.querySelector(`link[${CSS_MARKER}]:not(#${THEME_LINK_ID})`)) {
        pending.push(_appendCssLink(`${REVEAL_BASE}/reset.css`));
        pending.push(_appendCssLink(`${REVEAL_BASE}/reveal.css`));
        pending.push(_appendCssLink(`${REVEAL_BASE}/plugin/highlight/monokai.css`));
    }
    if (!document.getElementById(THEME_LINK_ID)) {
        pending.push(_appendCssLink(themeCssUrl(theme), THEME_LINK_ID));
    }
    await Promise.all(pending);
}

function removeRevealCss() {
    document.head.querySelectorAll(`link[${CSS_MARKER}]`).forEach(link => link.remove());
}

// ── Initialise ─────────────────────────────────────────────────────────────
export async function initialize(containerEl, dotNetRef, transition, theme, logLevel) {
    _logLevel = logLevel ?? 3;
    logInfo(`initialize — transition: ${transition}, theme: ${theme}`);

    _dotNetRef = dotNetRef;

    if (!containerEl) {
        logError('container element is null');
        return false;
    }

    if (_reveal) {
        logInfo('Already initialised — re-layout only');
        resizeLayout();
        return true;
    }

    await ensureRevealCss(theme ?? 'black');

    if (!await loadRevealJs()) return false;

    try {
        _reveal = new Reveal(containerEl, {
            width: 1280,
            height: 720,
            margin: 0.04,
            minScale: 0.1,
            maxScale: 2.0,
            controls: true,
            controlsTutorial: true,
            progress: true,
            slideNumber: 'c/t',
            hash: false,
            history: false,
            keyboard: true,
            overview: true,
            center: true,
            touch: true,
            loop: false,
            mouseWheel: true,
            embedded: true,
            help: true,
            showNotes: false,
            autoPlayMedia: false,
            transition: transition ?? 'slide',
            transitionSpeed: 'default',
            backgroundTransition: 'fade',
            viewDistance: 3,
            plugins: [RevealZoom, RevealNotes, RevealSearch, RevealHighlight]
        });

        // Resolve container refs — .etw-slide-container is the direct parent,
        // .etw-slides-wrapper is two levels up and is what we measure for space.
        _containerEl  = containerEl.parentElement;
        _wrapperJsEl  = _containerEl?.parentElement;

        // Set pixel dimensions BEFORE initialize() so that reveal.js reads
        // correct offsetWidth/offsetHeight during its own init; _reveal.layout()
        // inside resizeLayout() is a no-op at this stage (not initialized yet).
        resizeLayout();

        await _reveal.initialize();
        logInfo('Reveal.js initialised');

        // Re-layout now that reveal.js is fully initialised with correct sizes.
        resizeLayout();

        // Notify Blazor of initial state
        _notifySlideChanged();

        // Events
        _reveal.on('slidechanged', _notifySlideChanged);
        _reveal.on('overviewshown',  () => _dotNetRef?.invokeMethodAsync('OnOverviewChanged', true));
        _reveal.on('overviewhidden', () => _dotNetRef?.invokeMethodAsync('OnOverviewChanged', false));

        // Observe wrapper for size changes and re-apply pixel dimensions each time.
        if (_wrapperJsEl) {
            _resizeObserver = new ResizeObserver(() => resizeLayout());
            _resizeObserver.observe(_wrapperJsEl);
        }

        document.addEventListener('fullscreenchange', _onFullscreenChange);
        logInfo('Reveal.js fully wired');
        return true;
    } catch (e) {
        logError('Error during initialisation', e);
        _reveal = null;
        return false;
    }
}

function _notifySlideChanged() {
    if (!_dotNetRef || !_reveal) return;
    const state = _reveal.getState();
    const total = _reveal.getTotalSlides();
    _dotNetRef.invokeMethodAsync('OnSlideChanged', state.indexh, state.indexv ?? 0, total);
}

function _onFullscreenChange() {
    const isFs = !!document.fullscreenElement;
    _dotNetRef?.invokeMethodAsync('OnFullscreenChanged', isFs);
    // Re-compute pixel dimensions after fullscreen toggle (browser needs a frame to settle)
    setTimeout(() => resizeLayout(), 150);
}

// ── Navigation ─────────────────────────────────────────────────────────────
export function navigateNext() { _reveal?.next(); }
export function navigatePrev() { _reveal?.prev(); }
export function navigateRight() { _reveal?.right(); }
export function navigateLeft() { _reveal?.left(); }
export function navigateDown() { _reveal?.down(); }
export function navigateUp() { _reveal?.up(); }

export function navigateToSlide(h, v) {
    logDebug(`navigateToSlide(${h}, ${v})`);
    _reveal?.slide(h, v ?? 0, -1);
}

export function toggleOverview() { _reveal?.toggleOverview(); }
export function togglePause() { _reveal?.togglePause(); }

// ── Configuration ──────────────────────────────────────────────────────────
export function setTransition(transition) {
    logDebug(`setTransition: ${transition}`);
    _reveal?.configure({ transition });
}

export function setTheme(theme) {
    logDebug(`setTheme: ${theme}`);
    const link = document.getElementById(THEME_LINK_ID);
    if (link) {
        link.href = themeCssUrl(theme);
    } else {
        _appendCssLink(themeCssUrl(theme), THEME_LINK_ID);
    }
}

export function setSlideNumber(show) {
    _reveal?.configure({ slideNumber: show ? 'c/t' : false });
}

// ── Auto-play ──────────────────────────────────────────────────────────────
export function startAutoPlay(intervalMs) {
    stopAutoPlay();
    logInfo(`startAutoPlay: ${intervalMs}ms`);
    _autoPlayTimer = setInterval(() => {
        if (_reveal) {
            const hasNext = _reveal.availableRoutes().right || _reveal.availableRoutes().down;
            if (hasNext) {
                _reveal.next();
            } else {
                _reveal.slide(0, 0, -1); // wrap to beginning
            }
        }
    }, intervalMs);
}

export function stopAutoPlay() {
    if (_autoPlayTimer) {
        clearInterval(_autoPlayTimer);
        _autoPlayTimer = null;
        logInfo('stopAutoPlay');
    }
}

// ── Fullscreen ─────────────────────────────────────────────────────────────
export async function requestFullscreen(containerEl) {
    const target = containerEl ?? document.documentElement;
    if (!document.fullscreenElement) {
        await target.requestFullscreen?.();
    } else {
        await document.exitFullscreen?.();
    }
}

export function isFullscreen() {
    return !!document.fullscreenElement;
}

// ── Notes position ─────────────────────────────────────────────────────────
export function setNotesPosition(position) {
    if (!_reveal) return;
    _notesPosition = position;
    const el = _reveal.getRevealElement();
    const wrapper = el.closest('.etw-slides-wrapper');

    el.classList.remove('etw-notes-bottom');

    if (position === 'hidden') {
        _reveal.configure({ showNotes: false });
        if (wrapper) wrapper.style.overflow = 'hidden';
    } else if (position === 'right') {
        _reveal.configure({ showNotes: true });
        if (wrapper) wrapper.style.overflow = 'visible';
    } else if (position === 'bottom') {
        _reveal.configure({ showNotes: true });
        el.classList.add('etw-notes-bottom');
        if (wrapper) wrapper.style.overflow = 'visible';
    }
    resizeLayout();
}

export function getNotesPosition() { return _notesPosition; }

// ── Search ─────────────────────────────────────────────────────────────────
export function toggleSearch() {
    if (!_reveal) return;
    const plugin = _reveal.getPlugin('search');
    if (!plugin) { logError('Search plugin not loaded'); return; }

    const el = _reveal.getRevealElement();
    const searchBox = el.querySelector('.searchbox');

    if (searchBox && searchBox.style.display !== 'none' && searchBox.style.display !== '') {
        // Close: hide the box (plugin exposes no close(), target the DOM directly)
        searchBox.style.display = 'none';
    } else {
        plugin.open();
    }
}

// ── State ──────────────────────────────────────────────────────────────────
export function getState() {
    if (!_reveal) return null;
    const s = _reveal.getState();
    return { h: s.indexh, v: s.indexv ?? 0, total: _reveal.getTotalSlides() };
}

export function layout() {
    resizeLayout();
}

// ── Cleanup ────────────────────────────────────────────────────────────────
export function destroy() {
    logInfo('destroy');
    stopAutoPlay();
    if (_resizeObserver) {
        _resizeObserver.disconnect();
        _resizeObserver = null;
    }
    document.removeEventListener('fullscreenchange', _onFullscreenChange);
    if (_reveal) {
        try { _reveal.destroy(); } catch { }
        _reveal = null;
    }
    removeRevealCss();
    _dotNetRef = null;
    _containerEl = null;
    _wrapperJsEl = null;
}
