/* ===== Utilities kecil ===== */
function cgsTbodyAppend(node, tbody) { tbody.appendChild(node); }
const rootStyle = document.documentElement.style;

// ===== Sticky navbar shadow toggler =====
const navbarWrapper = document.getElementById('navbarWrapper');
const setStickinessShadow = () => {
    if (window.scrollY > 2) navbarWrapper.classList.add('is-stuck');
    else navbarWrapper.classList.remove('is-stuck');
};
setStickinessShadow();
window.addEventListener('scroll', setStickinessShadow, { passive: true });

// ===== Hitung tinggi navbar, tabs, actions untuk CSS vars =====
const tabsBar = document.getElementById('tabsBar');
const actionsDock = document.getElementById('actionsDock');
function updateHeightsVars() {
    const navH = navbarWrapper.offsetHeight || 60;
    const tabsH = tabsBar.offsetHeight || 52;
    const actH = actionsDock.offsetHeight || 56;
    rootStyle.setProperty('--nav-h', navH + 'px');
    rootStyle.setProperty('--tabs-h', tabsH + 'px');
    rootStyle.setProperty('--actions-h', actH + 'px');
}
window.addEventListener('load', updateHeightsVars);
window.addEventListener('resize', updateHeightsVars);

/* Perbarui tinggi secara reaktif bila konten berubah (wrap/unwarp) */
if (typeof ResizeObserver !== 'undefined') {
    const navResizeObserver = new ResizeObserver(updateHeightsVars);
    navResizeObserver.observe(navbarWrapper);

    const tabsResizeObserver = new ResizeObserver(updateHeightsVars);
    tabsResizeObserver.observe(tabsBar);

    const actionsResizeObserver = new ResizeObserver(updateHeightsVars);
    actionsResizeObserver.observe(actionsDock);
}

// ===== Sticky Tabs: shadow saat melekat =====
const tabsSentinel = document.getElementById('tabsSentinel');
const tabsIO = new IntersectionObserver(([entry]) => {
    if (entry && !entry.isIntersecting) tabsBar.classList.add('is-stuck');
    else tabsBar.classList.remove('is-stuck');
}, { threshold: 0 });
tabsIO.observe(tabsSentinel);

// ===== Sticky Form: shadow saat melekat =====
const formBar = document.getElementById('formBar');
const formSentinel = document.getElementById('formSentinel');
const formIO = new IntersectionObserver(([entry]) => {
    if (entry && !entry.isIntersecting) formBar.classList.add('is-stuck');
    else formBar.classList.remove('is-stuck');
}, { threshold: 0 });
formIO.observe(formSentinel);

// ===== STATE & HELPERS =====
const tabCogsBtn = document.getElementById('tabCogsBtn');
const tabSellingBtn = document.getElementById('tabSellingBtn');
const tabCalcBtn = document.getElementById('tabCalcBtn');
const tabCogs = document.getElementById('tabCogs');
const tabSelling = document.getElementById('tabSelling');
const tabCalc = document.getElementById('tabCalc');
const form = document.getElementById('boqForm');

// Theme switch (slider button)
const themeSwitchBtn = document.getElementById('themeSwitchBtn');
const toggleLabel = document.getElementById('toggleLabel');

function applyTheme(theme) {
    const root = document.documentElement;
    const isDark = theme === 'dark';
    root.classList.toggle('theme-dark', isDark);
    toggleLabel.textContent = isDark ? 'Dark' : 'Light';
    themeSwitchBtn.classList.toggle('on', isDark);
    themeSwitchBtn.setAttribute('aria-pressed', String(isDark));
    localStorage.setItem('boq_theme', isDark ? 'dark' : 'light');
    updateHeightsVars();
}
(function initTheme() {
    const saved = localStorage.getItem('boq_theme');
    if (saved === 'dark' || saved === 'light') applyTheme(saved);
    else if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) applyTheme('dark');
    else applyTheme('light');
})();
themeSwitchBtn.addEventListener('click', () => {
    applyTheme(document.documentElement.classList.contains('theme-dark') ? 'light' : 'dark');
});
themeSwitchBtn.addEventListener('keydown', (e) => {
    if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); themeSwitchBtn.click(); }
});

/* ===== AUTH (Supabase) ===== */
const SUPABASE_URL = 'https://lhrwrkcablnqewgndpsr.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_NFPcMihNQpPB10r9tRYZdg_6AaOL3lo';

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
    auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true
    }
});

const loginScreen = document.getElementById('loginScreen');
const appShell = document.getElementById('appShell');
const loginForm = document.getElementById('loginForm');
const loginEmailInput = document.getElementById('loginEmail');
const loginPasswordInput = document.getElementById('loginPassword');
const rememberLoginInput = document.getElementById('rememberLogin');
const loginError = document.getElementById('loginError');
const logoutBtn = document.getElementById('logoutBtn');
const loginThemeSwitchBtn = document.getElementById('loginThemeSwitchBtn');
const loginToggleLabel = document.getElementById('loginToggleLabel');

const appLoadingScreen = document.getElementById('appLoadingScreen');

const APP_LOADING_MIN_MS = 3000;
let appLoadingRunId = 0;

function wait(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

function showAppLoadingScreen() {
    if (!appLoadingScreen) return;
    document.body.classList.add('auth-loading');
    appLoadingScreen.classList.add('is-visible');
    appLoadingScreen.setAttribute('aria-hidden', 'false');
}

function hideAppLoadingScreen() {
    if (!appLoadingScreen) return;
    document.body.classList.remove('auth-loading');
    appLoadingScreen.classList.remove('is-visible');
    appLoadingScreen.setAttribute('aria-hidden', 'true');
}

async function runAppEntryTransition(sessionArg = null) {
    const runId = ++appLoadingRunId;

    showAppLoadingScreen();

    try {
        await Promise.all([
            setGreetingFromSession(sessionArg),
            wait(APP_LOADING_MIN_MS)
        ]);

        if (runId !== appLoadingRunId) return;

        showAppShell();
    } finally {
        if (runId === appLoadingRunId) {
            hideAppLoadingScreen();
        }
    }
}

const userMenu = document.getElementById('userMenu');
const userMenuTrigger = document.getElementById('userMenuTrigger');
const userMenuDropdown = document.getElementById('userMenuDropdown');
const userMenuGreeting = document.getElementById('userMenuGreeting');
const changePasswordBtn = document.getElementById('changePasswordBtn');

const changePasswordModal = document.getElementById('changePasswordModal');
const changePasswordCloseBtn = document.getElementById('changePasswordCloseBtn');
const cancelChangePasswordBtn = document.getElementById('cancelChangePasswordBtn');
const changePasswordForm = document.getElementById('changePasswordForm');
const currentPasswordInput = document.getElementById('currentPasswordInput');
const newPasswordInput = document.getElementById('newPasswordInput');
const confirmPasswordInput = document.getElementById('confirmPasswordInput');
const changePasswordError = document.getElementById('changePasswordError');

const USER_PROFILE_TABLE = 'user_profiles';
const USER_PROFILE_ID_COLUMN = 'user_id';
const USER_PROFILE_USERNAME_COLUMN = 'username';

let navbarGreetingRequestSeq = 0;

function setUserGreetingLoading() {
    if (!userMenuGreeting) return;

    userMenuGreeting.textContent = '';
    userMenuGreeting.classList.add('is-loading');
    userMenuGreeting.setAttribute('aria-busy', 'true');
}

function setUserGreetingText(text) {
    if (!userMenuGreeting) return;

    userMenuGreeting.classList.remove('is-loading');
    userMenuGreeting.setAttribute('aria-busy', 'false');
    userMenuGreeting.textContent = text || 'User';
}

function resetUserGreeting() {
    navbarGreetingRequestSeq++;
    setUserGreetingText('User');
}

function syncThemeToggles() {
    const isDark = document.documentElement.classList.contains('theme-dark');
    if (loginThemeSwitchBtn) {
        loginThemeSwitchBtn.classList.toggle('on', isDark);
        loginThemeSwitchBtn.setAttribute('aria-pressed', String(isDark));
    }
    if (loginToggleLabel) {
        loginToggleLabel.textContent = isDark ? 'Dark' : 'Light';
    }
}

const originalApplyTheme = applyTheme;
applyTheme = function (theme) {
    originalApplyTheme(theme);
    syncThemeToggles();
};

async function getStoredSession() {
    const { data, error } = await supabaseClient.auth.getSession();
    if (error) {
        console.error('Supabase getSession error:', error);
        return null;
    }
    return data.session || null;
}

async function clearSession() {
    const { error } = await supabaseClient.auth.signOut();
    if (error) {
        console.error('Supabase signOut error:', error);
    }
}

function getNavbarFallbackName(user) {
    return user?.email ? user.email.split('@')[0] : 'User';
}

async function fetchNavbarUsernameFromProfile(userId) {
    if (!userId) return null;

    const { data, error } = await supabaseClient
        .from(USER_PROFILE_TABLE)
        .select(USER_PROFILE_USERNAME_COLUMN)
        .eq(USER_PROFILE_ID_COLUMN, userId)
        .maybeSingle();

    if (error) {
        console.error('fetchNavbarUsernameFromProfile error:', error);
        return null;
    }

    const username = data?.[USER_PROFILE_USERNAME_COLUMN];
    if (typeof username !== 'string') return null;

    const trimmed = username.trim();
    return trimmed || null;
}

async function setGreetingFromSession(sessionArg = null) {
    if (!userMenuGreeting) return;

    const session = sessionArg || await getStoredSession();
    const user = session?.user;
    const currentRequestSeq = ++navbarGreetingRequestSeq;

    if (!user) {
        setUserGreetingText('User');
        return;
    }

    setUserGreetingLoading();

    const currentUserId = user.id;
    if (!currentUserId) {
        if (currentRequestSeq !== navbarGreetingRequestSeq) return;
        setUserGreetingText(getNavbarFallbackName(user));
        return;
    }

    void (async () => {
        const profileUsername = await fetchNavbarUsernameFromProfile(currentUserId);

        if (currentRequestSeq !== navbarGreetingRequestSeq) return;

        const resolvedName = profileUsername || getNavbarFallbackName(user);
        setUserGreetingText(resolvedName);
    })();
}

function openUserMenu() {
    if (!userMenu || !userMenuTrigger || !userMenuDropdown) return;
    userMenu.classList.add('open');
    userMenuTrigger.setAttribute('aria-expanded', 'true');
    userMenuDropdown.setAttribute('aria-hidden', 'false');
}

function closeUserMenu() {
    if (!userMenu || !userMenuTrigger || !userMenuDropdown) return;
    userMenu.classList.remove('open');
    userMenuTrigger.setAttribute('aria-expanded', 'false');
    userMenuDropdown.setAttribute('aria-hidden', 'true');
}

function toggleUserMenu() {
    if (!userMenu) return;
    if (userMenu.classList.contains('open')) closeUserMenu();
    else openUserMenu();
}

function openChangePasswordModal() {
    if (!changePasswordModal) return;
    changePasswordError.textContent = '';
    changePasswordForm.reset();
    changePasswordModal.style.display = 'block';
    changePasswordModal.setAttribute('aria-hidden', 'false');
    closeUserMenu();
    setTimeout(() => currentPasswordInput?.focus(), 0);
}

function closeChangePasswordModal() {
    if (!changePasswordModal) return;
    changePasswordModal.style.display = 'none';
    changePasswordModal.setAttribute('aria-hidden', 'true');
    changePasswordError.textContent = '';
}

function showLoginScreen() {
    appLoadingRunId++;
    hideAppLoadingScreen();
    resetUserGreeting();
    document.body.classList.add('auth-locked');
    loginScreen.classList.add('is-visible');
    loginScreen.setAttribute('aria-hidden', 'false');
    appShell.removeAttribute('aria-hidden');
    loginError.textContent = '';
    loginPasswordInput.value = '';
    syncThemeToggles();
    setTimeout(() => loginEmailInput.focus(), 0);
}

function showAppShell() {
    document.body.classList.remove('auth-locked');
    loginScreen.classList.remove('is-visible');
    loginScreen.setAttribute('aria-hidden', 'true');
    appShell.removeAttribute('aria-hidden');
}

async function initAuth() {
    const session = await getStoredSession();
    if (session?.user) {
        await runAppEntryTransition(session);
        void runStartupCloudPull(session);
    } else {
        showLoginScreen();
    }
}

loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();

    const email = loginEmailInput.value.trim();
    const password = loginPasswordInput.value;

    loginError.textContent = '';

    if (!email || !password) {
        loginError.textContent = 'Please enter your email and password.';
        return;
    }

    const { data, error } = await supabaseClient.auth.signInWithPassword({
        email,
        password
    });

    if (error || !data?.user) {
        loginError.textContent = 'Invalid credentials. Please try again.';
        loginPasswordInput.focus();
        loginPasswordInput.select();
        return;
    }

    // Untuk tahap ini checkbox "Remember me" hanya UI.
    // Session persistence sudah ditangani Supabase via persistSession: true.
    void rememberLoginInput?.checked;

    loginError.textContent = '';
    await runAppEntryTransition(data.session);
    void runStartupCloudPull(data.session);
});

if (logoutBtn) {
    logoutBtn.addEventListener('click', async () => {
        closeUserMenu();
        startupCloudPullRunId++;
        startupCloudPullDoneForUserId = null;
        clearTimeout(syncTimer);
        syncTimer = null;
        syncInFlight = false;
        await clearSession();
        showLoginScreen();
    });
}

if (userMenuTrigger) {
    userMenuTrigger.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleUserMenu();
    });
}

document.addEventListener('click', (e) => {
    if (!userMenu) return;
    if (!userMenu.contains(e.target)) closeUserMenu();
});

document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        closeUserMenu();
        closeChangePasswordModal();
    }
});

if (changePasswordBtn) {
    changePasswordBtn.addEventListener('click', openChangePasswordModal);
}

if (changePasswordCloseBtn) {
    changePasswordCloseBtn.addEventListener('click', closeChangePasswordModal);
}

if (cancelChangePasswordBtn) {
    cancelChangePasswordBtn.addEventListener('click', closeChangePasswordModal);
}

if (changePasswordModal) {
    changePasswordModal.addEventListener('click', (e) => {
        if (e.target === changePasswordModal) closeChangePasswordModal();
    });
}

if (changePasswordForm) {
    changePasswordForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const currentPassword = currentPasswordInput.value;
        const newPassword = newPasswordInput.value.trim();
        const confirmPassword = confirmPasswordInput.value.trim();

        changePasswordError.textContent = '';

        if (!currentPassword || !newPassword || !confirmPassword) {
            changePasswordError.textContent = 'Please complete all password fields.';
            return;
        }

        if (newPassword.length < 6) {
            changePasswordError.textContent = 'New password must be at least 6 characters.';
            newPasswordInput.focus();
            return;
        }

        if (newPassword !== confirmPassword) {
            changePasswordError.textContent = 'New password confirmation does not match.';
            confirmPasswordInput.focus();
            confirmPasswordInput.select();
            return;
        }

        const session = await getStoredSession();
        const email = session?.user?.email;

        if (!email) {
            changePasswordError.textContent = 'Session not found. Please login again.';
            return;
        }

        const { error: reauthError } = await supabaseClient.auth.signInWithPassword({
            email,
            password: currentPassword
        });

        if (reauthError) {
            changePasswordError.textContent = 'Current password is incorrect.';
            currentPasswordInput.focus();
            currentPasswordInput.select();
            return;
        }

        const { error: updateError } = await supabaseClient.auth.updateUser({
            password: newPassword
        });

        if (updateError) {
            changePasswordError.textContent = updateError.message || 'Failed to update password.';
            return;
        }

        closeChangePasswordModal();
        alert('Password updated successfully.');
    });
}

if (loginThemeSwitchBtn) {
    loginThemeSwitchBtn.addEventListener('click', () => {
        applyTheme(document.documentElement.classList.contains('theme-dark') ? 'light' : 'dark');
    });

    loginThemeSwitchBtn.addEventListener('keydown', (e) => {
        if (e.key === ' ' || e.key === 'Enter') {
            e.preventDefault();
            loginThemeSwitchBtn.click();
        }
    });
}

supabaseClient.auth.onAuthStateChange(async (_event, session) => {
    if (session?.user) {
        await setGreetingFromSession(session);
        if (!appLoadingScreen?.classList.contains('is-visible')) {
            showAppShell();
        }
    } else {
        showLoginScreen();
    }
});

syncThemeToggles();
initAuth();

let activeTab = 'cogs';
function setFormDisabled(disabled) {
    const controls = form.querySelectorAll('input, button, select, textarea');
    controls.forEach(el => { if (el.tagName.toLowerCase() !== 'datalist') disabled ? el.setAttribute('disabled', 'disabled') : el.removeAttribute('disabled'); });
    form.classList.toggle('form-disabled', disabled);
}
function switchTo(tab) {
    activeTab = tab;
    tabCogs.classList.add('hidden'); tabSelling.classList.add('hidden'); tabCalc.classList.add('hidden');
    [tabCogsBtn, tabSellingBtn, tabCalcBtn].forEach(btn => { btn.classList.remove('active'); btn.setAttribute('aria-selected', 'false'); });

    if (tab === 'cogs') {
        tabCogs.classList.remove('hidden'); tabCogsBtn.classList.add('active'); tabCogsBtn.setAttribute('aria-selected', 'true');
        setFormDisabled(false); // aktif di COGS
    }
    else if (tab === 'selling') {
        tabSelling.classList.remove('hidden'); tabSellingBtn.classList.add('active'); tabSellingBtn.setAttribute('aria-selected', 'true');
        setFormDisabled(false); // aktif juga di Selling
    }
    else {
        tabCalc.classList.remove('hidden'); tabCalcBtn.classList.add('active'); tabCalcBtn.setAttribute('aria-selected', 'true');
        setFormDisabled(true);  // tetap non-aktif di Calculation
        renderCalculation();
    }
    updateHeightsVars();
}
tabCogsBtn.addEventListener('click', () => switchTo('cogs'));
tabSellingBtn.addEventListener('click', () => switchTo('selling'));
tabCalcBtn.addEventListener('click', () => switchTo('calculation'));

const itemList = document.getElementById('itemList');
const priceInput = document.getElementById('price');
const marginInput = document.getElementById('margin');
const sellingInput = document.getElementById('sellingInput');
const qtyInput = document.getElementById('qty');
const unitInput = document.getElementById('unit');
const categoryInput = document.getElementById('category');
const nameInput = document.getElementById('itemName');
const grandCogsEl = document.getElementById('grandCogs');
const grandSellingEl = document.getElementById('grandSelling');
const projectSearch = document.getElementById('projectSearch');
const projectList = document.getElementById('projectList');
const navProject = document.getElementById('navProject');
const renameProjectBtn = document.getElementById('renameProjectBtn');
const downloadJSONBtn = document.getElementById('downloadJSON');
const restoreJSONBtn = document.getElementById('restoreJSON');
const restoreInput = document.getElementById('restoreInput');
const manageBtn = document.getElementById('manageBtn');
const manageModal = document.getElementById('manageModal');
const projectsModal = document.getElementById('projectsModal');
const modalClose = document.getElementById('modalClose');
const projectsClose = document.getElementById('projectsClose');
const manageItemList = document.getElementById('manageItemList');
const manageProjectsBtn = document.getElementById('manageProjectsBtn');
const manageProjectList = document.getElementById('manageProjectList');

const manageItemSearch = document.getElementById('manageItemSearch');
const manageProjectSearch = document.getElementById('manageProjectSearch');

const editBtn = document.getElementById('editBtn');
const deleteBtn = document.getElementById('deleteBtn');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn');
const loadProjectBtn = document.getElementById('loadProjectBtn');
const saveProjectBtn = document.getElementById('saveProjectBtn');
const newProjectBtn = document.getElementById('newProjectBtn');

const calcTotalCogsEl = document.getElementById('calcTotalCogs');
const calcTotalSellingEl = document.getElementById('calcTotalSelling');
const commissionCellInput = document.getElementById('commissionCellInput');
const calcGrossProfitEl = document.getElementById('calcGrossProfit');
const calcTotalMarginEl = document.getElementById('calcTotalMargin');

const lastSavedTextEl = document.getElementById('lastSavedText');
const lastSyncedTextEl = document.getElementById('lastSyncedText');
const LOCAL_META_KEY = 'boq_sync_meta_v1';

let items = JSON.parse(localStorage.getItem('boq_items')) || [];
let tableData = JSON.parse(localStorage.getItem('boq_working')) || [];
let selectedIndex = null;
let editingIndex = null;
let sellingInputManuallyEdited = false;
let currentProjectName = localStorage.getItem('boq_current_name') || "";
const PROJECTS_KEY = 'boq_projects_v2';
const UNSAVED_COMM_KEY = 'boq_unsaved_commission';
const CLOUD_STATE_TABLE = 'user_app_state';
const CLOUD_SYNC_DEBOUNCE_MS = 3000;

let startupCloudPullDoneForUserId = null;
let startupCloudPullRunId = 0;
let syncTimer = null;
let syncInFlight = false;
let suppressScheduledSync = false;

// ===== Reorder mode =====
const reorderBtn = document.getElementById('reorderBtn');
let reorderMode = false;
function updateReorderButton() {
    reorderBtn.textContent = reorderMode ? 'Done Reorder' : 'Reorder';
    reorderBtn.classList.toggle('secondary', !reorderMode);
    reorderBtn.title = reorderMode ? 'Selesai ubah urutan' : 'Aktifkan drag & drop untuk ubah urutan';
}
function clearDragVisuals() {
    document.querySelectorAll('.drag-over, .dragging').forEach(el => el.classList.remove('drag-over', 'dragging'));
}
function setReorderMode(on) {
    reorderMode = !!on;
    document.body.classList.toggle('reorder-mode', reorderMode); // gate DnD vs text-select
    clearDragVisuals();
    updateReorderButton();
    renderAll();
}
reorderBtn.addEventListener('click', () => setReorderMode(!reorderMode));

// ===== Category order =====
let categoryOrder = JSON.parse(localStorage.getItem('boq_category_order')) || [];
function getAllCategories() {
    const set = [];
    tableData.forEach(r => {
        const c = sanitize(r.category || 'Uncategorized');
        if (!set.includes(c)) set.push(c);
    });
    return set;
}
function normalizeCategoryOrder() {
    const cats = getAllCategories();
    categoryOrder = categoryOrder.filter(c => cats.includes(c));
    cats.forEach(c => { if (!categoryOrder.includes(c)) categoryOrder.push(c); });
    localStorage.setItem('boq_category_order', JSON.stringify(categoryOrder));
    if (currentProjectName) {
        const projects = loadProjectsObj();
        if (projects[currentProjectName]) {
            projects[currentProjectName].categoryOrder = categoryOrder.slice();
            saveProjectsObj(projects);
        }
    }
}

let currentCommission = 0;

function sanitize(s) { return String(s || '').replace(/<\/?[^>]+(>|$)/g, "").trim(); }
function getSafeFileName(name) { if (!name) return 'BOQ'; return (`BOQ - ${name}`).replace(/[\\\/:*?"<>|]/g, '').trim(); }
function updateNavProject() { navProject.textContent = currentProjectName ? `| ${currentProjectName}` : '(no project)'; }
function showAlert(msg) { alert(msg); }
function parseNumberFormatted(str) { if (str == null) return null; const raw = String(str).replace(/\./g, '').replace(/,/g, '').trim(); return raw === '' ? null : Number(raw); }
function formatThousandsInput(el) { el.addEventListener('input', () => { const raw = String(el.value).replace(/[^\d]/g, ''); el.value = raw ? Number(raw).toLocaleString('id-ID') : ''; }); }
formatThousandsInput(priceInput);
formatThousandsInput(sellingInput);

function excelColLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function updateDatalist() {
    const map = {};
    items.forEach(it => { if (!it || !it.name) return; const key = it.name.trim(); if (!map[key] || (it.updatedAt && map[key].updatedAt < it.updatedAt)) map[key] = it; });
    const arr = Object.values(map).slice().sort((a, b) => (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase()));
    itemList.innerHTML = '';
    arr.forEach(it => { const opt = document.createElement('option'); opt.value = it.name; itemList.appendChild(opt); });
}

function calcSellingPrice(cogs, marginPercent) { const m = Number(marginPercent) / 100; if (!isFinite(m) || m >= 1) return null; if (cogs === "" || cogs == null) return null; return Number(cogs) / (1 - m); }
function roundUpToThousand(n) { if (n == null || !isFinite(n)) return null; const eps = 1e-9; return Math.ceil((n - eps) / 1000) * 1000; }

function normalizeKey(name) { return String(name || '').trim().toLowerCase(); }
function getItemByName(name) { if (!name) return null; const key = normalizeKey(name); let best = null; for (const it of items) { if (!it || !it.name) continue; if (normalizeKey(it.name) === key) { if (!best || (it.updatedAt || 0) > (best.updatedAt || 0)) best = it; } } return best; }

function updateItemsWithEntry(name, unit, price, sellingPrice, category) {
    if (!name) return;
    const key = normalizeKey(name);
    let foundIndex = -1;
    for (let i = 0; i < items.length; i++) { if (normalizeKey(items[i].name) === key) { foundIndex = i; break; } }
    const now = Date.now();
    if (foundIndex >= 0) {
        const it = items[foundIndex];
        if (unit !== undefined) it.unit = unit;
        if (price !== undefined) it.price = price;
        if (sellingPrice !== undefined) it.sellingPrice = sellingPrice;
        if (category !== undefined) it.category = category;
        it.updatedAt = now;
        items[foundIndex] = it;
    } else {
        items.push({ name, unit: unit || "", price: (price !== undefined && price !== null) ? price : "", sellingPrice: (sellingPrice !== undefined && sellingPrice !== null) ? sellingPrice : "", category: category || "", updatedAt: now });
    }
    const map = {};
    items.forEach(it => { const k = normalizeKey(it.name); if (!map[k] || (it.updatedAt || 0) > (map[k].updatedAt || 0)) map[k] = it; });
    items = Object.values(map).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    localStorage.setItem('boq_items', JSON.stringify(items));
    updateDatalist();
}

function autofillFromItem(name) {
    const it = getItemByName(name);
    if (!it) return;
    if (it.unit !== undefined) unitInput.value = it.unit || "";
    if (it.price !== undefined && it.price !== "" && it.price !== null) {
        priceInput.value = Number(it.price).toLocaleString('id-ID');
        sellingInputManuallyEdited = false;
        computeAndDisplaySelling();
    } else {
        priceInput.value = '';
        if (it.sellingPrice !== undefined && it.sellingPrice !== "" && it.sellingPrice !== null) {
            sellingInput.value = Number(it.sellingPrice).toLocaleString('id-ID');
            sellingInputManuallyEdited = false;
        } else {
            sellingInput.value = '';
            sellingInputManuallyEdited = false;
        }
    }
    if (it.category !== undefined) categoryInput.value = it.category || "";
}

nameInput.addEventListener('change', (e) => {
    const v = sanitize(e.target.value);
    if (!v) return;
    const found = getItemByName(v);
    if (found) autofillFromItem(v);
});

sellingInput.addEventListener('input', () => { sellingInputManuallyEdited = true; });

function computeAndDisplaySelling(opts = {}) {
    const { force = false } = opts;
    const priceNum = parseNumberFormatted(priceInput.value);
    const marginNum = (marginInput.value !== '') ? Number(marginInput.value) : null;
    const canCompute = (priceNum != null && priceNum !== '' && marginNum != null && isFinite(marginNum) && marginNum < 100);
    if (!canCompute) return;

    const spRaw = calcSellingPrice(Number(priceNum), Number(marginNum));
    const spRounded = (spRaw != null) ? roundUpToThousand(spRaw) : null;

    if (spRounded != null && (force || !sellingInputManuallyEdited)) {
        sellingInput.value = Number(spRounded).toLocaleString('id-ID');
        sellingInputManuallyEdited = false;
    }
}
priceInput.addEventListener('input', () => { sellingInputManuallyEdited = false; computeAndDisplaySelling(); });
marginInput.addEventListener('input', () => { sellingInputManuallyEdited = false; computeAndDisplaySelling(); });

function computeTotals() {
    const totalCogs = tableData.reduce((acc, r) => { const p = (r.price != null && r.price !== "") ? Number(r.price) : 0; const q = Number(r.qty || 0); return acc + (Number(p || 0) * q); }, 0);
    const totalSelling = tableData.reduce((acc, r) => {
        let sp = null;
        if (r.sellingPrice != null && r.sellingPrice !== "") sp = Number(r.sellingPrice);
        else if (r.price != null && r.price !== "") sp = calcSellingPrice(Number(r.price), r.margin);
        const spr = sp != null ? roundUpToThousand(sp) : 0;
        const q = Number(r.qty || 0);
        return acc + (spr ? (spr * q) : 0);
    }, 0);
    return { totalCogs, totalSelling };
}

/* ===== Drag & Drop State ===== */
const dragState = { type: null, cat: null, fromPos: null, fromCat: null };

function moveItemInCategory(category, fromPos, toPos) {
    if (fromPos === toPos) return;
    const cat = sanitize(category || 'Uncategorized');
    const indices = [];
    tableData.forEach((r, i) => { if (sanitize(r.category || 'Uncategorized') === cat) indices.push(i); });
    if (!indices.length) return;
    if (fromPos < 0 || fromPos >= indices.length || toPos < 0 || toPos >= indices.length) return;

    const fromIndex = indices[fromPos];
    const [item] = tableData.splice(fromIndex, 1);

    const indicesAfter = [];
    tableData.forEach((r, i) => { if (sanitize(r.category || 'Uncategorized') === cat) indicesAfter.push(i); });
    const toIndex = indicesAfter[toPos];

    tableData.splice(toIndex, 0, item);

    localStorage.setItem('boq_working', JSON.stringify(tableData));
    renderAll();
    persistProject();
}

function addRowDragHandlers(tr, cat, posInCat) {
    tr.classList.add('draggable');
    tr.setAttribute('draggable', 'true');

    tr.addEventListener('dragstart', (e) => {
        dragState.type = 'row';
        dragState.cat = cat;
        dragState.fromPos = posInCat;
        tr.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        try { e.dataTransfer.setData('text/plain', `row|${cat}|${posInCat}`); } catch (_) { }
    });

    tr.addEventListener('dragend', () => {
        tr.classList.remove('dragging');
        document.querySelectorAll('tr.drag-over').forEach(x => x.classList.remove('drag-over'));
        dragState.type = null; dragState.cat = null; dragState.fromPos = null;
    });

    tr.addEventListener('dragover', (e) => {
        if (dragState.type === 'row' && dragState.cat === cat) {
            e.preventDefault();
            tr.classList.add('drag-over');
            e.dataTransfer.dropEffect = 'move';
        }
    });

    tr.addEventListener('dragleave', () => tr.classList.remove('drag-over'));

    tr.addEventListener('drop', (e) => {
        e.preventDefault();
        tr.classList.remove('drag-over');
        if (!(dragState.type === 'row' && dragState.cat === cat && dragState.fromPos != null)) return;
        const toPos = posInCat;
        moveItemInCategory(cat, dragState.fromPos, toPos);
    });
}

function addSectionDragHandlers(trSec, cat) {
    trSec.classList.add('draggable');
    trSec.setAttribute('draggable', 'true');

    trSec.addEventListener('dragstart', (e) => {
        dragState.type = 'group';
        dragState.fromCat = cat;
        trSec.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        try { e.dataTransfer.setData('text/plain', `group|${cat}`); } catch (_) { }
    });

    trSec.addEventListener('dragend', () => {
        trSec.classList.remove('dragging');
        document.querySelectorAll('tr.section-row.drag-over').forEach(x => x.classList.remove('drag-over'));
        dragState.type = null; dragState.fromCat = null;
    });

    trSec.addEventListener('dragover', (e) => {
        if (dragState.type === 'group') {
            e.preventDefault();
            trSec.classList.add('drag-over');
            e.dataTransfer.dropEffect = 'move';
        }
    });

    trSec.addEventListener('dragleave', () => trSec.classList.remove('drag-over'));

    trSec.addEventListener('drop', (e) => {
        e.preventDefault();
        trSec.classList.remove('drag-over');
        if (dragState.type !== 'group' || !dragState.fromCat) return;
        const fromIdx = categoryOrder.indexOf(dragState.fromCat);
        const toIdx = categoryOrder.indexOf(cat);
        if (fromIdx === -1 || toIdx === -1 || fromIdx === toIdx) return;
        const [moved] = categoryOrder.splice(fromIdx, 1);
        categoryOrder.splice(toIdx, 0, moved);
        localStorage.setItem('boq_category_order', JSON.stringify(categoryOrder));
        persistProject();
        renderAll();
    });
}

function buildGroupedByCategory() {
    const grouped = {};
    tableData.forEach(r => {
        const cat = sanitize(r.category || 'Uncategorized');
        if (!grouped[cat]) grouped[cat] = [];
        grouped[cat].push(r);
    });
    normalizeCategoryOrder();
    return { grouped, orderedCats: categoryOrder.slice() };
}

function renderCogsTable() {
    const cogsTbody = document.querySelector('#cogsTable tbody');
    cogsTbody.innerHTML = '';
    const { grouped, orderedCats } = buildGroupedByCategory();
    let idx = 1; let grand = 0;

    orderedCats.forEach(cat => {
        const trSec = document.createElement('tr'); trSec.className = 'section-row';
        trSec.dataset.cat = cat;
        const tdSec = document.createElement('td'); tdSec.colSpan = 7; tdSec.textContent = cat;
        trSec.appendChild(tdSec);
        if (reorderMode) addSectionDragHandlers(trSec, cat);
        cgsTbodyAppend(trSec, cogsTbody);

        (grouped[cat] || []).forEach((row, posInCat) => {
            const tr = document.createElement('tr');

            const tdNo = document.createElement('td'); tdNo.textContent = idx++;
            const tdName = document.createElement('td'); tdName.className = 'left'; tdName.textContent = sanitize(row.name);
            const tdQty = document.createElement('td'); tdQty.textContent = row.qty !== "" ? row.qty : "";
            const tdUnit = document.createElement('td'); tdUnit.textContent = row.unit || "";
            const tdCogs = document.createElement('td'); tdCogs.style.textAlign = 'right';
            const tdTotal = document.createElement('td'); tdTotal.style.textAlign = 'right';
            const tdMargin = document.createElement('td'); tdMargin.style.textAlign = 'right';

            const priceNum = (row.price != null && row.price !== "") ? Number(row.price) : null;
            const qtyNum = (row.qty != null && row.qty !== "") ? Number(row.qty) : null;
            const totalCogs = (priceNum != null && qtyNum != null) ? Number(priceNum) * Number(qtyNum) : null;
            tdCogs.textContent = (priceNum != null && isFinite(priceNum)) ? Number(priceNum).toLocaleString('id-ID') : "";
            tdTotal.textContent = (totalCogs != null && isFinite(totalCogs)) ? Number(totalCogs).toLocaleString('id-ID') : "";
            tdMargin.textContent = (row.margin != null && row.margin !== "") ? Number(row.margin).toFixed(2) : "";
            if (totalCogs != null && isFinite(totalCogs)) grand += Number(totalCogs);

            tr.appendChild(tdNo); tr.appendChild(tdName); tr.appendChild(tdQty); tr.appendChild(tdUnit); tr.appendChild(tdCogs); tr.appendChild(tdTotal); tr.appendChild(tdMargin);

            tr.addEventListener('click', () => {
                selectedIndex = tableData.indexOf(row);
                document.querySelectorAll('#cogsTable tbody tr, #sellingTable tbody tr').forEach(r => r.classList.remove('selected'));
                tr.classList.add('selected');
            });

            if (reorderMode) addRowDragHandlers(tr, cat, posInCat);

            cgsTbodyAppend(tr, cogsTbody);
        });
    });
    grandCogsEl.textContent = grand.toLocaleString('id-ID');
}

function renderSellingTable() {
    const sellingTbody = document.querySelector('#sellingTable tbody');
    sellingTbody.innerHTML = '';
    const { grouped, orderedCats } = buildGroupedByCategory();
    let idx = 1; let grand = 0;

    orderedCats.forEach(cat => {
        const trSec = document.createElement('tr'); trSec.className = 'section-row';
        trSec.dataset.cat = cat;
        const tdSec = document.createElement('td'); tdSec.colSpan = 6; tdSec.textContent = cat;
        trSec.appendChild(tdSec);
        if (reorderMode) addSectionDragHandlers(trSec, cat);
        sellingTbody.appendChild(trSec);

        (grouped[cat] || []).forEach((row, posInCat) => {
            const tr = document.createElement('tr');

            const tdNo = document.createElement('td'); tdNo.textContent = idx++;
            const tdName = document.createElement('td'); tdName.className = 'left'; tdName.textContent = sanitize(row.name);
            const tdQty = document.createElement('td'); tdQty.textContent = row.qty !== "" ? row.qty : "";
            const tdUnit = document.createElement('td'); tdUnit.textContent = row.unit || "";
            const tdSelling = document.createElement('td'); tdSelling.style.textAlign = 'right';
            const tdTotalSelling = document.createElement('td'); tdTotalSelling.style.textAlign = 'right';

            let sellingPriceRaw = null;
            if (row.sellingPrice != null && row.sellingPrice !== "") sellingPriceRaw = Number(row.sellingPrice);
            else { const priceNum = (row.price != null && row.price !== "") ? Number(row.price) : null; sellingPriceRaw = (priceNum != null) ? calcSellingPrice(priceNum, row.margin) : null; }
            const sellingPriceRounded = sellingPriceRaw != null ? roundUpToThousand(sellingPriceRaw) : null;
            const qtyNum = (row.qty != null && row.qty !== "") ? Number(row.qty) : null;
            const totalSelling = (sellingPriceRounded != null && qtyNum != null) ? Number(sellingPriceRounded) * Number(qtyNum) : null;
            tdSelling.textContent = (sellingPriceRounded != null && isFinite(sellingPriceRounded)) ? Number(sellingPriceRounded).toLocaleString('id-ID') : "";
            tdTotalSelling.textContent = (totalSelling != null && isFinite(totalSelling)) ? Number(totalSelling).toLocaleString('id-ID') : "";
            if (totalSelling != null && isFinite(totalSelling)) grand += Number(totalSelling);

            tr.appendChild(tdNo); tr.appendChild(tdName); tr.appendChild(tdQty); tr.appendChild(tdUnit); tr.appendChild(tdSelling); tr.appendChild(tdTotalSelling);

            tr.addEventListener('click', () => {
                selectedIndex = tableData.indexOf(row);
                document.querySelectorAll('#cogsTable tbody tr, #sellingTable tbody tr').forEach(r => r.classList.remove('selected'));
                tr.classList.add('selected');
            });

            if (reorderMode) addRowDragHandlers(tr, cat, posInCat);

            sellingTbody.appendChild(tr);
        });
    });
    grandSellingEl.textContent = grand.toLocaleString('id-ID');
}

function loadProjectsObj() { try { return JSON.parse(localStorage.getItem(PROJECTS_KEY)) || {}; } catch (e) { return {}; } }
function saveProjectsObj(obj) { localStorage.setItem(PROJECTS_KEY, JSON.stringify(obj)); }

function serializeAppState() {
    return {
        projects: loadProjectsObj(),
        items: Array.isArray(items) ? items : [],
        working: Array.isArray(tableData) ? tableData : [],
        currentProjectName: currentProjectName || '',
        categoryOrder: Array.isArray(categoryOrder) ? categoryOrder : [],
        unsavedCommission: Number(currentCommission || 0),
        meta: {
            schemaVersion: 1
        }
    };
}

function applyAppState(snapshot) {
    const safe = snapshot && typeof snapshot === 'object' ? snapshot : {};

    const nextProjects =
        safe.projects && typeof safe.projects === 'object' ? safe.projects : {};
    const nextItems =
        Array.isArray(safe.items) ? safe.items : [];
    const nextWorking =
        Array.isArray(safe.working) ? safe.working : [];
    const nextCurrentProjectName =
        typeof safe.currentProjectName === 'string' ? safe.currentProjectName : '';
    const nextCategoryOrder =
        Array.isArray(safe.categoryOrder) ? safe.categoryOrder : [];
    const nextUnsavedCommission =
        Number(safe.unsavedCommission || 0);

    saveProjectsObj(nextProjects);
    localStorage.setItem('boq_items', JSON.stringify(nextItems));
    localStorage.setItem('boq_working', JSON.stringify(nextWorking));
    localStorage.setItem('boq_category_order', JSON.stringify(nextCategoryOrder));

    if (nextCurrentProjectName) {
        localStorage.setItem('boq_current_name', nextCurrentProjectName);
    } else {
        localStorage.removeItem('boq_current_name');
    }

    if (nextUnsavedCommission === 0) {
        localStorage.removeItem(UNSAVED_COMM_KEY);
    } else {
        localStorage.setItem(UNSAVED_COMM_KEY, String(nextUnsavedCommission));
    }

    items = nextItems;
    tableData = nextWorking;
    categoryOrder = nextCategoryOrder;
    currentProjectName = nextCurrentProjectName;
    currentCommission = nextUnsavedCommission;

    updateDatalist();
    updateProjectList();
    updateNavProject();
    renderAll();
    updateFooterStatus();
}

function renderCalculation() {
    const totals = computeTotals();
    const totalCogs = totals.totalCogs || 0;
    const totalSelling = totals.totalSelling || 0;
    const commission = Number(currentCommission || 0);
    const grossProfit = totalSelling - totalCogs - commission;
    const totalMargin = (totalSelling === 0) ? 0 : (grossProfit / totalSelling);
    calcTotalCogsEl.textContent = Number(totalCogs || 0).toLocaleString('id-ID');
    calcTotalSellingEl.textContent = Number(totalSelling || 0).toLocaleString('id-ID');
    commissionCellInput.value = commission ? Number(commission).toLocaleString('id-ID') : '';
    calcGrossProfitEl.textContent = Number(grossProfit || 0).toLocaleString('id-ID');
    calcTotalMarginEl.textContent = (isFinite(totalMargin) ? (totalMargin * 100).toFixed(2) : '0.00') + '%';
}

(function setupCommissionCell() {
    commissionCellInput.addEventListener('input', () => {
        const rawDigits = String(commissionCellInput.value).replace(/[^\d]/g, '');
        commissionCellInput.value = rawDigits ? Number(rawDigits).toLocaleString('id-ID') : '';
        const parsed = parseNumberFormatted(commissionCellInput.value) || 0;
        currentCommission = Number(parsed);
        persistProject();
        renderCalculation();
    });
})();

function loadCommissionForCurrentProject() {
    if (currentProjectName) {
        const projects = loadProjectsObj();
        const p = projects[currentProjectName];
        if (p && typeof p === 'object') { currentCommission = Number(p.commission || 0); return; }
    }
    const unsaved = parseNumberFormatted(localStorage.getItem(UNSAVED_COMM_KEY));
    currentCommission = unsaved != null ? Number(unsaved) : 0;
}

function getLocalMeta() {
    try {
        const parsed = JSON.parse(localStorage.getItem(LOCAL_META_KEY));
        return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
        return {};
    }
}

function saveLocalMeta(metaPatch) {
    const current = getLocalMeta();
    const next = { ...current, ...metaPatch };
    localStorage.setItem(LOCAL_META_KEY, JSON.stringify(next));
    return next;
}

function formatLastSaved(ts) {
    if (!ts) return '—';
    try {
        return new Date(ts).toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' });
    } catch {
        return '—';
    }
}

function updateFooterLastSaved() {
    if (!lastSavedTextEl) return;
    if (!currentProjectName) {
        lastSavedTextEl.textContent = '—';
        return;
    }
    const projects = loadProjectsObj();
    const p = projects[currentProjectName];
    lastSavedTextEl.textContent = p && p.lastSaved ? formatLastSaved(p.lastSaved) : '—';
}

function updateFooterLastSynced() {
    if (!lastSyncedTextEl) return;
    const meta = getLocalMeta();
    lastSyncedTextEl.textContent = meta.lastSyncedAt
        ? formatLastSaved(meta.lastSyncedAt)
        : 'Not yet';
}

function updateFooterStatus() {
    updateFooterLastSaved();
    updateFooterLastSynced();
}

function markLastSynced(ts = Date.now()) {
    saveLocalMeta({ lastSyncedAt: ts });
    updateFooterLastSynced();
}

function setSyncedFooterText(text) {
    if (!lastSyncedTextEl) return;
    lastSyncedTextEl.textContent = text;
}

async function getCurrentUserId(sessionArg = null) {
    const session = sessionArg || await getStoredSession();
    return session?.user?.id || null;
}

function getSnapshotClientUpdatedAt(snapshot) {
    if (!snapshot || typeof snapshot !== 'object') return null;

    const projects = snapshot.projects && typeof snapshot.projects === 'object'
        ? snapshot.projects
        : {};

    let latest = 0;
    Object.values(projects).forEach((p) => {
        const ts = Number(p?.lastSaved || 0);
        if (ts > latest) latest = ts;
    });

    return latest || null;
}

function isSnapshotStructurallyEmpty(snapshot) {
    if (!snapshot || typeof snapshot !== 'object') return true;

    const projects = snapshot.projects && typeof snapshot.projects === 'object' ? snapshot.projects : {};
    const itemsArr = Array.isArray(snapshot.items) ? snapshot.items : [];
    const workingArr = Array.isArray(snapshot.working) ? snapshot.working : [];
    const categoryArr = Array.isArray(snapshot.categoryOrder) ? snapshot.categoryOrder : [];
    const currentName = typeof snapshot.currentProjectName === 'string' ? snapshot.currentProjectName.trim() : '';
    const unsavedComm = Number(snapshot.unsavedCommission || 0);

    return (
        Object.keys(projects).length === 0 &&
        itemsArr.length === 0 &&
        workingArr.length === 0 &&
        categoryArr.length === 0 &&
        currentName === '' &&
        unsavedComm === 0
    );
}

async function fetchCloudState(sessionArg = null) {
    const userId = await getCurrentUserId(sessionArg);
    if (!userId) return null;

    const { data, error } = await supabaseClient
        .from(CLOUD_STATE_TABLE)
        .select('state, client_updated_at, updated_at')
        .eq('user_id', userId)
        .maybeSingle();

    if (error) {
        console.error('fetchCloudState error:', error);
        return null;
    }

    if (!data) return null;

    return {
        state: data.state || null,
        clientUpdatedAt: data.client_updated_at || null,
        updatedAt: data.updated_at || null
    };
}

async function pushCloudStateManual(sessionArg = null) {
    const ok = await pushCloudState(sessionArg);

    if (!ok) {
        alert('Push ke cloud gagal. Lihat console.');
        return false;
    }

    alert('Push ke cloud berhasil.');
    return true;
}

async function pullCloudStateManual(options = {}) {
    const { forceReplace = false } = options;

    const cloud = await fetchCloudState();
    if (!cloud || !cloud.state) {
        alert('Data cloud belum ada.');
        return false;
    }

    const localSnapshot = serializeAppState();
    const localTs = getSnapshotClientUpdatedAt(localSnapshot);
    const cloudTs = getSnapshotClientUpdatedAt(cloud.state);
    const localEmpty = isSnapshotStructurallyEmpty(localSnapshot);
    const cloudEmpty = isSnapshotStructurallyEmpty(cloud.state);

    if (cloudEmpty) {
        alert('Snapshot cloud kosong.');
        return false;
    }

    if (localEmpty) {
        applyAppState(cloud.state);
        if (cloud.updatedAt) {
            const parsed = Date.parse(cloud.updatedAt);
            if (!Number.isNaN(parsed)) markLastSynced(parsed);
            else updateFooterLastSynced();
        } else {
            updateFooterLastSynced();
        }
        alert('Data cloud berhasil dimuat ke local.');
        return true;
    }

    if (forceReplace) {
        applyAppState(cloud.state);
        if (cloud.updatedAt) {
            const parsed = Date.parse(cloud.updatedAt);
            if (!Number.isNaN(parsed)) markLastSynced(parsed);
            else updateFooterLastSynced();
        } else {
            updateFooterLastSynced();
        }
        alert('Data local berhasil ditimpa dari cloud.');
        return true;
    }

    if (cloudTs && (!localTs || cloudTs > localTs)) {
        const ok = confirm('Data cloud lebih baru dari local. Timpa local dengan data cloud?');
        if (!ok) return false;

        applyAppState(cloud.state);
        if (cloud.updatedAt) {
            const parsed = Date.parse(cloud.updatedAt);
            if (!Number.isNaN(parsed)) markLastSynced(parsed);
            else updateFooterLastSynced();
        } else {
            updateFooterLastSynced();
        }
        alert('Data cloud berhasil dimuat ke local.');
        return true;
    }

    if (localTs && (!cloudTs || localTs > cloudTs)) {
        alert('Data local lebih baru dari cloud. Tidak dilakukan overwrite.');
        return false;
    }

    applyAppState(cloud.state);
    if (cloud.updatedAt) {
        const parsed = Date.parse(cloud.updatedAt);
        if (!Number.isNaN(parsed)) markLastSynced(parsed);
        else updateFooterLastSynced();
    } else {
        updateFooterLastSynced();
    }
    alert('Data cloud berhasil dimuat ke local.');
    return true;
}

async function runStartupCloudPull(sessionArg = null) {
    const session = sessionArg || await getStoredSession();
    const userId = session?.user?.id || null;

    if (!userId) return false;
    if (startupCloudPullDoneForUserId === userId) return false;

    const runId = ++startupCloudPullRunId;
    startupCloudPullDoneForUserId = userId;

    const cloud = await fetchCloudState(session);
    if (runId !== startupCloudPullRunId) return false;

    if (!cloud || !cloud.state) {
        updateFooterLastSynced();
        return false;
    }

    const localSnapshot = serializeAppState();
    const localTs = getSnapshotClientUpdatedAt(localSnapshot);
    const cloudTs = getSnapshotClientUpdatedAt(cloud.state);
    const localEmpty = isSnapshotStructurallyEmpty(localSnapshot);
    const cloudEmpty = isSnapshotStructurallyEmpty(cloud.state);

    if (cloudEmpty) {
        updateFooterLastSynced();
        return false;
    }

    if (localEmpty) {
        applyAppState(cloud.state);

        if (runId !== startupCloudPullRunId) return false;

        if (cloud.updatedAt) {
            const parsed = Date.parse(cloud.updatedAt);
            if (!Number.isNaN(parsed)) markLastSynced(parsed);
            else updateFooterLastSynced();
        } else {
            updateFooterLastSynced();
        }

        return true;
    }

    if (cloudTs && (!localTs || cloudTs > localTs)) {
        applyAppState(cloud.state);

        if (runId !== startupCloudPullRunId) return false;

        if (cloud.updatedAt) {
            const parsed = Date.parse(cloud.updatedAt);
            if (!Number.isNaN(parsed)) markLastSynced(parsed);
            else updateFooterLastSynced();
        } else {
            updateFooterLastSynced();
        }

        return true;
    }

    updateFooterLastSynced();
    return false;
}

async function pushCloudState(sessionArg = null) {
    const userId = await getCurrentUserId(sessionArg);
    if (!userId) return false;

    const snapshot = serializeAppState();
    const clientUpdatedAtMs = getSnapshotClientUpdatedAt(snapshot);
    const clientUpdatedAtIso = clientUpdatedAtMs
        ? new Date(clientUpdatedAtMs).toISOString()
        : new Date().toISOString();

    const payload = {
        user_id: userId,
        state: snapshot,
        client_updated_at: clientUpdatedAtIso,
        app_version: 'snapshot-v1'
    };

    const { error } = await supabaseClient
        .from(CLOUD_STATE_TABLE)
        .upsert(payload, { onConflict: 'user_id' });

    if (error) {
        console.error('pushCloudState error:', error);
        return false;
    }

    markLastSynced(Date.now());
    return true;
}

function scheduleCloudSync() {
    if (suppressScheduledSync) return;

    clearTimeout(syncTimer);
    setSyncedFooterText('Pending');

    syncTimer = setTimeout(async () => {
        if (syncInFlight) return;

        syncInFlight = true;
        setSyncedFooterText('Syncing...');

        try {
            const ok = await pushCloudState();
            if (!ok) {
                setSyncedFooterText('Failed');
            }
        } catch (err) {
            console.error('scheduleCloudSync error:', err);
            setSyncedFooterText('Failed');
        } finally {
            syncInFlight = false;
        }
    }, CLOUD_SYNC_DEBOUNCE_MS);
}

function touchCloudSync() {
    scheduleCloudSync();
}

window.__boqCloudDebug = {
    fetch: fetchCloudState,
    push: pushCloudStateManual,
    pull: pullCloudStateManual
};

function persistProject() {
    if (currentProjectName) {
        const projects = loadProjectsObj();
        projects[currentProjectName] = {
            data: tableData,
            commission: Number(currentCommission || 0),
            categoryOrder: categoryOrder.slice(),
            lastSaved: Date.now()
        };
        saveProjectsObj(projects);
        updateFooterLastSaved();
        touchCloudSync();
    } else {
        if (currentCommission === 0) localStorage.removeItem(UNSAVED_COMM_KEY);
        else localStorage.setItem(UNSAVED_COMM_KEY, String(currentCommission));
        touchCloudSync();
    }
}

function renderAll() { normalizeCategoryOrder(); renderCogsTable(); renderSellingTable(); renderCalculation(); }

function setEditingMode(idx) { editingIndex = (typeof idx === 'number') ? idx : null; updateEditButton(); }
function cancelEdit() { setEditingMode(null); selectedIndex = null; sellingInputManuallyEdited = false; form.reset(); document.querySelectorAll('#cogsTable tbody tr, #sellingTable tbody tr').forEach(r => r.classList.remove('selected')); nameInput.focus(); }
function updateEditButton() { editBtn.textContent = (editingIndex === null) ? 'Edit' : 'Cancel Edit'; }
function updateEditBtnInit() { editBtn.textContent = (editingIndex === null) ? 'Edit' : 'Cancel Edit'; }

function clearSelectionOnly() {
    selectedIndex = null;
    document.querySelectorAll('#cogsTable tbody tr, #sellingTable tbody tr').forEach(r => r.classList.remove('selected'));
}

/* Unselect row when clicking outside the tables */
document.addEventListener('click', (e) => {
    const insideAnyTable = e.target.closest('#cogsTable') || e.target.closest('#sellingTable');
    if (!insideAnyTable) {
        clearSelectionOnly();
    }
});

form.addEventListener('submit', e => {
    e.preventDefault();
    if (form.classList.contains('form-disabled')) return;

    const name = sanitize(nameInput.value);
    const category = sanitize(categoryInput.value);
    if (!name || !category) { showAlert('Nama item dan kategori harus diisi.'); return; }
    const qtyRaw = qtyInput.value;
    const qty = qtyRaw !== "" ? Number(qtyRaw) : "";
    const unit = sanitize(unitInput.value);
    const priceRaw = priceInput.value.replace(/\./g, '').replace(/,/g, '');
    const price = priceRaw !== "" ? Number(priceRaw) : "";
    const marginRaw = marginInput.value;
    const margin = marginRaw !== "" ? Number(marginRaw) : 0;
    const sellingRawStr = sellingInput.value.replace(/\./g, '').replace(/,/g, '');
    const sellingProvided = sellingRawStr !== "" ? Number(sellingRawStr) : "";
    if (margin >= 100) { if (!confirm('Margin 100% atau lebih akan membuat Harga Jual tidak terdefinisi. Lanjutkan?')) return; }
    const totalCogs = (qty !== "" && price !== "") ? Number(qty) * Number(price) : "";
    let sellingPriceComputed = "";
    if (sellingProvided !== "") sellingPriceComputed = sellingProvided;
    else if (price !== "" && margin < 100) { const sp = calcSellingPrice(price, margin); sellingPriceComputed = (sp != null) ? sp : ""; }
    else sellingPriceComputed = "";
    const sellingPriceRoundedForTotal = (sellingPriceComputed !== "" && sellingPriceComputed != null) ? roundUpToThousand(Number(sellingPriceComputed)) : "";
    const totalSelling = (qty !== "" && sellingPriceRoundedForTotal !== "" && sellingPriceRoundedForTotal != null) ? Number(qty) * Number(sellingPriceRoundedForTotal) : "";
    const newRow = { name, qty: qty !== "" ? qty : "", unit, price: price !== "" ? price : "", total: totalCogs !== "" ? totalCogs : "", margin: margin !== "" ? margin : 0, sellingPrice: (sellingPriceComputed !== "" && sellingPriceComputed != null) ? Number(sellingPriceComputed) : "", totalSellingPrice: (totalSelling !== "" && totalSelling != null) ? totalSelling : "", category };
    if (editingIndex !== null && typeof editingIndex === 'number') { tableData[editingIndex] = newRow; setEditingMode(null); } else { tableData.push(newRow); }
    selectedIndex = null; sellingInputManuallyEdited = false;
    localStorage.setItem('boq_working', JSON.stringify(tableData));
    normalizeCategoryOrder();

    const storePrice = (price !== "" && price != null) ? Number(price) : "";
    const storeSelling = (sellingProvided !== "" && sellingProvided != null) ? Number(sellingProvided) : ((sellingPriceComputed !== "" && sellingPriceComputed != null) ? Number(sellingPriceComputed) : "");
    updateItemsWithEntry(name, unit || "", storePrice !== "" ? storePrice : "", storeSelling !== "" ? storeSelling : "", category || "");
    form.reset(); nameInput.focus(); renderAll();
    persistProject();
    updateEditButton();
});

editBtn.addEventListener('click', () => {
    if (editingIndex !== null) { cancelEdit(); updateEditButton(); return; }
    if (selectedIndex === null) { alert('Pilih baris dulu!'); return; }
    setEditingMode(selectedIndex);
    const r = tableData[selectedIndex];
    nameInput.value = r.name;
    qtyInput.value = r.qty !== "" ? r.qty : "";
    unitInput.value = r.unit || "";
    priceInput.value = r.price !== "" ? Number(r.price).toLocaleString('id-ID') : "";
    marginInput.value = (r.margin !== undefined && r.margin !== null && r.margin !== "") ? r.margin : "";
    sellingInputManuallyEdited = false;

    if (r.sellingPrice != null && r.sellingPrice !== "") {
        sellingInput.value = Number(r.sellingPrice).toLocaleString('id-ID');
    } else {
        const hasCOGS = r.price !== "" && r.price != null;
        const hasMargin = (r.margin !== undefined && r.margin !== null && r.margin !== "");
        if (hasCOGS && hasMargin) computeAndDisplaySelling();
        else sellingInput.value = '';
    }

    categoryInput.value = r.category || "";
    nameInput.focus();
    requestAnimationFrame(() => {
        nameInput.setSelectionRange(0, 0);
        nameInput.scrollLeft = 0;
    });
    updateEditButton();
});

deleteBtn.addEventListener('click', () => {
    if (selectedIndex === null) { alert('Pilih baris dulu!'); return; }
    if (!confirm('Hapus baris ini?')) return;
    tableData.splice(selectedIndex, 1);
    localStorage.setItem('boq_working', JSON.stringify(tableData));
    selectedIndex = null; setEditingMode(null);
    normalizeCategoryOrder();
    renderAll();
    persistProject();
    updateEditButton();
});

exportExcelBtn.addEventListener('click', exportExcel);
exportPdfBtn.addEventListener('click', exportPDF);

async function exportExcel() {
    const filename = getSafeFileName(currentProjectName) + ".xlsx";
    const workbook = new ExcelJS.Workbook();

    // SHEET: Selling
    const sheetSelling = workbook.addWorksheet("Selling");
    sheetSelling.columns = [
        { key: "no", width: 6 }, { key: "name", width: 40 }, { key: "qty", width: 10 },
        { key: "unit", width: 12 }, { key: "sellingPrice", width: 18 }, { key: "totalSelling", width: 18 }
    ];
    sheetSelling.insertRow(1, [currentProjectName ? `BOQ - ${currentProjectName}` : 'BOQ']);
    sheetSelling.insertRow(2, []);
    const lastColSell = sheetSelling.columns.length;
    sheetSelling.mergeCells(`A1:${excelColLetter(lastColSell)}1`);
    const tCellSell = sheetSelling.getRow(1).getCell(1);
    tCellSell.font = { name: 'Arial', bold: true };
    tCellSell.alignment = { horizontal: 'left', vertical: 'middle' };

    const headerSell = ["No", "Nama Item", "Qty", "Unit", "Harga Satuan", "Total"];
    sheetSelling.insertRow(3, headerSell);
    const headerRowSell = sheetSelling.getRow(3);
    headerRowSell.eachCell((cell) => {
        cell.font = { name: 'Arial', bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });

    const groupedS = {};
    tableData.forEach(r => { const cat = sanitize(r.category || "Uncategorized"); if (!groupedS[cat]) groupedS[cat] = []; groupedS[cat].push(r); });
    normalizeCategoryOrder();
    let noS = 1;
    categoryOrder.forEach(cat => {
        const secRow = sheetSelling.addRow([cat]);
        sheetSelling.mergeCells(`A${secRow.number}:${excelColLetter(lastColSell)}${secRow.number}`);
        for (let i = 1; i <= lastColSell; i++) {
            const cell = secRow.getCell(i);
            cell.font = { name: 'Arial', bold: true };
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEEEEE" } };
            cell.alignment = { horizontal: "left", vertical: "middle" };
            cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        }
        (groupedS[cat] || []).forEach(r => {
            let spRaw = null;
            if (r.sellingPrice != null && r.sellingPrice !== "") spRaw = Number(r.sellingPrice);
            else if (r.price != null && r.price !== "") spRaw = calcSellingPrice(Number(r.price), r.margin);
            const spRounded = spRaw != null ? roundUpToThousand(spRaw) : "";
            const totalSp = (spRounded !== "") ? Number(r.qty || 0) * spRounded : "";
            const rowObj = [noS++, r.name, r.qty, r.unit, (spRounded !== "" ? spRounded : ""), (totalSp !== "" ? totalSp : "")];
            const dr = sheetSelling.addRow(rowObj);
            dr.eachCell((cell, colNum) => {
                cell.font = { name: 'Arial' };
                cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
                if (colNum >= 5) { cell.numFmt = "#,##0"; cell.alignment = { horizontal: "right", vertical: "middle" }; }
                else if (colNum === 2) { cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true }; }
                else { cell.alignment = { horizontal: "center", vertical: "middle" }; }
            });
        });
    });

    const sumSelling = tableData.reduce((a, b) => {
        let sp = null;
        if (b.sellingPrice != null && b.sellingPrice !== "") sp = Number(b.sellingPrice);
        else if (b.price != null && b.price !== "") sp = calcSellingPrice(Number(b.price), b.margin);
        const spr = sp != null ? roundUpToThousand(sp) : 0;
        return a + (spr ? (Number(b.qty || 0) * spr) : 0);
    }, 0);
    const footRowSell = sheetSelling.addRow([]);
    sheetSelling.mergeCells(`A${footRowSell.number}:${excelColLetter(lastColSell - 1)}${footRowSell.number}`);
    const labelCellS = footRowSell.getCell(1);
    labelCellS.value = "Grand Total";
    labelCellS.font = { name: 'Arial', bold: true };
    labelCellS.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3D6" } };
    labelCellS.alignment = { horizontal: "center", vertical: "middle" };
    labelCellS.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    const totalCellS = footRowSell.getCell(lastColSell);
    totalCellS.value = sumSelling;
    totalCellS.numFmt = "#,##0";
    totalCellS.font = { name: 'Arial', bold: true };
    totalCellS.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3D6" } };
    totalCellS.alignment = { horizontal: "right", vertical: "middle" };
    totalCellS.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

    // SHEET: COGS
    const sheetCogs = workbook.addWorksheet("COGS");
    sheetCogs.columns = [
        { key: "no", width: 6 }, { key: "name", width: 40 }, { key: "qty", width: 10 },
        { key: "unit", width: 12 }, { key: "price", width: 18 }, { key: "totalCogs", width: 18 },
        { key: "margin", width: 12 }
    ];
    sheetCogs.insertRow(1, [currentProjectName ? `BOQ - ${currentProjectName}` : 'BOQ']);
    sheetCogs.insertRow(2, []);
    const lastColCogs = sheetCogs.columns.length;
    sheetCogs.mergeCells(`A1:${excelColLetter(lastColCogs)}1`);
    const tCellC = sheetCogs.getRow(1).getCell(1);
    tCellC.font = { name: 'Arial', bold: true };
    tCellC.alignment = { horizontal: 'left', vertical: 'middle' };

    const headerCogs = ["No", "Nama Item", "Qty", "Unit", "Harga COGS", "Total COGS", "Margin (%)"];
    sheetCogs.insertRow(3, headerCogs);
    const headerRowCogs = sheetCogs.getRow(3);
    headerRowCogs.eachCell((cell) => {
        cell.font = { name: 'Arial', bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });

    const groupedC = {};
    tableData.forEach(r => { const cat = sanitize(r.category || "Uncategorized"); if (!groupedC[cat]) groupedC[cat] = []; groupedC[cat].push(r); });
    normalizeCategoryOrder();
    let noC = 1;
    categoryOrder.forEach(cat => {
        const secRow = sheetCogs.addRow([cat]);
        sheetCogs.mergeCells(`A${secRow.number}:${excelColLetter(lastColCogs)}${secRow.number}`);
        for (let i = 1; i <= lastColCogs; i++) {
            const cell = secRow.getCell(i);
            cell.font = { name: 'Arial', bold: true };
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEEEEE" } };
            cell.alignment = { horizontal: "left", vertical: "middle" };
            cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        }
        (groupedC[cat] || []).forEach(r => {
            const totalCogs = (r.price != null && r.price !== "" && r.qty != null && r.qty !== "") ? Number(r.price) * Number(r.qty) : "";
            const rowObj = [noC++, r.name, r.qty, r.unit, (r.price !== "" && r.price != null) ? Number(r.price) : "", (totalCogs !== "" ? totalCogs : ""), (r.margin !== undefined && r.margin != null) ? r.margin : ""];
            const dr = sheetCogs.addRow(rowObj);
            dr.eachCell((cell, colNum) => {
                cell.font = { name: 'Arial' };
                cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
                if (colNum >= 5 && colNum <= 6) { cell.numFmt = "#,##0"; cell.alignment = { horizontal: "right", vertical: "middle" }; }
                else if (colNum === 2) { cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true }; }
                else { cell.alignment = { horizontal: "center", vertical: "middle" }; }
            });
        });
    });

    const sumCogs = tableData.reduce((a, b) => a + (Number(b.price || 0) * Number(b.qty || 0)), 0);
    const footRowCogs = sheetCogs.addRow([]);
    sheetCogs.mergeCells(`A${footRowCogs.number}:${excelColLetter(lastColCogs - 1)}${footRowCogs.number}`);
    const labelCellC = footRowCogs.getCell(1);
    labelCellC.value = "Grand Total COGS";
    labelCellC.font = { name: 'Arial', bold: true };
    labelCellC.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3D6" } };
    labelCellC.alignment = { horizontal: "center", vertical: "middle" };
    labelCellC.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    const totalCellC = footRowCogs.getCell(lastColCogs);
    totalCellC.value = sumCogs;
    totalCellC.numFmt = "#,##0";
    totalCellC.font = { name: 'Arial', bold: true };
    totalCellC.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3D6" } };
    totalCellC.alignment = { horizontal: "right", vertical: "middle" };
    totalCellC.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

    // SHEET: Calculation
    const sheetCalc = workbook.addWorksheet("Calculation");
    sheetCalc.columns = [
        { key: "totalCogs", width: 18 }, { key: "totalSelling", width: 18 },
        { key: "commission", width: 18 }, { key: "grossProfit", width: 18 }, { key: "totalMargin", width: 18 }
    ];
    sheetCalc.insertRow(1, [currentProjectName ? `BOQ - ${currentProjectName}` : 'BOQ']);
    sheetCalc.insertRow(2, []);
    const lastColCalc = sheetCalc.columns.length;
    sheetCalc.mergeCells(`A1:${excelColLetter(lastColCalc)}1`);
    const tCellCalc = sheetCalc.getRow(1).getCell(1);
    tCellCalc.font = { name: 'Arial', bold: true };
    tCellCalc.alignment = { horizontal: 'left', vertical: 'middle' };

    const headerCalc = ["Total COGS", "Total Selling", "Commission", "Gross Profit", "Total Margin"];
    sheetCalc.insertRow(3, headerCalc);
    const headerRowCalc = sheetCalc.getRow(3);
    headerRowCalc.eachCell((cell) => {
        cell.font = { name: 'Arial', bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });

    const totals = computeTotals();
    const totalCogsVal = Number(totals.totalCogs || 0);
    const totalSellingVal = Number(totals.totalSelling || 0);
    const commissionVal = Number(currentCommission || 0);
    const grossProfitVal = totalSellingVal - totalCogsVal - commissionVal;
    const totalMarginPercent = (totalSellingVal === 0) ? 0 : (grossProfitVal / totalSellingVal * 100);

    const dataRowCalc = sheetCalc.addRow([totalCogsVal, totalSellingVal, commissionVal, grossProfitVal, (isFinite(totalMarginPercent) ? totalMarginPercent.toFixed(2) + '%' : '0.00%')]);
    dataRowCalc.eachCell((cell, colNum) => {
        cell.font = { name: 'Arial' };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        if (colNum <= 4) { cell.numFmt = "#,##0"; cell.alignment = { horizontal: "right", vertical: "middle" }; }
        else { cell.alignment = { horizontal: "center", vertical: "middle" }; }
    });

    const footCalc = sheetCalc.addRow([]);
    sheetCalc.mergeCells(`A${footCalc.number}:${excelColLetter(lastColCalc)}${footCalc.number}`);
    const labelCellCalc = footCalc.getCell(1);
    labelCellCalc.value = "Calculation Summary";
    labelCellCalc.font = { name: 'Arial', bold: true };
    labelCellCalc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3D6" } };
    labelCellCalc.alignment = { horizontal: "center", vertical: "middle" };
    labelCellCalc.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
}

function exportPDF() {
    if (!tableData.length && activeTab !== 'calculation') { alert('Tidak ada data untuk diekspor.'); return; }
    let filenameSuffix = '';
    if (activeTab === 'cogs') filenameSuffix = ' (COGS)';
    else if (activeTab === 'selling') filenameSuffix = ' (Selling)';
    else if (activeTab === 'calculation') filenameSuffix = ' (Calculation)';
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    doc.setFontSize(14);
    const title = currentProjectName ? `BOQ - ${currentProjectName}` : 'BOQ';
    doc.text(title, 14, 15);

    const grouped = {};
    tableData.forEach(r => {
        const cat = sanitize(r.category || 'Uncategorized');
        if (!grouped[cat]) grouped[cat] = [];
        grouped[cat].push(r);
    });
    normalizeCategoryOrder();
    const body = []; let rowNoGlobal = 1;

    if (activeTab === 'cogs') {
        categoryOrder.forEach(cat => {
            body.push([{ content: cat, colSpan: 7, styles: { fillColor: [238, 238, 238], fontStyle: "bold", halign: "left", valign: "middle" } }]);
            (grouped[cat] || []).forEach(r => {
                const totalCogs = (r.price != null && r.price !== "" && r.qty != null && r.qty !== "") ? Number(r.price) * Number(r.qty) : "";
                body.push([rowNoGlobal++, r.name, r.qty, r.unit,
                { content: r.price != null && r.price !== "" ? Number(r.price).toLocaleString('id-ID') : "", styles: { halign: 'right', valign: 'middle' } },
                { content: totalCogs !== "" ? Number(totalCogs).toLocaleString('id-ID') : "", styles: { halign: 'right', valign: 'middle' } },
                { content: r.margin != null && r.margin !== "" ? Number(r.margin).toFixed(2) : "", styles: { halign: 'right', valign: 'middle' } }
                ]);
            });
        });
        const grandCogs = tableData.reduce((a, b) => a + (Number(b.price || 0) * Number(b.qty || 0)), 0);
        body.push([{ content: "Grand Total COGS", colSpan: 6, styles: { fontStyle: "bold", halign: "center", valign: "middle", fillColor: [254, 243, 214] } }, { content: grandCogs.toLocaleString('id-ID'), styles: { fontStyle: "bold", halign: "right", valign: "middle", fillColor: [254, 243, 214] } }]);
        const head = ["No", "Nama Item", "Qty", "Unit", "Harga COGS", "Total COGS", "Margin (%)"];
        doc.autoTable({ head: [head], body, startY: 22, theme: 'grid', headStyles: { fillColor: [31, 78, 120], textColor: [255, 255, 255], halign: 'center', valign: 'middle' }, columnStyles: { 0: { halign: 'center' }, 2: { halign: 'center' }, 3: { halign: 'center' } }, styles: { fontSize: 10, cellPadding: 3 } });
    }
    else if (activeTab === 'selling') {
        categoryOrder.forEach(cat => {
            body.push([{ content: cat, colSpan: 6, styles: { fillColor: [238, 238, 238], fontStyle: "bold", halign: "left", valign: "middle" } }]);
            (grouped[cat] || []).forEach(r => {
                let spRaw = null;
                if (r.sellingPrice != null && r.sellingPrice !== "") spRaw = Number(r.sellingPrice);
                else if (r.price != null && r.price !== "") spRaw = calcSellingPrice(Number(r.price), r.margin);
                const spRounded = spRaw != null ? roundUpToThousand(spRaw) : "";
                const totalSp = (spRounded !== "") ? Number(r.qty || 0) * spRounded : "";
                body.push([rowNoGlobal++, r.name, r.qty, r.unit,
                { content: spRounded !== "" ? Number(spRounded).toLocaleString('id-ID') : "", styles: { halign: 'right', valign: 'middle' } },
                { content: totalSp !== "" ? Number(totalSp).toLocaleString('id-ID') : "", styles: { halign: 'right', valign: 'middle' } }
                ]);
            });
        });
        const grandSelling = tableData.reduce((a, b) => {
            let sp = null;
            if (b.sellingPrice != null && b.sellingPrice !== "") sp = Number(b.sellingPrice);
            else if (b.price != null && b.price !== "") sp = calcSellingPrice(Number(b.price), b.margin);
            const spr = sp != null ? roundUpToThousand(sp) : 0;
            return a + (spr ? (Number(b.qty || 0) * spr) : 0);
        }, 0);
        body.push([{ content: "Grand Total", colSpan: 5, styles: { fontStyle: "bold", halign: "center", valign: "middle", fillColor: [254, 243, 214] } }, { content: grandSelling.toLocaleString('id-ID'), styles: { fontStyle: "bold", halign: "right", valign: "middle", fillColor: [254, 243, 214] } }]);
        const head = ["No", "Nama Item", "Qty", "Unit", "Harga Satuan", "Total"];
        doc.autoTable({ head: [head], body, startY: 22, theme: 'grid', headStyles: { fillColor: [31, 78, 120], textColor: [255, 255, 255], halign: 'center', valign: 'middle' }, columnStyles: { 0: { halign: 'center' }, 2: { halign: 'center' }, 3: { halign: 'center' } }, styles: { fontSize: 10, cellPadding: 3 } });
    }
    else if (activeTab === 'calculation') {
        const totals = computeTotals();
        const totalCogs = Number(totals.totalCogs || 0);
        const totalSelling = Number(totals.totalSelling || 0);
        const commission = Number(currentCommission || 0);
        const grossProfit = totalSelling - totalCogs - commission;
        const totalMargin = (totalSelling === 0) ? 0 : (grossProfit / totalSelling);
        const head = ["Total COGS", "Total Selling", "Commission", "Gross Profit", "Total Margin"];
        const row = [
            { content: Number(totalCogs).toLocaleString('id-ID'), styles: { halign: 'right', valign: 'middle' } },
            { content: Number(totalSelling).toLocaleString('id-ID'), styles: { halign: 'right', valign: 'middle' } },
            { content: Number(commission).toLocaleString('id-ID'), styles: { halign: 'right', valign: 'middle' } },
            { content: Number(grossProfit).toLocaleString('id-ID'), styles: { halign: 'right', valign: 'middle' } },
            { content: (isFinite(totalMargin) ? (totalMargin * 100).toFixed(2) + '%' : '0.00%'), styles: { halign: 'center', valign: 'middle' } }
        ];
        doc.autoTable({ head: [head], body: [row], startY: 22, theme: 'grid', headStyles: { fillColor: [31, 78, 120], textColor: [255, 255, 255], halign: 'center', valign: 'middle' }, styles: { fontSize: 10, cellPadding: 3 } });
    }
    const filename = getSafeFileName(currentProjectName) + (filenameSuffix ? filenameSuffix : '') + ".pdf";
    doc.save(filename);
}

// MANAGE ITEMS & PROJECTS (tanpa deklarasi ganda)
manageBtn.addEventListener('click', () => {
    manageItemSearch.value = '';
    populateManageItems();
    manageModal.style.display = 'block';
    manageModal.setAttribute('aria-hidden', 'false');
    manageItemSearch.focus();
});
modalClose.addEventListener('click', () => { manageModal.style.display = 'none'; manageModal.setAttribute('aria-hidden', 'true'); });
projectsClose.addEventListener('click', () => { projectsModal.style.display = 'none'; projectsModal.setAttribute('aria-hidden', 'true'); });
manageProjectsBtn.addEventListener('click', () => {
    manageProjectSearch.value = '';
    populateManageProjects();
    projectsModal.style.display = 'block';
    projectsModal.setAttribute('aria-hidden', 'false');
    manageProjectSearch.focus();
});

function populateManageItems() {
    manageItemList.innerHTML = '';
    const q = String(manageItemSearch.value || '').trim().toLowerCase();
    const map = {};
    items.forEach(it => { if (!it || !it.name) return; const key = it.name.trim(); if (!map[key] || (it.updatedAt && map[key].updatedAt < it.updatedAt)) map[key] = it; });
    const arr = Object.values(map).slice().sort((a, b) => (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase()));
    const filtered = arr.filter(it => !q || (it.name || '').toLowerCase().includes(q));
    filtered.forEach((it) => {
        const li = document.createElement('li');
        const span = document.createElement('span'); span.textContent = it.name;
        const renameBtnEl = document.createElement('button'); renameBtnEl.textContent = 'Rename'; renameBtnEl.className = 'btn ghost';
        renameBtnEl.addEventListener('click', () => {
            const newName = sanitize(prompt('Ubah nama item:', it.name) || '');
            if (!newName) return;
            const oldKey = normalizeKey(it.name);
            items.forEach(x => { if (normalizeKey(x.name) === oldKey) x.name = newName; });
            localStorage.setItem('boq_items', JSON.stringify(items));
            updateDatalist();
            touchCloudSync();
            populateManageItems();
        });
        const delBtnEl = document.createElement('button'); delBtnEl.textContent = 'Hapus'; delBtnEl.className = 'btn secondary';
        delBtnEl.addEventListener('click', () => {
            if (!confirm(`Hapus item "${it.name}"?`)) return;
            const key = normalizeKey(it.name);
            items = items.filter(x => normalizeKey(x.name) !== key);
            localStorage.setItem('boq_items', JSON.stringify(items));
            updateDatalist();
            touchCloudSync();
            populateManageItems();
        });
        li.appendChild(span); li.appendChild(renameBtnEl); li.appendChild(delBtnEl);
        manageItemList.appendChild(li);
    });
}

function populateManageProjects() {
    manageProjectList.innerHTML = '';
    const q = String(manageProjectSearch.value || '').trim().toLowerCase();
    const projects = loadProjectsObj();
    const names = Object.keys(projects || {}).slice().sort((a, b) => (a || '').toLowerCase().localeCompare((b || '').toLowerCase()));
    const filtered = names.filter(n => !q || (n || '').toLowerCase().includes(q));
    filtered.forEach(name => {
        const li = document.createElement('li');
        const span = document.createElement('span'); span.textContent = name;
        const delBtnEl = document.createElement('button'); delBtnEl.textContent = 'Hapus'; delBtnEl.className = 'btn secondary';
        delBtnEl.addEventListener('click', () => {
            if (!confirm(`Hapus project "${name}"?`)) return;
            delete projects[name];
            saveProjectsObj(projects);
            touchCloudSync();
            populateManageProjects();
            updateProjectList();
            updateFooterLastSaved();
        });
        const renameBtnEl = document.createElement('button'); renameBtnEl.textContent = 'Rename'; renameBtnEl.className = 'btn ghost';
        renameBtnEl.addEventListener('click', () => {
            const newName = sanitize(prompt('Ubah nama project:', name) || '');
            if (!newName) return;
            if (projects[newName]) { alert('Nama project sudah ada.'); return; }
            projects[newName] = projects[name];
            projects[newName].lastSaved = Date.now();
            delete projects[name];
            saveProjectsObj(projects);
            touchCloudSync();
            if (currentProjectName === name) { currentProjectName = newName; localStorage.setItem('boq_current_name', currentProjectName); updateNavProject(); updateFooterLastSaved(); }
            populateManageProjects(); updateProjectList();
        });
        li.appendChild(span); li.appendChild(renameBtnEl); li.appendChild(delBtnEl);
        manageProjectList.appendChild(li);
    });
}

window.addEventListener('click', (e) => {
    if (e.target === manageModal) { manageModal.style.display = 'none'; manageModal.setAttribute('aria-hidden', 'true'); }
    if (e.target === projectsModal) { projectsModal.style.display = 'none'; projectsModal.setAttribute('aria-hidden', 'true'); }
});

function updateProjectList() {
    const projects = loadProjectsObj();
    projectList.innerHTML = '';
    const keys = Object.keys(projects || {}).slice().sort((a, b) => (a || '').toLowerCase().localeCompare((b || '').toLowerCase()));
    keys.forEach(k => { const opt = document.createElement('option'); opt.value = k; projectList.appendChild(opt); });
}

manageItemSearch.addEventListener('input', populateManageItems);
manageProjectSearch.addEventListener('input', populateManageProjects);

saveProjectBtn.addEventListener('click', saveProject);
loadProjectBtn.addEventListener('click', () => loadProject());
newProjectBtn.addEventListener('click', () => newProject());

function saveProject() {
    if (!tableData.length) { alert('Tidak ada data untuk disimpan.'); return; }
    let name = prompt('Masukkan nama project:', currentProjectName || '');
    if (!name) return;
    name = sanitize(name);
    const projects = loadProjectsObj();
    if (projects[name] && !confirm('Nama project sudah ada. Timpa?')) return;
    projects[name] = {
        data: tableData,
        commission: Number(currentCommission || 0),
        categoryOrder: categoryOrder.slice(),
        lastSaved: Date.now()
    };
    saveProjectsObj(projects);
    currentProjectName = name; localStorage.setItem('boq_current_name', currentProjectName);
    updateProjectList();
    updateNavProject();
    updateFooterLastSaved();
    touchCloudSync();
    alert(`Project "${name}" berhasil disimpan.`);
}

function loadProject() {
    const name = sanitize(projectSearch.value || '');
    if (!name) { alert('Masukkan nama project terlebih dahulu.'); return; }
    const projects = loadProjectsObj();
    if (!projects[name]) { alert('Project tidak ditemukan.'); return; }
    const stored = projects[name];
    if (Array.isArray(stored)) { tableData = stored; currentCommission = 0; categoryOrder = []; }
    else if (stored && typeof stored === 'object') {
        tableData = stored.data || [];
        currentCommission = Number(stored.commission || 0);
        categoryOrder = (stored.categoryOrder && Array.isArray(stored.categoryOrder)) ? stored.categoryOrder.slice() : [];
    } else { tableData = []; currentCommission = 0; categoryOrder = []; }
    form.reset(); sellingInputManuallyEdited = false; setEditingMode(null); selectedIndex = null;
    localStorage.setItem('boq_working', JSON.stringify(tableData));
    localStorage.setItem('boq_category_order', JSON.stringify(categoryOrder));
    currentProjectName = name; localStorage.setItem('boq_current_name', currentProjectName);
    updateNavProject(); updateProjectList(); renderAll(); projectSearch.value = '';
    updateFooterLastSaved();
    touchCloudSync();
    alert(`Project "${name}" berhasil dimuat.`);
    updateEditButton();
}

function newProject() {
    if (!confirm('Buat project baru? Semua data saat ini akan dihapus.')) return;
    tableData = []; selectedIndex = null; setEditingMode(null); currentProjectName = ''; currentCommission = 0;
    categoryOrder = [];
    form.reset(); sellingInputManuallyEdited = false;
    localStorage.setItem('boq_working', JSON.stringify(tableData));
    localStorage.setItem('boq_category_order', JSON.stringify(categoryOrder));
    localStorage.removeItem('boq_current_name');
    updateNavProject(); renderAll(); updateEditButton();
    updateFooterLastSaved();
    touchCloudSync();
    alert('Project baru siap diisi.');
}

renameProjectBtn.addEventListener('click', () => {
    if (!currentProjectName) { alert('Tidak ada project aktif. Simpan project dulu agar bisa rename.'); return; }
    const newName = sanitize(prompt('Ubah nama project:', currentProjectName) || '');
    if (!newName) return;
    const projects = loadProjectsObj();
    if (projects[newName] && newName !== currentProjectName && !confirm('Nama project sudah ada. Timpa?')) return;
    if (projects[currentProjectName]) delete projects[currentProjectName];
    projects[newName] = {
        data: tableData,
        commission: Number(currentCommission || 0),
        categoryOrder: categoryOrder.slice(),
        lastSaved: Date.now()
    };
    saveProjectsObj(projects);
    currentProjectName = newName;
    localStorage.setItem('boq_current_name', currentProjectName);
    updateProjectList();
    updateNavProject();
    updateFooterLastSaved();
    touchCloudSync();
    alert('Nama project diubah.');
});

// Backup / Restore
downloadJSONBtn.addEventListener('click', () => {
    const payload = {
        exportedAt: new Date().toISOString(),
        ...serializeAppState()
    };

    const ts = new Date();
    const pad = n => String(n).padStart(2, '0');
    const filename = `boq_backup_${ts.getFullYear()}${pad(ts.getMonth() + 1)}${pad(ts.getDate())}_${pad(ts.getHours())}${pad(ts.getMinutes())}.json`;

    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
});

restoreJSONBtn.addEventListener('click', () => { restoreInput.value = ''; restoreInput.click(); });
restoreInput.addEventListener('change', (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        let parsed = null;
        try { parsed = JSON.parse(String(e.target.result)); } catch (err) { alert('File tidak valid: bukan JSON.'); return; }
        let restoreObj = { projects: {}, items: [], working: [], currentProjectName: '', unsavedCommission: 0, categoryOrder: [] };
        if (parsed && typeof parsed === 'object') {
            if (parsed.projects && typeof parsed.projects === 'object') {
                restoreObj.projects = parsed.projects || {};
                restoreObj.items = Array.isArray(parsed.items) ? parsed.items : (parsed.items || []);
                restoreObj.working = Array.isArray(parsed.working) ? parsed.working : (parsed.working || []);
                restoreObj.currentProjectName = parsed.currentProjectName || (parsed.project || '') || '';
                restoreObj.unsavedCommission = Number(parsed.unsavedCommission || parsed.commission || 0);
                restoreObj.categoryOrder = Array.isArray(parsed.categoryOrder) ? parsed.categoryOrder : [];
            } else if (parsed.project && parsed.data) {
                const projName = sanitize(parsed.project) || ('Imported Project ' + new Date().toISOString());
                restoreObj.projects = loadProjectsObj();
                restoreObj.projects[projName] = { data: Array.isArray(parsed.data) ? parsed.data : [], commission: Number(parsed.commission || 0), categoryOrder: Array.isArray(parsed.categoryOrder) ? parsed.categoryOrder : [], lastSaved: Date.now() };
                restoreObj.items = items || [];
                restoreObj.working = restoreObj.projects[projName].data;
                restoreObj.currentProjectName = projName;
                restoreObj.unsavedCommission = Number(parsed.commission || 0);
                restoreObj.categoryOrder = restoreObj.projects[projName].categoryOrder || [];
            } else {
                restoreObj.projects = parsed.projects || parsed.projectList || {};
                restoreObj.items = parsed.items || parsed.boq_items || [];
                restoreObj.working = parsed.working || parsed.boq_working || [];
                restoreObj.currentProjectName = parsed.currentProjectName || parsed.current || '';
                restoreObj.unsavedCommission = Number(parsed.unsavedCommission || parsed.commission || 0);
                restoreObj.categoryOrder = Array.isArray(parsed.categoryOrder) ? parsed.categoryOrder : [];
            }
        }
        if (!confirm('Restore akan menimpa semua data lokal (projects, items, working, category order, current project). Lanjutkan?')) { return; }
        try {
            applyAppState({
                projects: restoreObj.projects || {},
                items: restoreObj.items || [],
                working: restoreObj.working || [],
                currentProjectName: restoreObj.currentProjectName || '',
                categoryOrder: restoreObj.categoryOrder || [],
                unsavedCommission: Number(restoreObj.unsavedCommission || 0),
                meta: {
                    schemaVersion: 1
                }
            });

            touchCloudSync();
            alert('Restore berhasil. Data lokal telah diperbarui.');
        } catch (err) {
            console.error(err);
            alert('Terjadi kesalahan saat menyimpan data ke localStorage.');
        }
    };
    reader.readAsText(f);
});

// INIT
function init() {
    updateHeightsVars();
    updateProjectList();
    updateDatalist();
    loadCommissionForCurrentProject();
    normalizeCategoryOrder();
    updateReorderButton();
    renderAll();
    switchTo('selling');
    updateEditBtnInit();
    updateFooterStatus();
    updateNavProject();
}
init();

/* ===== USER-SCOPED LOCAL STORAGE PATCH ===== */

(() => {
    const SCOPED_KEYS = new Set([
        'boq_projects_v2',
        'boq_unsaved_commission',
        'boq_sync_meta_v1',
        'boq_items',
        'boq_working',
        'boq_current_name',
        'boq_category_order'
    ]);

    const LEGACY_TO_SUFFIX = {
        boq_projects_v2: 'projects',
        boq_unsaved_commission: 'unsaved_commission',
        boq_sync_meta_v1: 'sync_meta',
        boq_items: 'items',
        boq_working: 'working',
        boq_current_name: 'current_name',
        boq_category_order: 'category_order'
    };

    const originalGetItem = localStorage.getItem.bind(localStorage);
    const originalSetItem = localStorage.setItem.bind(localStorage);
    const originalRemoveItem = localStorage.removeItem.bind(localStorage);

    let scopedStorageUserId = null;

    function setScopedStorageUserId(userId) {
        scopedStorageUserId = userId || null;
    }

    function getScopedNamespace(userId = scopedStorageUserId) {
        return userId ? `boq:user:${userId}` : 'boq:guest';
    }

    function getScopedKeyFromLegacyKey(legacyKey, userId = scopedStorageUserId) {
        const suffix = LEGACY_TO_SUFFIX[legacyKey];
        if (!suffix) return legacyKey;
        return `${getScopedNamespace(userId)}:${suffix}`;
    }

    function readRawJson(rawKey, fallback) {
        try {
            const raw = originalGetItem(rawKey);
            return raw === null ? fallback : JSON.parse(raw);
        } catch (e) {
            return fallback;
        }
    }

    function legacyGlobalSnapshotExists() {
        return Object.keys(LEGACY_TO_SUFFIX).some((legacyKey) => originalGetItem(legacyKey) !== null);
    }

    function scopedSnapshotExists(userId = scopedStorageUserId) {
        return Object.keys(LEGACY_TO_SUFFIX).some((legacyKey) => {
            const scopedKey = getScopedKeyFromLegacyKey(legacyKey, userId);
            return originalGetItem(scopedKey) !== null;
        });
    }

    function migrateLegacyGlobalToUser(userId = scopedStorageUserId) {
        if (!userId) return;
        if (scopedSnapshotExists(userId)) return;
        if (!legacyGlobalSnapshotExists()) return;

        Object.keys(LEGACY_TO_SUFFIX).forEach((legacyKey) => {
            const raw = originalGetItem(legacyKey);
            if (raw !== null) {
                const scopedKey = getScopedKeyFromLegacyKey(legacyKey, userId);
                originalSetItem(scopedKey, raw);
            }
        });
    }

    function hydrateRuntimeFromScopedStorage() {
        items = readRawJson(getScopedKeyFromLegacyKey('boq_items'), []);
        tableData = readRawJson(getScopedKeyFromLegacyKey('boq_working'), []);
        currentProjectName = originalGetItem(getScopedKeyFromLegacyKey('boq_current_name')) || '';
        categoryOrder = readRawJson(getScopedKeyFromLegacyKey('boq_category_order'), []);

        selectedIndex = null;
        editingIndex = null;
        sellingInputManuallyEdited = false;

        if (typeof loadCommissionForCurrentProject === 'function') {
            loadCommissionForCurrentProject();
        }

        if (typeof setEditingMode === 'function') {
            setEditingMode(null);
        }

        if (typeof updateProjectList === 'function') updateProjectList();
        if (typeof updateDatalist === 'function') updateDatalist();
        if (typeof updateNavProject === 'function') updateNavProject();
        if (typeof renderAll === 'function') renderAll();
        if (typeof updateFooterLastSaved === 'function') updateFooterLastSaved();
        if (typeof updateEditButton === 'function') updateEditButton();
    }

    function patchLocalStorageMethods() {
        localStorage.getItem = function (key) {
            if (SCOPED_KEYS.has(key)) {
                return originalGetItem(getScopedKeyFromLegacyKey(key));
            }
            return originalGetItem(key);
        };

        localStorage.setItem = function (key, value) {
            if (SCOPED_KEYS.has(key)) {
                return originalSetItem(getScopedKeyFromLegacyKey(key), value);
            }
            return originalSetItem(key, value);
        };

        localStorage.removeItem = function (key) {
            if (SCOPED_KEYS.has(key)) {
                return originalRemoveItem(getScopedKeyFromLegacyKey(key));
            }
            return originalRemoveItem(key);
        };
    }

    async function bootstrapScopedStorageFromSession() {
        const session = await getStoredSession();
        const userId = session?.user?.id || null;

        setScopedStorageUserId(userId);

        if (userId) {
            migrateLegacyGlobalToUser(userId);
        }

        hydrateRuntimeFromScopedStorage();
    }

    patchLocalStorageMethods();

    supabaseClient.auth.onAuthStateChange(async (_event, session) => {
        const userId = session?.user?.id || null;
        setScopedStorageUserId(userId);

        if (userId) {
            migrateLegacyGlobalToUser(userId);
        }

        hydrateRuntimeFromScopedStorage();
    });

    bootstrapScopedStorageFromSession();
})();