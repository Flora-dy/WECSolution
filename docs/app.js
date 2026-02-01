const DEFAULT_LANG = "CN";
const STORAGE_LANG = "wecare_pages_lang";
const STORAGE_VIEW = "wecare_pages_view";
const STORAGE_OK = "wecare_pages_authed";

// Lightweight client-side password gate (not secure).
const SHARED_PASSWORD = "WECARE888";

const el = (id) => document.getElementById(id);

function setSelectedSeg(container, value) {
  container.querySelectorAll(".seg-btn").forEach((btn) => {
    btn.setAttribute("aria-selected", btn.dataset.lang === value || btn.dataset.view === value ? "true" : "false");
  });
}

function setHero(lang) {
  const title = el("hero-title");
  const lines = el("hero-lines");
  if (lang === "EN") {
    title.textContent = "WECARE Health & Wellness Solutions";
    lines.innerHTML = [
      `<div class="hero-strong">Tailored for Global Brands</div>`,
      `<div>Built on scientific rigor, consumer insights, and global expertise.</div>`,
      `<div class="hero-spacer">From proprietary probiotic strains to advanced formulations and end-to-end delivery.</div>`,
    ].join("");
    el("label-cat").textContent = "Health Area";
    el("label-sub").textContent = "Use Case";
    el("viewer-title").textContent = "Full Solution";
    el("viewer-sub").textContent = "Two-page PDF preview";
    el("viewer-foot").textContent = "Tip: On mobile, PDFs may open in the browser’s built-in viewer.";
  } else {
    title.textContent = "WECARE 健康与营养解决方案";
    lines.innerHTML = [
      `<div class="hero-strong">服务全球品牌</div>`,
      `<div>以科学为基础，以洞察为导向，以专业能力贯穿全流程</div>`,
      `<div class="hero-spacer">从自有益生菌菌株到配方开发与商业化落地</div>`,
    ].join("");
    el("label-cat").textContent = "功能方向";
    el("label-sub").textContent = "应用场景";
    el("viewer-title").textContent = "完整解决方案";
    el("viewer-sub").textContent = "PDF 双页预览";
    el("viewer-foot").textContent = "提示：手机端可能会调用浏览器内置 PDF 阅读器。";
  }
}

function safeText(s) {
  return String(s ?? "").trim();
}

function getLabel(obj, lang) {
  if (!obj) return "";
  if (typeof obj === "string") return obj;
  return safeText(obj[lang] ?? obj.CN ?? obj.EN ?? "");
}

function buildPdfUrl(pdfPath, page) {
  const p = Math.max(1, Number(page || 1));
  // `#page=` works in Chrome’s built-in PDF viewer.
  return `${pdfPath}#page=${p}`;
}

async function loadData() {
  const resp = await fetch("./data/solutions.json", { cache: "no-store" });
  if (!resp.ok) throw new Error("Failed to load data");
  return await resp.json();
}

function maybeGate() {
  const ok = sessionStorage.getItem(STORAGE_OK) === "1";
  const overlay = el("pw-overlay");
  if (ok) {
    overlay.hidden = true;
    return;
  }
  overlay.hidden = false;
  const input = el("pw-input");
  const btn = el("pw-btn");
  const err = el("pw-err");

  const submit = () => {
    const v = safeText(input.value);
    if (v === SHARED_PASSWORD) {
      sessionStorage.setItem(STORAGE_OK, "1");
      overlay.hidden = true;
      return;
    }
    err.hidden = false;
  };

  btn.onclick = () => submit();
  input.onkeydown = (e) => {
    if (e.key === "Enter") submit();
  };
}

function init() {
  maybeGate();

  const lang = (localStorage.getItem(STORAGE_LANG) || DEFAULT_LANG).toUpperCase() === "EN" ? "EN" : "CN";
  const view = (localStorage.getItem(STORAGE_VIEW) || "both").toLowerCase();

  const langSeg = document.querySelector(".hero-actions .seg");
  const viewSeg = document.querySelector(".viewer-actions .seg");

  setSelectedSeg(langSeg, lang);
  setSelectedSeg(viewSeg, view);
  setHero(lang);

  Promise.resolve()
    .then(() => loadData())
    .then((data) => mount(data, lang, view))
    .catch((e) => {
      console.error(e);
      el("pill").textContent = "Data load failed";
    });

  langSeg.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-lang]");
    if (!btn) return;
    const next = btn.dataset.lang === "EN" ? "EN" : "CN";
    localStorage.setItem(STORAGE_LANG, next);
    location.reload();
  });

  viewSeg.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-view]");
    if (!btn) return;
    const next = btn.dataset.view;
    localStorage.setItem(STORAGE_VIEW, next);
    location.reload();
  });
}

function mount(data, lang, view) {
  const categories = Array.isArray(data.categories) ? data.categories : [];
  const pdfPath = safeText(data?.pdf?.[lang]) || safeText(data?.pdf?.CN);
  const catSel = el("cat");
  const subSel = el("sub");
  const pill = el("pill");

  const catOptions = categories.map((c) => ({
    key: safeText(c.key),
    label: getLabel(c.label, lang) || safeText(c.key),
    accent1: safeText(c.accent1),
    accent2: safeText(c.accent2),
    scenarios: Array.isArray(c.scenarios) ? c.scenarios : [],
  }));

  catSel.innerHTML = catOptions
    .map((c) => `<option value="${encodeURIComponent(c.key)}">${c.label}</option>`)
    .join("");

  function setTheme(c) {
    if (c?.accent1) document.documentElement.style.setProperty("--accent1", c.accent1);
    if (c?.accent2) document.documentElement.style.setProperty("--accent2", c.accent2);
  }

  function fillSubs(catKey) {
    const cat = catOptions.find((x) => x.key === catKey) || catOptions[0];
    const scenarios = cat?.scenarios || [];
    setTheme(cat);

    subSel.innerHTML = scenarios
      .map((s) => {
        const k = safeText(s.key);
        const label = getLabel(s.label, lang) || k;
        return `<option value="${encodeURIComponent(k)}">${label}</option>`;
      })
      .join("");

    if (scenarios.length) {
      subSel.value = encodeURIComponent(safeText(scenarios[0].key));
    }

    return { cat, scenarios };
  }

  function render(catKey, scenKey) {
    const cat = catOptions.find((x) => x.key === catKey) || catOptions[0];
    const scenarios = cat?.scenarios || [];
    const scen = scenarios.find((s) => safeText(s.key) === scenKey) || scenarios[0];
    const catLabel = getLabel(cat?.label, lang) || catKey;
    const scenLabel = getLabel(scen?.title, lang) || getLabel(scen?.label, lang) || scenKey;
    pill.textContent = `${catLabel} · ${scenLabel}`;

    const p1 = Number(scen?.page1 || 1);
    const p2 = Number(scen?.page2 || p1 + 1);
    const frame1 = el("frame1");
    const frame2 = el("frame2");
    const frame2Wrap = el("frame2-wrap");

    const url1 = buildPdfUrl(pdfPath, p1);
    const url2 = buildPdfUrl(pdfPath, p2);

    frame1.src = url1;
    frame2.src = url2;

    const openPdf = el("open-pdf");
    openPdf.href = url1;

    if (view === "left") {
      frame2Wrap.style.display = "none";
    } else if (view === "right") {
      frame1.src = url2;
      frame2Wrap.style.display = "none";
      openPdf.href = url2;
    } else {
      frame2Wrap.style.display = "";
    }
  }

  const initialCatKey = catOptions[0]?.key || "";
  catSel.value = encodeURIComponent(initialCatKey);
  const { cat } = fillSubs(initialCatKey);
  const firstScenario = (cat?.scenarios || [])[0];
  const initialScenKey = safeText(firstScenario?.key || "");
  render(initialCatKey, initialScenKey);

  catSel.addEventListener("change", () => {
    const k = decodeURIComponent(catSel.value || "");
    const r = fillSubs(k);
    const first = (r.cat?.scenarios || [])[0];
    render(k, safeText(first?.key || ""));
  });

  subSel.addEventListener("change", () => {
    const catKey = decodeURIComponent(catSel.value || "");
    const scenKey = decodeURIComponent(subSel.value || "");
    render(catKey, scenKey);
  });
}

init();
