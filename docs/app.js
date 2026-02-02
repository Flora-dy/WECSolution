const DEFAULT_LANG = "CN";
const DEFAULT_SERIES = "WecLac";
const DEFAULT_VIEW = "both";

const STORAGE_LANG = "wecare_pages_lang";
const STORAGE_SERIES = "wecare_pages_series";
const STORAGE_VIEW = "wecare_pages_view";

const el = (id) => document.getElementById(id);

function safeText(v) {
  return String(v ?? "").trim();
}

function setSelected(container, predicate) {
  container.querySelectorAll(".seg-btn").forEach((btn) => {
    btn.setAttribute("aria-selected", predicate(btn) ? "true" : "false");
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
    el("solution-core-title").textContent = "Core Formula";
    el("solution-clinical-title").textContent = "Clinical Studies";
    el("solution-spec-title").textContent = "Specifications";
    el("viewer-title").textContent = "Full Solution";
    el("viewer-sub").textContent = "Two-page preview";
    el("open-pdf").textContent = "Open PDF";
  } else {
    title.textContent = "WECARE 健康与营养解决方案";
    lines.innerHTML = [
      `<div class="hero-strong">服务全球品牌</div>`,
      `<div>以科学为基础，以洞察为导向，以专业能力贯穿全流程</div>`,
      `<div class="hero-spacer">从自有益生菌菌株到配方开发与商业化落地</div>`,
    ].join("");
    el("label-cat").textContent = "功能方向";
    el("label-sub").textContent = "应用场景";
    el("solution-core-title").textContent = "核心配方";
    el("solution-clinical-title").textContent = "临床研究";
    el("solution-spec-title").textContent = "规格";
    el("viewer-title").textContent = "完整解决方案";
    el("viewer-sub").textContent = "双页预览";
    el("open-pdf").textContent = "打开 PDF";
  }
}

function getLabel(obj, lang) {
  if (!obj) return "";
  if (typeof obj === "string") return obj;
  return safeText(obj[lang] ?? obj.CN ?? obj.EN ?? "");
}

async function loadData() {
  const resp = await fetch("./data/pages_data.json", { cache: "no-store" });
  if (!resp.ok) throw new Error("Failed to load pages_data.json");
  return await resp.json();
}

function showPanel(series) {
  const map = {
    WecLac: "panel-weclac",
    "WecPro® Formula": "panel-formula",
    "WecPro® Solution": "panel-solution",
  };
  Object.values(map).forEach((id) => (el(id).hidden = true));
  el(map[series] || map.WecLac).hidden = false;
}

function openModal(html) {
  const modal = el("modal");
  el("modal-body").innerHTML = html || "";
  modal.hidden = false;
  document.body.style.overflow = "hidden";
}

function closeModal() {
  const modal = el("modal");
  modal.hidden = true;
  el("modal-body").innerHTML = "";
  document.body.style.overflow = "";
}

function renderWecLac(data, lang) {
  const grid = el("weclac-grid");
  const areas = el("weclac-areas");
  const strains = Array.isArray(data?.weclac?.strains) ? data.weclac.strains : [];
  const coreSet = new Set(Array.isArray(data?.weclac?.core_codes) ? data.weclac.core_codes : []);
  const byCode = new Map(strains.map((s) => [safeText(s.code), s]));
  grid.innerHTML = strains
    .map((s) => {
      const code = safeText(s.code);
      const name = getLabel(s.base_name, lang);
      const latinHtml = safeText(s.latin_html || "");
      const icon = safeText(s.icon || "");
      const coreStar = coreSet.has(code) ? " <span title='Core' style='color:rgba(236,72,153,0.9)'>★</span>" : "";
      const detailsLabel = lang === "EN" ? "Details" : "介绍";
      const kFeature = lang === "EN" ? "Highlights" : "特点";
      const kClinical = lang === "EN" ? "Clinical" : "临床";
      const kPatent = lang === "EN" ? "Patents" : "专利";
      const kSpec = lang === "EN" ? "Specification" : "规格";
      const feature = safeText(s.feature?.[lang] ?? s.feature ?? "");
      const clinical = safeText(s.clinical?.[lang] ?? s.clinical ?? "");
      const patent = safeText(s.patent?.[lang] ?? s.patent ?? "");
      const spec = safeText(s.spec?.[lang] ?? s.spec ?? "");

      return `
        <div class="ip-wrap">
          <div class="ip-card">
            <div class="ip-avatar" role="button" tabindex="0" data-code="${encodeURIComponent(code)}">
              ${icon ? `<img src="${icon}" alt="${code}"/>` : `<div class="code-pill">${code}</div>`}
            </div>
            <div class="ip-code">
              ${latinHtml ? `<span class="ip-latin">${latinHtml}</span>` : ""}
              <span class="code-pill">${code}${coreStar}</span>
            </div>
            ${name ? `<div class="ip-name">${escapeHtml(name)}</div>` : ""}
            <details class="ip-details">
              <summary>${detailsLabel}</summary>
              <div class="ip-kv">
                <div class="ip-k">${kFeature}</div><div class="ip-v">${escapeHtml(feature) || "—"}</div>
                <div class="ip-k">${kClinical}</div><div class="ip-v">${escapeHtml(clinical) || "—"}</div>
                <div class="ip-k">${kPatent}</div><div class="ip-v">${escapeHtml(patent) || "—"}</div>
                <div class="ip-k">${kSpec}</div><div class="ip-v">${escapeHtml(spec) || "—"}</div>
              </div>
            </details>
          </div>
        </div>
      `;
    })
    .join("");

  const openByCode = (code) => {
    const s = byCode.get(code);
    if (!s) return;
    const name = getLabel(s.base_name, lang);
    const latinHtml = safeText(s.latin_html || "");
    const coreStar = coreSet.has(code) ? " <span title='Core' style='color:rgba(236,72,153,0.9)'>★</span>" : "";
    const kFeature = lang === "EN" ? "Highlights" : "特点";
    const kClinical = lang === "EN" ? "Clinical" : "临床";
    const kPatent = lang === "EN" ? "Patents" : "专利";
    const kSpec = lang === "EN" ? "Specification" : "规格";
    const feature = safeText(s.feature?.[lang] ?? s.feature ?? "");
    const clinical = safeText(s.clinical?.[lang] ?? s.clinical ?? "");
    const patent = safeText(s.patent?.[lang] ?? s.patent ?? "");
    const spec = safeText(s.spec?.[lang] ?? s.spec ?? "");

    const modalHtml = `
      <div class="section-title">${latinHtml ? latinHtml + " " : ""}<span class="code-pill">${code}</span>${coreStar}</div>
      ${name ? `<div style="color:rgba(15,23,42,0.65);font-weight:800;margin-top:-6px">${escapeHtml(name)}</div>` : ""}
      <div style="margin-top:12px" class="kv-table">
        <div class="kv-grid">
          <div class="kv-k">${kFeature}</div><div class="kv-v">${escapeHtml(feature) || "—"}</div>
          <div class="kv-k">${kClinical}</div><div class="kv-v">${escapeHtml(clinical) || "—"}</div>
          <div class="kv-k">${kPatent}</div><div class="kv-v">${escapeHtml(patent) || "—"}</div>
          <div class="kv-k">${kSpec}</div><div class="kv-v">${escapeHtml(spec) || "—"}</div>
        </div>
      </div>
    `;
    openModal(modalHtml);
  };

  grid.querySelectorAll("[data-code]").forEach((node) => {
    const code = decodeURIComponent(node.getAttribute("data-code") || "");
    node.addEventListener("click", () => openByCode(code));
    node.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openByCode(code);
      }
    });
  });

  const areasList = Array.isArray(data?.weclac?.areas) ? data.weclac.areas : [];
  const areaDetails = Array.isArray(data?.weclac?.area_details) ? data.weclac.area_details : [];
  if (!areasList.length) {
    areas.hidden = true;
    return;
  }
  areas.hidden = false;
  const title = lang === "EN" ? "Supported Application Areas" : "功能方向";
  const chips = areasList.map((a) => `<span class="chip">${escapeHtml(getLabel(a, lang) || a)}</span>`).join("");
  let detailsHtml = "";
  if (areaDetails.length) {
    const summary = lang === "EN" ? "Show area details" : "展开方向详情";
    const rows = areaDetails
      .map((row) => {
        const d = getLabel(row.direction, lang) || safeText(row.direction);
        const desc = safeHtml(row?.desc_html?.[lang] ?? row?.desc_html ?? "");
        if (!d) return "";
        return `<div class="ip-k">${escapeHtml(d)}</div><div class="ip-v">${desc || "—"}</div>`;
      })
      .join("");
    detailsHtml = `
      <details class="ip-details" style="margin-top:12px">
        <summary>${escapeHtml(summary)}</summary>
        <div class="ip-kv" style="grid-template-columns:190px 1fr">${rows}</div>
      </details>
    `;
  }
  areas.innerHTML = `<div class="section-title">${title}</div><div style="display:flex;gap:10px;flex-wrap:wrap">${chips}</div>${detailsHtml}`;
}

function renderFormula(data, lang) {
  const list = el("formula-list");
  const items = Array.isArray(data?.formula?.items) ? data.formula.items : [];
  list.innerHTML = items
    .map((it) => {
      const direction = getLabel(it.direction_label, lang) || safeText(it.direction || "");
      const product = getLabel(it.product, lang) || safeText(it.product || "");
      const benefits = getLabel(it.benefit, lang) || safeText(it.benefit || "");
      const coreFormula = getLabel(it.core_formula, lang) || safeText(it.core_formula || "");
      const detailsLabel = lang === "EN" ? "Details" : "介绍";
      const kBenefit = lang === "EN" ? "Benefits" : "健康功效";
      const kCore = lang === "EN" ? "Core Formula" : "核心配方";
      const badge = product ? `<span class="f-badge">${escapeHtml(product)}</span>` : "";
      return `
        <details class="f-row-wrap">
          <summary class="f-row" style="list-style:none">
            <div class="f-left">
              <div class="f-dot"></div>
              <div style="min-width:0">
                <div class="f-title">${escapeHtml(direction)}</div>
              </div>
            </div>
            <div class="f-actions">
              ${badge}
              <span class="f-cta">${detailsLabel}</span>
            </div>
          </summary>
          <div class="f-expand">
            <div class="kv-table">
              <div class="kv-grid">
                <div class="kv-k">${kBenefit}</div><div class="kv-v">${escapeHtml(benefits) || "—"}</div>
                <div class="kv-k">${kCore}</div><div class="kv-v">${safeHtml(coreFormula) || "—"}</div>
              </div>
            </div>
          </div>
        </details>
      `;
    })
    .join("");
}

function renderSolution(data, lang, view) {
  const categories = Array.isArray(data?.solutions?.categories) ? data.solutions.categories : [];
  const pdfLink = safeText(data?.solutions?.pdf?.[lang] || data?.solutions?.pdf?.CN || "");
  el("open-pdf").href = pdfLink || "#";

  const catSel = el("cat");
  const subSel = el("sub");
  const pill = el("pill");

  const catOptions = categories.map((c) => ({
    key: safeText(c.key),
    label: getLabel(c.label, lang) || safeText(c.key),
    accent1: safeText(c.accent1),
    accent2: safeText(c.accent2),
    scenarios: Array.isArray(c.scenarios) ? c.scenarios : [],
    core: c.core || {},
  }));

  if (!catOptions.length) {
    catSel.innerHTML = "";
    subSel.innerHTML = "";
    pill.textContent = lang === "EN" ? "No data" : "无数据";
    return;
  }

  catSel.innerHTML = catOptions
    .map((c) => `<option value="${encodeURIComponent(c.key)}">${escapeHtml(c.label)}</option>`)
    .join("");

  function setTheme(cat) {
    if (cat?.accent1) document.documentElement.style.setProperty("--accent1", cat.accent1);
    if (cat?.accent2) document.documentElement.style.setProperty("--accent2", cat.accent2);
  }

  function fillSubs(catKey) {
    const cat = catOptions.find((x) => x.key === catKey) || catOptions[0];
    setTheme(cat);
    const scenarios = cat.scenarios || [];
    subSel.innerHTML = scenarios
      .map((s) => {
        const key = safeText(s.key);
        const label = getLabel(s.label, lang) || key;
        return `<option value="${encodeURIComponent(key)}">${escapeHtml(label)}</option>`;
      })
      .join("");
    if (scenarios.length) subSel.value = encodeURIComponent(safeText(scenarios[0].key));
    return cat;
  }

  function render(catKey, scenKey) {
    const cat = catOptions.find((x) => x.key === catKey) || catOptions[0];
    const scenarios = cat.scenarios || [];
    const scen = scenarios.find((s) => safeText(s.key) === scenKey) || scenarios[0];

    const catLabel = getLabel(cat.label, lang) || catKey;
    const scenLabel = getLabel(scen.title, lang) || getLabel(scen.label, lang) || scenKey;
    pill.textContent = `${catLabel} · ${scenLabel}`;

    // Core formula section
    const coreName = getLabel(cat.core?.name, lang);
    const coreFormula = getLabel(cat.core?.formula, lang);
    const sep = lang === "EN" ? ": " : "：";
    const coreLine = coreName && coreFormula ? `<strong>${escapeHtml(coreName)}</strong>${sep}${safeHtml(coreFormula)}` : safeHtml(coreFormula || coreName || "");
    el("solution-core").innerHTML = coreLine || (lang === "EN" ? "—" : "—");

    const highlights = Array.isArray(scen.highlights?.[lang] ?? scen.highlights) ? (scen.highlights?.[lang] ?? scen.highlights) : [];
    el("solution-highlights").innerHTML = highlights.length ? `<ul>${highlights.slice(0, 4).map((x) => `<li>${safeHtml(x)}</li>`).join("")}</ul>` : "";

    // Clinical studies
    const clinical = Array.isArray(scen.clinical) ? scen.clinical : [];
    const clinicalCard = el("solution-clinical-card");
    if (!clinical.length) {
      clinicalCard.hidden = true;
    } else {
      clinicalCard.hidden = false;
      const rows = clinical
        .map((row) => {
          const k = getLabel(row.key, lang) || safeText(row.key);
          const badges = (Array.isArray(row.ids) ? row.ids : []).map((x) => safeHtml(x)).join("");
          return `<div class="kv-k">${escapeHtml(k)}</div><div class="kv-v"><div style="display:flex;gap:8px;flex-wrap:wrap">${badges}</div></div>`;
        })
        .join("");
      el("solution-clinical").innerHTML = `<div class="kv-grid" style="grid-template-columns:130px 1fr">${rows}</div>`;
    }

    // Specs
    const specs = Array.isArray(scen.specs) ? scen.specs : [];
    const specCard = el("solution-spec-card");
    if (!specs.length) {
      specCard.hidden = true;
    } else {
      specCard.hidden = false;
      el("solution-spec").innerHTML = specs
        .map((s) => {
          const title = getLabel(s.title, lang) || "";
          const clinical = getLabel(s.clinical, lang) || "";
          const exc = Array.isArray(s.excipients?.[lang] ?? s.excipients) ? (s.excipients?.[lang] ?? s.excipients) : [];
          const total = getLabel(s.total, lang) || "";
          const lClinical = lang === "EN" ? "Clinical Strain Formula" : "临床菌配方";
          const lExc = lang === "EN" ? "Excipients (mg)" : "辅料配方（mg）";
          const lTotal = lang === "EN" ? "Total Weight" : "总克重";
          const sep2 = lang === "EN" ? ": " : "：";
          const excHtml = exc.map((x) => `<div>• ${escapeHtml(x)}</div>`).join("");
          return `
            <div class="spec-box">
              <div class="spec-title">${escapeHtml(title)}</div>
              ${clinical ? `<div class="spec-meta">${lClinical}</div><div style="font-size:0.92rem;line-height:1.45">${safeHtml(clinical)}</div>` : ""}
              ${excHtml ? `<div class="spec-meta">${lExc}</div><div style="font-size:0.92rem;line-height:1.45">${excHtml}</div>` : ""}
              ${total ? `<div class="spec-meta" style="display:flex;gap:8px;align-items:baseline"><span>${lTotal}${sep2}</span><span style="color:var(--text);font-weight:850;font-size:0.92rem">${escapeHtml(total)}</span></div>` : ""}
            </div>
          `;
        })
        .join("");
    }

    // PDF preview images
    const p1 = Number(scen.page1 || 1);
    const p2 = Number(scen.page2 || p1 + 1);
    const base = `./assets/pages/${lang}`;
    const img1 = el("img1");
    const img2 = el("img2");
    img1.src = `${base}/${String(p1).padStart(3, "0")}.png`;
    img2.src = `${base}/${String(p2).padStart(3, "0")}.png`;

    const img2Wrap = el("img2-wrap");
    if (view === "left") {
      img2Wrap.style.display = "none";
    } else if (view === "right") {
      img1.src = `${base}/${String(p2).padStart(3, "0")}.png`;
      img2Wrap.style.display = "none";
    } else {
      img2Wrap.style.display = "";
    }
  }

  const initialCatKey = catOptions[0].key;
  catSel.value = encodeURIComponent(initialCatKey);
  const cat = fillSubs(initialCatKey);
  const initialScenKey = safeText((cat.scenarios?.[0] || {}).key || "");
  render(initialCatKey, initialScenKey);

  catSel.addEventListener("change", () => {
    const k = decodeURIComponent(catSel.value || "");
    const c = fillSubs(k);
    const first = c.scenarios?.[0] || {};
    render(k, safeText(first.key || ""));
  });
  subSel.addEventListener("change", () => {
    const catKey = decodeURIComponent(catSel.value || "");
    const scenKey = decodeURIComponent(subSel.value || "");
    render(catKey, scenKey);
  });
}

function escapeHtml(text) {
  return safeText(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function safeHtml(html) {
  // Our JSON is built locally from trusted sources (your PPTX/Excel). Keep formatting (italics, links).
  return safeText(html);
}

function init() {
  const lang = (localStorage.getItem(STORAGE_LANG) || DEFAULT_LANG).toUpperCase() === "EN" ? "EN" : "CN";
  const series = localStorage.getItem(STORAGE_SERIES) || DEFAULT_SERIES;
  const view = (localStorage.getItem(STORAGE_VIEW) || DEFAULT_VIEW).toLowerCase();

  const heroLangSeg = document.querySelector(".hero-actions .seg");
  const heroSeriesSeg = document.querySelector(".hero-series .seg");
  const viewSeg = document.querySelector(".viewer-actions .seg");

  setSelected(heroLangSeg, (btn) => btn.dataset.lang === lang);
  setSelected(heroSeriesSeg, (btn) => btn.dataset.series === series);
  setSelected(viewSeg, (btn) => btn.dataset.view === view);
  setHero(lang);
  showPanel(series);

  // Series base theme (Solutions will override per category on selection)
  if (series === "WecLac") {
    document.documentElement.style.setProperty("--accent1", "#7C3AED");
    document.documentElement.style.setProperty("--accent2", "#FF2D55");
  } else if (series === "WecPro® Formula") {
    document.documentElement.style.setProperty("--accent1", "#4F46E5");
    document.documentElement.style.setProperty("--accent2", "#EC4899");
  } else {
    document.documentElement.style.setProperty("--accent1", "#6366F1");
    document.documentElement.style.setProperty("--accent2", "#EC4899");
  }

  el("modal").addEventListener("click", (e) => {
    if (e.target?.dataset?.close === "1") closeModal();
  });
  document.querySelectorAll("[data-close='1']").forEach((btn) => btn.addEventListener("click", closeModal));
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeModal();
  });

  heroLangSeg.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-lang]");
    if (!btn) return;
    const next = btn.dataset.lang === "EN" ? "EN" : "CN";
    localStorage.setItem(STORAGE_LANG, next);
    location.reload();
  });
  heroSeriesSeg.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-series]");
    if (!btn) return;
    localStorage.setItem(STORAGE_SERIES, btn.dataset.series);
    location.reload();
  });
  viewSeg.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-view]");
    if (!btn) return;
    localStorage.setItem(STORAGE_VIEW, btn.dataset.view);
    location.reload();
  });

  Promise.resolve()
    .then(() => loadData())
    .then((data) => {
      renderWecLac(data, lang);
      renderFormula(data, lang);
      renderSolution(data, lang, view);
    })
    .catch((e) => {
      console.error(e);
      el("weclac-grid").innerHTML = `<div class="card">Data load failed</div>`;
    });
}

init();
