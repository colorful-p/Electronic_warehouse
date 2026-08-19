/* ============================================================
   app.js - 物料库网站核心逻辑（全动态版）
   ★ 完全跟随 data.js 的列自动渲染 ★
   Excel 有多少列 → data.js 就有多少字段 → 页面显示多少内容
   特殊列智能识别：编号(主键) / 型号 / 厂家 / 数量 / 位置 / 备注 / Datasheet
   ============================================================ */

(function () {
  "use strict";

  // ---- 状态管理 ----
  let filteredList = [];
  let activeFilter = "all";
  let scannerStream = null;

  // ---- DOM 元素 ----
  const $ = (sel) => document.querySelector(sel);
  const el = {
    cardList: $("#cardList"),
    searchInput: $("#searchInput"),
    statsBar: $("#statsBar"),
    totalCount: $("#totalCount"),
    emptyState: $("#emptyState"),
    modalOverlay: $("#modalOverlay"),
    modalIdLabel: $("#modalIdLabel"),
    modalTitle: $("#modalTitle"),
    modalSubtitle: $("#modalSubtitle"),
    detailGrid: $("#detailGrid"),
    actionButtons: $("#actionButtons"),
    modalClose: $("#modalClose"),
    scannerOverlay: $("#scannerOverlay"),
    scannerVideo: $("#scannerVideo"),
    scannerClose: $("#scannerClose"),
    scannerUnsupported: $("#scannerUnsupported"),
    themeToggle: $("#themeToggle"),
    fab: $("#fab"),
    toast: $("#toast"),
  };

  // ============================================================
  //  列信息 —— 完全从数据自动识别，不写死任何字段
  // ============================================================
  const columns = Object.keys(materials[0] || {});

  // 智能识别特殊列（按表头文字匹配，找不到就返回 null，不强求）
  const colId = findCol(/编号|料号|id$/i);
  const colModel = findCol(/型号|产品|名称|model/i);
  const colMaker = findCol(/厂家|品牌|厂商|maker|supplier|供应商/i);
  const colQty = findCol(/数量|库存|qty/i);
  const colPlace = findCol(/位置|仓位|位号|place/i);
  const colRemark = findCol(/备注|说明|remark/i);
  const colDatasheet = findCol(/datasheet|数据手册|手册|链接|url/i);

  // 主键：优先"编号"列，没有就用第一列
  const primaryKey = colId || columns[0] || "id";
  // 卡片副标题列：优先"型号"类列，且不能和主键重复
  const subtitleKey =
    colModel && colModel !== primaryKey
      ? colModel
      : columns.find(function (c) {
          return c !== primaryKey && c !== colQty && c !== colRemark;
        }) || null;

  function findCol(re) {
    for (let i = 0; i < columns.length; i++) {
      if (re.test(columns[i])) return columns[i];
    }
    return null;
  }
  function getId(m) {
    return m[primaryKey] != null ? String(m[primaryKey]) : "";
  }
  function getVal(m, col) {
    return col && m[col] != null ? String(m[col]) : "";
  }
  // 该列是否全部是数字（决定数量预警是否启用）
  function isNumericCol(col) {
    if (!col) return false;
    return materials.every(function (m) {
      const v = getVal(m, col);
      if (v === "") return true;
      const num = parseFloat(v);
      return !isNaN(num) && String(num) === v.trim();
    });
  }
  const qtyNumeric = isNumericCol(colQty);

  // ============================================================
  //  初始化
  // ============================================================
  function init() {
    loadTheme();
    renderStats();
    renderCards(materials);
    bindEvents();
    handleURLParam();
  }

  // ============================================================
  //  主题 (暗黑模式)
  // ============================================================
  function loadTheme() {
    const saved = localStorage.getItem("material-theme") || "light";
    document.documentElement.setAttribute("data-theme", saved);
    updateThemeIcon(saved);
  }

  function toggleTheme() {
    const current = document.documentElement.getAttribute("data-theme");
    const next = current === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("material-theme", next);
    updateThemeIcon(next);
  }

  function updateThemeIcon(theme) {
    if (el.themeToggle) {
      el.themeToggle.textContent = theme === "dark" ? "\u2600\uFE0F" : "\u{1F319}";
    }
  }

  // ============================================================
  //  统计栏（有"厂家"类列才分组，否则只显示总数）
  // ============================================================
  function renderStats() {
    const all = materials.length;

    let html = `
      <div class="stat-chip active" data-filter="all">
        <span>全部</span>
        <span class="stat-num">${all}</span>
      </div>
    `;

    if (colMaker) {
      const makers = {};
      materials.forEach(function (m) {
        const v = getVal(m, colMaker);
        if (v) makers[v] = (makers[v] || 0) + 1;
      });
      Object.keys(makers)
        .sort()
        .forEach(function (maker) {
          html += `
            <div class="stat-chip" data-filter="${escapeAttr(maker)}">
              <span>${escapeHtml(maker)}</span>
              <span class="stat-num">${makers[maker]}</span>
            </div>
          `;
        });
    }

    el.statsBar.innerHTML = html;

    el.statsBar.querySelectorAll(".stat-chip").forEach(function (chip) {
      chip.addEventListener("click", function () {
        el.statsBar.querySelectorAll(".stat-chip").forEach(function (c) {
          c.classList.remove("active");
        });
        this.classList.add("active");
        activeFilter = this.getAttribute("data-filter");
        applyFilters();
      });
    });
  }

  // ============================================================
  //  渲染卡片列表（动态：所有非空字段都展示）
  // ============================================================
  function renderCards(list) {
    if (list.length === 0) {
      el.cardList.innerHTML = "";
      el.emptyState.style.display = "block";
      return;
    }

    el.emptyState.style.display = "none";

    el.cardList.innerHTML = list
      .map(function (m) {
        const id = getId(m);

        // 数量徽章（仅当数量列是纯数字时启用预警）
        let qtyBadge = "";
        if (colQty && qtyNumeric && getVal(m, colQty) !== "") {
          const qtyNum = parseInt(getVal(m, colQty), 10) || 0;
          let qtyClass = "low";
          let qtyIcon = "\u26A0\uFE0F";
          if (qtyNum >= 100) {
            qtyClass = "high";
            qtyIcon = "\u2705";
          } else if (qtyNum >= 20) {
            qtyClass = "mid";
            qtyIcon = "\u26A0\uFE0F";
          }
          qtyBadge = `
            <div class="card-qty-badge ${qtyClass}">
              ${qtyIcon} ${qtyNum}
            </div>
          `;
        }

        // 副标题
        const sub = subtitleKey ? getVal(m, subtitleKey) : "";

        // 卡片底部：除 主键/副标题/数量 外的所有非空字段（完全跟 Excel 走）
        const tags = columns
          .filter(function (c) {
            return c !== primaryKey && c !== subtitleKey && c !== colQty;
          })
          .map(function (c) {
            const v = getVal(m, c);
            if (v === "") return "";
            return `
              <span class="card-tag">
                <span class="card-tag-key">${escapeHtml(c)}</span>
                ${escapeHtml(v)}
              </span>
            `;
          })
          .join("");

        return `
          <div class="card" data-id="${escapeAttr(id)}">
            <div class="card-header">
              <div class="card-title-area">
                <div class="card-id">${escapeHtml(id)}</div>
                ${sub ? `<div class="card-model">${escapeHtml(sub)}</div>` : ""}
              </div>
              ${qtyBadge}
            </div>
            ${tags ? `<div class="card-footer">${tags}</div>` : ""}
          </div>
        `;
      })
      .join("");

    el.cardList.querySelectorAll(".card").forEach(function (card) {
      card.addEventListener("click", function () {
        const id = this.getAttribute("data-id");
        const item = materials.find(function (m) {
          return getId(m) === id;
        });
        if (item) openDetail(item);
      });
    });
  }

  // ============================================================
  //  搜索 + 筛选（搜索范围 = 所有列）
  // ============================================================
  function applyFilters() {
    const kw = el.searchInput.value.trim().toLowerCase();

    filteredList = materials.filter(function (m) {
      // 厂家筛选
      if (activeFilter !== "all" && getVal(m, colMaker) !== activeFilter) {
        return false;
      }
      // 关键词搜索：遍历所有字段
      if (kw === "") return true;

      return columns.some(function (c) {
        const v = getVal(m, c);
        return v && v.toLowerCase().indexOf(kw) > -1;
      });
    });

    renderCards(filteredList);
    updateTotalCount();
  }

  function updateTotalCount() {
    el.totalCount.textContent = filteredList.length || materials.length;
  }

  // ============================================================
  //  详情弹窗（动态渲染所有列）
  // ============================================================
  function openDetail(m) {
    const id = getId(m);

    // 更新 URL (方便分享/扫码)
    const newUrl = window.location.pathname + "?id=" + encodeURIComponent(id);
    history.replaceState(null, "", newUrl);

    // 填充头部
    el.modalIdLabel.textContent = id;
    el.modalTitle.textContent = subtitleKey ? getVal(m, subtitleKey) || "-" : "-";
    el.modalSubtitle.textContent = colMaker ? getVal(m, colMaker) : "";

    // 详情网格：按 Excel 列顺序逐列渲染
    const detailRows = columns
      .filter(function (c) {
        return c !== primaryKey;
      })
      .map(function (c) {
        const v = getVal(m, c);
        if (v === "") return "";

        // Datasheet 列：值是网址 → 渲染成可点击链接
        if (c === colDatasheet && /^https?:\/\//i.test(v)) {
          return `
            <div class="detail-item">
              <div class="detail-label">${escapeHtml(c)}</div>
              <div class="detail-value">
                <a href="${escapeAttr(v)}" target="_blank" rel="noopener"
                   style="color: var(--primary); word-break: break-all;">
                  ${escapeHtml(v)}
                </a>
              </div>
            </div>
          `;
        }

        const isRemark = c === colRemark;
        return `
          <div class="detail-item ${isRemark ? "remark" : ""}">
            <div class="detail-label">${escapeHtml(c)}</div>
            <div class="detail-value ${isRemark ? "remark" : ""}">${escapeHtml(v)}</div>
          </div>
        `;
      })
      .join("");

    // 主键也要显示（除非它就是标题）
    const idRow = `
      <div class="detail-item">
        <div class="detail-label">${escapeHtml(primaryKey)}</div>
        <div class="detail-value">${escapeHtml(id)}</div>
      </div>
    `;

    el.detailGrid.innerHTML = idRow + detailRows;

    // 操作按钮：Datasheet（可选）
    const dsUrl = colDatasheet ? getVal(m, colDatasheet) : "";
    const hasDatasheet = /^https?:\/\//i.test(dsUrl);

    el.actionButtons.innerHTML = `
      ${hasDatasheet
        ? `<a class="action-btn action-btn-primary" href="${escapeAttr(dsUrl)}" target="_blank" rel="noopener">
             \u{1F4D6} 查看 Datasheet
           </a>`
        : ""
      }
    `;

    el.modalOverlay.classList.add("active");
    document.body.style.overflow = "hidden";
  }

  function closeDetail() {
    el.modalOverlay.classList.remove("active");
    document.body.style.overflow = "";
    history.replaceState(null, "", window.location.pathname);
  }

  // ============================================================
  //  URL 参数处理 (扫码后自动打开)
  // ============================================================
  function handleURLParam() {
    const params = new URLSearchParams(window.location.search);
    const id = params.get("id");
    if (id) {
      const item = materials.find(function (m) {
        return getId(m) === id;
      });
      if (item) {
        setTimeout(function () {
          openDetail(item);
        }, 300);
      } else {
        showToast("\u672A\u627E\u5230\u7F16\u53F7: " + id);
      }
    }
  }

  // ============================================================
  //  扫码功能
  // ============================================================
  async function startScanner() {
    if (!("BarcodeDetector" in window)) {
      el.scannerUnsupported.style.display = "flex";
      el.scannerOverlay.classList.add("active");
      return;
    }

    try {
      scannerStream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: "environment" },
      });
      el.scannerVideo.srcObject = scannerStream;
      el.scannerVideo.setAttribute("playsinline", "true");
      el.scannerVideo.play();
      el.scannerOverlay.classList.add("active");

      const detector = new BarcodeDetector({
        formats: ["qr_code", "code_128", "code_39", "ean_13"],
      });

      let scanning = true;

      async function detectLoop() {
        if (!scanning) return;
        try {
          const barcodes = await detector.detect(el.scannerVideo);
          if (barcodes.length > 0) {
            const rawValue = barcodes[0].rawValue;
            scanning = false;
            stopScanner();
            handleScanResult(rawValue);
            return;
          }
        } catch (e) {
          // 忽略检测错误
        }
        requestAnimationFrame(detectLoop);
      }
      detectLoop();
    } catch (err) {
      el.scannerUnsupported.style.display = "flex";
      el.scannerOverlay.classList.add("active");
    }
  }

  function stopScanner() {
    if (scannerStream) {
      scannerStream.getTracks().forEach(function (t) {
        t.stop();
      });
      scannerStream = null;
    }
    el.scannerOverlay.classList.remove("active");
  }

  function handleScanResult(value) {
    try {
      if (value.indexOf("?id=") > -1) {
        const url = new URL(value);
        const id = url.searchParams.get("id");
        if (id) {
          const item = materials.find(function (m) {
            return getId(m) === id;
          });
          if (item) {
            openDetail(item);
            showToast("\u626B\u7801\u6210\u529F: " + id);
            return;
          }
        }
      }
      const item = materials.find(function (m) {
        return getId(m) === value.trim();
      });
      if (item) {
        openDetail(item);
        showToast("\u626B\u7801\u6210\u529F: " + getId(item));
        return;
      }
      showToast("\u672A\u627E\u5230\u5339\u914D\u7684\u5143\u4EF6: " + value);
    } catch (e) {
      showToast("\u626B\u7801\u7ED3\u679C: " + value);
    }
  }

  // ============================================================
  //  Toast 提示
  // ============================================================
  let toastTimer = null;
  function showToast(msg) {
    el.toast.textContent = msg;
    el.toast.classList.add("show");
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      el.toast.classList.remove("show");
    }, 2500);
  }

  // ============================================================
  //  HTML 转义
  // ============================================================
  function escapeHtml(str) {
    if (str == null) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
  function escapeAttr(str) {
    return escapeHtml(str);
  }

  // ============================================================
  //  事件绑定
  // ============================================================
  function bindEvents() {
    let searchTimer = null;
    el.searchInput.addEventListener("input", function () {
      if (searchTimer) clearTimeout(searchTimer);
      searchTimer = setTimeout(applyFilters, 200);
    });

    el.themeToggle.addEventListener("click", toggleTheme);
    el.modalClose.addEventListener("click", closeDetail);
    el.modalOverlay.addEventListener("click", function (e) {
      if (e.target === el.modalOverlay) closeDetail();
    });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") {
        if (el.modalOverlay.classList.contains("active")) closeDetail();
        if (el.scannerOverlay.classList.contains("active")) stopScanner();
      }
    });

    el.fab.addEventListener("click", startScanner);
    el.scannerClose.addEventListener("click", stopScanner);

    el.totalCount.textContent = materials.length;
  }

  // ============================================================
  //  启动
  // ============================================================
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
