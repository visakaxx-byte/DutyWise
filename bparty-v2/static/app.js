(function () {
  const FILE_API_ORIGIN = (window.localStorage && window.localStorage.getItem("bparty_api_origin")) || "http://127.0.0.1:8012";
  const API_ORIGIN = window.location.protocol === "file:" ? FILE_API_ORIGIN.replace(/\/$/, "") : "";
  const PRECHECK_ENDPOINT = `${API_ORIGIN}/api/precheck`;
  const JOB_ENDPOINT = `${API_ORIGIN}/api/jobs`;
  const form = document.querySelector("#processForm");
  const manifestInput = document.querySelector("#manifestInput");
  const billInput = document.querySelector("#billInput");
  const targetTaxInput = document.querySelector("#targetTaxInput");
  const targetCountInput = document.querySelector("#targetCountInput");
  const manifestName = document.querySelector("#manifestName");
  const billName = document.querySelector("#billName");
  const submitButton = document.querySelector("#submitButton");
  const confirmButton = document.querySelector("#confirmButton");
  const resetButton = document.querySelector("#resetButton");
  const statusPill = document.querySelector("#statusPill");
  const precheckPanel = document.querySelector("#precheckPanel");
  const precheckBadge = document.querySelector("#precheckBadge");
  const precheckGrid = document.querySelector("#precheckGrid");
  const precheckReasons = document.querySelector("#precheckReasons");
  const statsGrid = document.querySelector("#statsGrid");
  const categoryList = document.querySelector("#categoryList");
  const flowList = document.querySelector("#flowList");
  const outputTable = document.querySelector("#outputTable");
  const downloadLink = document.querySelector("#downloadLink");
  const taskId = document.querySelector("#taskId");
  let currentPrecheckId = "";
  let currentTaskId = "";
  let pollTimer = null;

  const statLabels = {
    profile_hint: "模式",
    constraint_status: "约束状态",
    weight_source: "重量锚点",
    weight_evidence: "重量证据",
    weight_confidence: "重量置信度",
    manifest_rows: "清单行数",
    input_categories: "输入品类",
    input_ctns: "输入箱数",
    output_ctns: "输出箱数",
    input_real_weight: "清单实重",
    output_gross_weight: "输出毛重",
    input_declared_value: "原申报金额",
    total_value_usd: "申报总价",
    target_tax_amount: "期望税金",
    estimated_tax_amount: "预计税金",
    tax_gap: "税金差额",
    target_item_count: "目标输出行",
    output_rows: "输出行",
    plausibility_warnings: "约束提示",
    realism_status: "合理性校验",
    realism_warnings: "合理性提示",
    bill_products: "提单品类",
    qualified_manifest_candidates: "客户可用品名",
    replacement_candidates_used: "替换品名",
    filtered_candidates: "排除候选",
    bill_gross_weight_ignored: "提单毛重(忽略)",
    llm_used: "LLM参与",
    llm_parse_used: "LLM解析",
    llm_parse_cache_reused: "复用解析缓存",
    llm_plausibility_used: "LLM合理性估算",
    llm_generation_used: "LLM生成",
    llm_draft_attempts: "LLM草案轮次",
    crawler_used: "爬虫",
    status: "状态",
    phase: "阶段",
    progress: "进度",
    message: "说明",
  };

  const precheckLabels = {
    success_probability_percent: "成功概率",
    estimated_minutes: "预计分钟",
    estimated_seconds: "预计秒数",
    estimated_usable_items: "预计可用条目",
    estimated_positive_tax_items: "预计正税率条目",
    suggested_item_count: "建议条目数",
    target_item_count: "目标输出行",
    target_tax_amount: "期望税金",
    manifest_rows: "清单行数",
    manifest_candidate_count: "客户候选",
    replacement_candidate_count: "替换候选",
    bill_product_count: "提单品类",
    qualified_bill_products: "提单合格品类",
    filtered_bill_products: "提单排除品类",
    sampled_qualified_manifest: "抽样客户合格",
    sampled_qualified_replacement: "抽样替换合格",
    estimated_query_items: "预计查询候选",
    manifest_total_weight: "Excel总重量",
    manifest_weight_source: "重量来源",
    manifest_weight_evidence: "重量证据",
  };

  function setStatus(text, state) {
    statusPill.textContent = text;
    statusPill.className = `status-pill ${state || ""}`.trim();
  }

  function formatBytes(file) {
    if (!file) return "未选择文件";
    if (file.size < 1024) return `${file.name} (${file.size} B)`;
    if (file.size < 1024 * 1024) {
      return `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    }
    return `${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
  }

  function updateFileNames() {
    manifestName.textContent = formatBytes(manifestInput.files[0]);
    billName.textContent = formatBytes(billInput.files[0]);
  }

  function clearPollTimer() {
    if (pollTimer) {
      window.clearTimeout(pollTimer);
      pollTimer = null;
    }
  }

  function resetPrecheck() {
    currentPrecheckId = "";
    precheckPanel.classList.add("hidden");
    precheckBadge.className = "precheck-badge";
    precheckBadge.textContent = "未评估";
    precheckGrid.className = "stats-grid empty";
    precheckGrid.textContent = "暂无评估";
    precheckReasons.className = "reason-list empty";
    precheckReasons.textContent = "暂无风险提示";
    confirmButton.classList.add("hidden");
    confirmButton.disabled = false;
  }

  function resetResults() {
    clearPollTimer();
    currentTaskId = "";
    taskId.textContent = "";
    statsGrid.className = "stats-grid empty";
    statsGrid.textContent = "暂无结果";
    categoryList.className = "tag-list empty";
    categoryList.textContent = "暂无结果";
    flowList.className = "flow-list empty";
    flowList.textContent = "暂无结果";
    downloadLink.classList.add("hidden");
    downloadLink.removeAttribute("href");
    renderRows([]);
  }

  function unwrapResponse(payload) {
    if (payload && typeof payload === "object" && payload.data) {
      return payload.data;
    }
    return payload || {};
  }

  function pickArray(source, keys) {
    for (const key of keys) {
      if (Array.isArray(source[key])) return source[key];
    }
    return [];
  }

  function normalizeDownloadUrl(url) {
    if (!url || typeof url !== "string") return "";
    try {
      return new URL(url, API_ORIGIN || window.location.origin).href;
    } catch {
      return url;
    }
  }

  function networkErrorMessage(error, fallback) {
    const message = error && error.message ? error.message : fallback;
    if (window.location.protocol === "file:" && /fetch|network|load failed|请求失败/i.test(message)) {
      return `当前页面是本地 file:// 打开，已尝试连接 ${FILE_API_ORIGIN}。请确认本地服务正在运行，或直接打开 ${FILE_API_ORIGIN}/ 使用。`;
    }
    return message || fallback;
  }

  function renderGrid(container, stats, labels) {
    const entries = Object.entries(stats || {}).filter(([, value]) => {
      return value === null || ["string", "number", "boolean"].includes(typeof value);
    });

    if (entries.length === 0) {
      container.className = "stats-grid empty";
      container.textContent = "暂无统计";
      return;
    }

    container.className = "stats-grid";
    container.innerHTML = entries
      .map(([key, value]) => {
        const label = labels[key] || key;
        const display = typeof value === "number" ? Number(value.toFixed(4)).toLocaleString() : String(value);
        return `<div class="stat"><div class="stat-label">${escapeHtml(label)}</div><div class="stat-value">${escapeHtml(display)}</div></div>`;
      })
      .join("");
  }

  function renderStats(stats) {
    renderGrid(statsGrid, stats, statLabels);
  }

  function renderCategories(categories) {
    if (!categories.length) {
      categoryList.className = "tag-list empty";
      categoryList.textContent = "暂无提单品类";
      return;
    }

    categoryList.className = "tag-list";
    categoryList.innerHTML = categories
      .map((item) => {
        const text = typeof item === "object" ? item.category || item.name || JSON.stringify(item) : item;
        return `<span class="tag">${escapeHtml(String(text))}</span>`;
      })
      .join("");
  }

  function renderFlow(flow) {
    if (!Array.isArray(flow) || flow.length === 0) {
      flowList.className = "flow-list empty";
      flowList.textContent = "暂无流程";
      return;
    }

    flowList.className = "flow-list";
    flowList.innerHTML = flow
      .map((item) => {
        const stage = item.stage || "stage";
        const status = item.status || "unknown";
        const details = Object.entries(item)
          .filter(([key]) => key !== "stage" && key !== "status")
          .map(([key, value]) => `${key}: ${formatCell(value)}`)
          .join(" · ");
        return `<div class="flow-step"><span>${escapeHtml(stage)}</span><strong>${escapeHtml(status)}</strong><small>${escapeHtml(details)}</small></div>`;
      })
      .join("");
  }

  function renderRows(rows) {
    const thead = outputTable.querySelector("thead");
    const tbody = outputTable.querySelector("tbody");

    if (!rows.length) {
      thead.innerHTML = "";
      tbody.innerHTML = '<tr><td class="empty-cell">暂无输出行</td></tr>';
      return;
    }

    const columns = Array.from(
      rows.reduce((set, row) => {
        Object.keys(row || {}).forEach((key) => set.add(key));
        return set;
      }, new Set())
    );

    thead.innerHTML = `<tr>${columns.map((key) => `<th>${escapeHtml(key)}</th>`).join("")}</tr>`;
    tbody.innerHTML = rows
      .map((row) => {
        return `<tr>${columns.map((key) => `<td>${escapeHtml(formatCell(row[key]))}</td>`).join("")}</tr>`;
      })
      .join("");
  }

  function renderPrecheck(precheck) {
    const status = precheck.status || "unknown";
    const statusText = {
      green: "可执行",
      yellow: "有风险",
      red: "不可执行",
    }[status] || "未知";

    precheckPanel.classList.remove("hidden");
    precheckBadge.className = `precheck-badge ${status}`;
    precheckBadge.textContent = statusText;
    renderGrid(
      precheckGrid,
      {
        success_probability_percent: precheck.success_probability_percent,
        estimated_minutes: precheck.estimated_minutes,
        estimated_usable_items: precheck.estimated_usable_items,
        estimated_positive_tax_items: precheck.estimated_positive_tax_items,
        suggested_item_count: precheck.suggested_item_count,
        target_item_count: precheck.target_item_count,
        bill_product_count: precheck.bill_product_count,
        qualified_bill_products: precheck.qualified_bill_products,
        filtered_bill_products: precheck.filtered_bill_products,
        manifest_candidate_count: precheck.manifest_candidate_count,
        replacement_candidate_count: precheck.replacement_candidate_count,
        estimated_query_items: precheck.estimated_query_items,
      },
      precheckLabels,
    );

    const reasons = Array.isArray(precheck.risk_reasons) ? precheck.risk_reasons : [];
    if (!reasons.length) {
      precheckReasons.className = "reason-list empty";
      precheckReasons.textContent = "未发现明显风险";
      return;
    }
    precheckReasons.className = "reason-list";
    precheckReasons.innerHTML = reasons
      .map((reason) => `<div class="reason-item ${status === "red" ? "error" : ""}">${escapeHtml(reason)}</div>`)
      .join("");
  }

  function renderJob(job) {
    taskId.textContent = job.task_id ? `任务 ${job.task_id}` : "";
    const status = job.status || "unknown";
    if (job.stats) {
      renderStats(job.stats);
    } else {
      renderStats({
        status,
        phase: job.phase || "",
        progress: job.progress || 0,
        message: job.message || "",
      });
    }

    if (Array.isArray(job.flow)) {
      renderFlow(job.flow);
    } else if (job.progress_detail) {
      renderFlow([job.progress_detail]);
    }

    if (Array.isArray(job.output_rows)) {
      renderRows(job.output_rows);
    }

    const downloadUrl = normalizeDownloadUrl(job.download_url);
    if (downloadUrl) {
      downloadLink.href = downloadUrl;
      downloadLink.classList.remove("hidden");
    }
  }

  function formatCell(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "object") return JSON.stringify(value);
    return String(value);
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function buildPayload() {
    const manifest = manifestInput.files[0];
    const bill = billInput.files[0];
    const targetTaxAmount = Number(targetTaxInput.value);
    const targetItemCount = Number(targetCountInput.value);
    const payload = new FormData();

    payload.append("manifest", manifest);
    payload.append("manifest_file", manifest);
    payload.append("file", manifest);
    payload.append("bill", bill);
    payload.append("bill_file", bill);
    payload.append("bl_file", bill);
    payload.append("target_tax_amount", String(targetTaxAmount));
    payload.append("target_item_count", String(targetItemCount));
    payload.append("profile", "auto");
    payload.append("mode", "auto");

    return payload;
  }

  async function submitForm(event) {
    event.preventDefault();
    clearPollTimer();

    if (!manifestInput.files[0] || !billInput.files[0]) {
      setStatus("缺少文件", "error");
      return;
    }
    const targetTaxAmount = Number(targetTaxInput.value);
    const targetItemCount = Number(targetCountInput.value);
    if (!Number.isFinite(targetTaxAmount) || targetTaxAmount <= 0) {
      setStatus("税金无效", "error");
      return;
    }
    if (!Number.isInteger(targetItemCount) || targetItemCount < 1 || targetItemCount > 30) {
      setStatus("条目无效", "error");
      return;
    }

    submitButton.disabled = true;
    confirmButton.classList.add("hidden");
    setStatus("预处理中", "running");
    resetPrecheck();
    resetResults();

    try {
      const response = await fetch(PRECHECK_ENDPOINT, {
        method: "POST",
        body: buildPayload(),
      });
      const contentType = response.headers.get("content-type") || "";
      const payload = contentType.includes("application/json") ? await response.json() : { message: await response.text() };

      if (!response.ok || (payload.code && payload.code !== 200)) {
        throw new Error(payload.detail || payload.message || `请求失败：${response.status}`);
      }

      const data = unwrapResponse(payload);
      currentPrecheckId = data.precheck_id || "";
      const precheck = data.precheck || {};
      const categories = pickArray(data.bill || {}, ["products"]);
      taskId.textContent = currentPrecheckId ? `预处理 ${currentPrecheckId}` : "";
      renderPrecheck(precheck);
      renderCategories(categories);

      if (precheck.can_start && currentPrecheckId) {
        confirmButton.classList.remove("hidden");
        confirmButton.disabled = false;
        setStatus(precheck.status === "yellow" ? "预检有风险" : "预检可执行", precheck.status === "yellow" ? "running" : "done");
      } else {
        confirmButton.classList.add("hidden");
        setStatus("预检未通过", "error");
      }
    } catch (error) {
      resetResults();
      resetPrecheck();
      setStatus("处理失败", "error");
      statsGrid.className = "stats-grid empty";
      statsGrid.textContent = networkErrorMessage(error, "请求失败");
    } finally {
      submitButton.disabled = false;
    }
  }

  async function startJob() {
    if (!currentPrecheckId) {
      setStatus("缺少预处理", "error");
      return;
    }

    confirmButton.disabled = true;
    submitButton.disabled = true;
    setStatus("任务启动中", "running");

    try {
      const payload = new FormData();
      payload.append("precheck_id", currentPrecheckId);
      const response = await fetch(JOB_ENDPOINT, {
        method: "POST",
        body: payload,
      });
      const contentType = response.headers.get("content-type") || "";
      const body = contentType.includes("application/json") ? await response.json() : { message: await response.text() };
      if (!response.ok || (body.code && body.code !== 200)) {
        throw new Error(body.detail || body.message || `请求失败：${response.status}`);
      }

      const data = unwrapResponse(body);
      currentTaskId = data.task_id;
      renderJob(data);
      setStatus("后台处理中", "running");
      pollJob();
    } catch (error) {
      confirmButton.disabled = false;
      submitButton.disabled = false;
      setStatus("启动失败", "error");
      statsGrid.className = "stats-grid empty";
      statsGrid.textContent = networkErrorMessage(error, "任务启动失败");
    }
  }

  async function pollJob() {
    if (!currentTaskId) return;
    try {
      const response = await fetch(`${JOB_ENDPOINT}/${encodeURIComponent(currentTaskId)}`);
      const payload = await response.json();
      if (!response.ok || (payload.code && payload.code !== 200)) {
        throw new Error(payload.detail || payload.message || `请求失败：${response.status}`);
      }

      const data = unwrapResponse(payload);
      renderJob(data);
      if (data.status === "succeeded") {
        clearPollTimer();
        setStatus("已完成", "done");
        submitButton.disabled = false;
        confirmButton.disabled = true;
        return;
      }
      if (data.status === "failed") {
        clearPollTimer();
        setStatus("处理失败", "error");
        statsGrid.className = "stats-grid empty";
        statsGrid.textContent = data.error || data.message || "任务失败";
        submitButton.disabled = false;
        confirmButton.disabled = false;
        return;
      }
      pollTimer = window.setTimeout(pollJob, 2000);
    } catch (error) {
      clearPollTimer();
      setStatus("轮询失败", "error");
      statsGrid.className = "stats-grid empty";
      statsGrid.textContent = networkErrorMessage(error, "任务状态获取失败");
      submitButton.disabled = false;
      confirmButton.disabled = false;
    }
  }

  function invalidatePrecheck() {
    clearPollTimer();
    resetPrecheck();
    resetResults();
    setStatus("待上传");
  }

  manifestInput.addEventListener("change", () => {
    updateFileNames();
    invalidatePrecheck();
  });
  billInput.addEventListener("change", () => {
    updateFileNames();
    invalidatePrecheck();
  });
  targetTaxInput.addEventListener("input", invalidatePrecheck);
  targetCountInput.addEventListener("input", invalidatePrecheck);
  confirmButton.addEventListener("click", startJob);
  form.addEventListener("submit", submitForm);
  resetButton.addEventListener("click", () => {
    clearPollTimer();
    form.reset();
    updateFileNames();
    resetPrecheck();
    resetResults();
    setStatus("待上传");
  });

  updateFileNames();
  resetResults();
})();
