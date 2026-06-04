(function () {
  const FILE_API_ORIGIN = (window.localStorage && window.localStorage.getItem("bparty_api_origin")) || "http://127.0.0.1:8012";
  const API_ORIGIN = window.location.protocol === "file:" ? FILE_API_ORIGIN.replace(/\/$/, "") : "";
  const JOB_ENDPOINT = `${API_ORIGIN}/api/jobs`;
  const AUTH_ENDPOINT = `${API_ORIGIN}/api/login`;
  const LOGOUT_ENDPOINT = `${API_ORIGIN}/api/logout`;
  const ME_ENDPOINT = `${API_ORIGIN}/api/me`;
  const ACCOUNT_ENDPOINT = `${API_ORIGIN}/api/accounts`;
  const PERMISSION_ENDPOINT = `${API_ORIGIN}/api/permissions`;
  const AUTH_STORAGE_KEY = "bparty_auth_token";
  const POLL_INTERVAL_MS = 2000;

  const body = document.body;
  const loginOverlay = document.querySelector("#loginOverlay");
  const appShell = document.querySelector("#appShell");
  const loginForm = document.querySelector("#loginForm");
  const loginUsername = document.querySelector("#loginUsername");
  const loginPassword = document.querySelector("#loginPassword");
  const loginButton = document.querySelector("#loginButton");
  const loginMessage = document.querySelector("#loginMessage");
  const form = document.querySelector("#processForm");
  const manifestInput = document.querySelector("#manifestInput");
  const billInput = document.querySelector("#billInput");
  const taskNoInput = document.querySelector("#taskNoInput");
  const targetTaxInput = document.querySelector("#targetTaxInput");
  const targetCountInput = document.querySelector("#targetCountInput");
  const manifestName = document.querySelector("#manifestName");
  const billName = document.querySelector("#billName");
  const submitButton = document.querySelector("#submitButton");
  const resetButton = document.querySelector("#resetButton");
  const refreshButton = document.querySelector("#refreshButton");
  const statusPill = document.querySelector("#statusPill");
  const queueSummary = document.querySelector("#queueSummary");
  const queueList = document.querySelector("#queueList");
  const resultSummary = document.querySelector("#resultSummary");
  const resultList = document.querySelector("#resultList");
  const userBadge = document.querySelector("#userBadge");
  const logoutButton = document.querySelector("#logoutButton");
  const permissionButton = document.querySelector("#permissionButton");
  const permissionModal = document.querySelector("#permissionModal");
  const closePermissionButton = document.querySelector("#closePermissionButton");
  const permissionMessage = document.querySelector("#permissionMessage");
  const permissionList = document.querySelector("#permissionList");

  let pollTimer = null;
  let authToken = window.localStorage ? window.localStorage.getItem(AUTH_STORAGE_KEY) || "" : "";
  let currentUser = null;
  let knownActiveTaskIds = new Set();
  let cancelledTaskIds = new Set();
  let deletedResultTaskIds = new Set();
  let cancelingTaskIds = new Set();
  let deletingResultTaskIds = new Set();
  let previewTaskIds = new Set();
  let previewLoadingTaskIds = new Set();
  let previewErrorByTaskId = new Map();
  let completedJobs = new Map();
  let lastQueueJobs = [];
  let queueMeta = {
    max_concurrency: 5,
    running_count: 0,
    queued_count: 0,
    jobs: [],
  };

  const statusText = {
    queued: "排队中",
    running: "运行中",
    succeeded: "已完成",
    failed: "处理失败",
  };

  function authHeaders() {
    return authToken ? { Authorization: `Bearer ${authToken}` } : {};
  }

  function persistAuthToken(token) {
    authToken = token || "";
    if (!window.localStorage) return;
    if (authToken) {
      window.localStorage.setItem(AUTH_STORAGE_KEY, authToken);
    } else {
      window.localStorage.removeItem(AUTH_STORAGE_KEY);
    }
  }

  function authedUrl(url) {
    if (!authToken) return url;
    try {
      const target = new URL(url, API_ORIGIN || window.location.origin);
      target.searchParams.set("token", authToken);
      return target.href;
    } catch {
      const separator = String(url).includes("?") ? "&" : "?";
      return `${url}${separator}token=${encodeURIComponent(authToken)}`;
    }
  }

  function showLogin(message = "") {
    clearPollTimer();
    currentUser = null;
    body.classList.add("login-required");
    body.classList.remove("auth-loading");
    loginOverlay.classList.remove("hidden");
    appShell.classList.add("hidden");
    loginMessage.textContent = message;
    permissionButton.classList.add("hidden");
    userBadge.textContent = "未登录";
    loginUsername.focus();
  }

  function showApp(user) {
    currentUser = user;
    body.classList.remove("auth-loading", "login-required");
    loginOverlay.classList.add("hidden");
    appShell.classList.remove("hidden");
    userBadge.textContent = `${user.display_name || user.user_id} · ${user.can_view_all_tasks ? "可看全部任务" : "仅看本人任务"}`;
    permissionButton.classList.toggle("hidden", user.role !== "admin");
  }

  function requireLoggedIn() {
    return Boolean(currentUser && authToken);
  }

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

  function scheduleQueuePoll() {
    clearPollTimer();
    if (lastQueueJobs.length > 0) {
      pollTimer = window.setTimeout(() => {
        refreshQueue({ detectCompleted: true, keepPolling: true });
      }, POLL_INTERVAL_MS);
    }
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

  function clampProgress(value) {
    const progress = Number(value);
    if (!Number.isFinite(progress)) return 0;
    return Math.max(0, Math.min(100, progress));
  }

  function formatProgress(value) {
    const progress = clampProgress(value);
    return Number.isInteger(progress) ? String(progress) : progress.toFixed(1);
  }

  function formatRuntime(job) {
    const seconds = Number(job.elapsed_seconds);
    if (!Number.isFinite(seconds)) return "未开始";
    if (seconds < 60) return `${Math.round(seconds)} 秒`;
    return `${(seconds / 60).toFixed(1)} 分钟`;
  }

  function formatDateTime(value) {
    if (!value) return "暂无创建时间";
    const normalized = String(value).includes("T") ? String(value) : String(value).replace(" ", "T");
    const date = new Date(normalized);
    if (Number.isNaN(date.getTime())) return String(value);
    const pad = (part) => String(part).padStart(2, "0");
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }

  function queuePositionText(job) {
    if (job.status === "running") {
      return `运行槽位 ${job.running_position || 1}/${queueMeta.max_concurrency || 5}`;
    }
    if (job.status === "queued") {
      return `队列第 ${job.queue_position || 1} 位`;
    }
    return statusText[job.status] || "未知状态";
  }

  function displayTaskNo(job) {
    return String(job.display_task_no || job.task_no || "历史任务");
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

  function renderCategoryTags(categories) {
    if (!categories.length) {
      return '<span class="muted">暂无提单品类</span>';
    }

    return categories
      .map((item) => {
        const text = typeof item === "object" ? item.category || item.name || JSON.stringify(item) : item;
        return `<span class="tag">${escapeHtml(String(text))}</span>`;
      })
      .join("");
  }

  function pickMetric(source, keys, fallback = "") {
    for (const key of keys) {
      const value = source && source[key];
      if (value !== null && value !== undefined && value !== "") {
        return value;
      }
    }
    return fallback;
  }

  function formatMetric(value, suffix = "") {
    if (value === null || value === undefined || value === "") return "暂无";
    if (typeof value === "number") {
      return `${Number(value.toFixed(4)).toLocaleString()}${suffix}`;
    }
    return `${String(value)}${suffix}`;
  }

  function deliverySummaryHtml(job) {
    const stats = job.stats || {};
    const rows = Array.isArray(job.output_rows) ? job.output_rows : [];
    const outputRows = pickMetric(stats, ["output_rows"], rows.length || "");
    const targetRows = pickMetric(stats, ["target_item_count"], "");
    const estimatedTax = pickMetric(stats, ["estimated_tax_amount"], "");
    const targetTax = pickMetric(stats, ["target_tax_amount"], "");
    const taxGap = pickMetric(stats, ["tax_gap"], "");
    const elapsed = pickMetric(stats, ["task_elapsed_minutes", "elapsed_minutes"], "");
    const status = job.status || "unknown";

    if (status !== "succeeded" && !rows.length && !Object.keys(stats).length) {
      return '<span class="muted">暂无交付摘要</span>';
    }

    const cards = [
      ["输出行数", targetRows ? `${formatMetric(outputRows)} / ${formatMetric(targetRows)}` : formatMetric(outputRows), "生成结果与目标行数"],
      ["预计税金", formatMetric(estimatedTax), targetTax ? `目标 ${formatMetric(targetTax)}` : "按结果估算"],
      ["税金差额", formatMetric(taxGap), "预计税金减目标税金"],
      ["任务耗时", elapsed ? `${formatMetric(elapsed)} 分钟` : formatRuntime(job), "从运行开始到结束"],
    ];

    return cards
      .map(([label, value, hint]) => {
        return `
          <div class="delivery-card">
            <span>${escapeHtml(label)}</span>
            <strong>${escapeHtml(value)}</strong>
            <small>${escapeHtml(hint)}</small>
          </div>
        `;
      })
      .join("");
  }

  function previewCell(value) {
    const text = formatCell(value);
    return text || "—";
  }

  function renderOutputPreview(job) {
    const rows = Array.isArray(job.output_rows) ? job.output_rows : [];
    if (job.task_id && previewLoadingTaskIds.has(job.task_id)) {
      return '<div class="preview-empty">正在加载预览...</div>';
    }
    if (job.task_id && previewErrorByTaskId.has(job.task_id)) {
      return `<div class="preview-empty">${escapeHtml(previewErrorByTaskId.get(job.task_id))}</div>`;
    }
    if (!rows.length) {
      return '<div class="preview-empty">暂无可预览的输出行</div>';
    }
    const visibleRows = rows.slice(0, 30);
    const columns = [
      ["中文品名", "中文品名"],
      ["英文品名", "英文品名"],
      ["商品编码", "HTS"],
      ["材质", "材质"],
      ["数量", "数量"],
      ["单价", "单价"],
      ["总价", "总价"],
      ["毛重", "毛重"],
      ["税率", "税率"],
      ["税金", "税金"],
    ];
    return `
      <div class="preview-scroll" role="region" aria-label="输出 Excel 预览">
        <table class="output-preview-table">
          <thead>
            <tr>
              ${columns.map(([, label]) => `<th>${escapeHtml(label)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${visibleRows.map((row) => `
              <tr>
                ${columns.map(([key]) => `<td>${escapeHtml(previewCell(row[key]))}</td>`).join("")}
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderSourceFileLinks(job) {
    if ((job.status || "") !== "failed" || !job.source_files || typeof job.source_files !== "object") {
      return "";
    }
    const fileOrder = [
      ["manifest", "清单"],
      ["bill", "提单"],
    ];
    return fileOrder
      .map(([kind, defaultLabel]) => {
        const file = job.source_files[kind];
        if (!file || !file.download_url) {
          return "";
        }
        const href = normalizeDownloadUrl(file.download_url);
        if (!href) {
          return "";
        }
        const label = file.label || defaultLabel;
        const filename = file.filename || label;
        return `
          <a
            class="source-file-link"
            href="${escapeHtml(authedUrl(href))}"
            title="${escapeHtml(filename)}"
            download
          >
            下载${escapeHtml(label)}
          </a>
        `;
      })
      .join("");
  }

  function resultCardHtml(job) {
    const status = job.status || "unknown";
    const message = job.error || job.message || "";
    const categories = pickArray(job, ["bill_categories"]);
    const createdAt = formatDateTime(job.created_at || job.updated_at);
    const downloadUrl = normalizeDownloadUrl(job.download_url);
    const taskLabel = displayTaskNo(job);
    const isDeleting = Boolean(job.task_id && deletingResultTaskIds.has(job.task_id));
    const hasPreview = status === "succeeded" && Boolean(job.task_id);
    const previewOpen = Boolean(job.task_id && previewTaskIds.has(job.task_id));
    const previewButton = hasPreview
      ? `
        <button
          type="button"
          class="preview-toggle-button"
          data-action="toggle-preview"
          data-task-id="${escapeHtml(job.task_id)}"
          aria-expanded="${previewOpen ? "true" : "false"}"
          aria-controls="preview-${escapeHtml(job.task_id)}"
        >
          ${previewOpen ? "取消预览" : "预览结果"}
        </button>
      `
      : "";
    const download = downloadUrl && status === "succeeded"
      ? `<a class="download-link" href="${escapeHtml(authedUrl(downloadUrl))}" download>下载结果</a>`
      : "";
    const sourceFileLinks = renderSourceFileLinks(job);
    const deleteButton = job.task_id
      ? `
        <button
          type="button"
          class="delete-result-button"
          data-action="delete-result"
          data-task-id="${escapeHtml(job.task_id)}"
          ${isDeleting ? "disabled" : ""}
          aria-label="删除任务 ${escapeHtml(taskLabel)} 的结果记录"
        >
          ${isDeleting ? "删除中" : "删除记录"}
        </button>
      `
      : "";
    const delivery = status === "failed"
      ? ""
      : `
        <div class="delivery-summary">
          ${deliverySummaryHtml(job)}
        </div>
      `;

    return `
      <article class="result-card ${escapeHtml(status)}">
        <div class="result-card-head">
          <div>
            <div class="result-title-row">
              <h3>任务 ${escapeHtml(taskLabel)}</h3>
              <span class="queue-status ${escapeHtml(status)}">${escapeHtml(statusText[status] || status)}</span>
            </div>
            <div class="result-meta-row">创建时间 ${escapeHtml(createdAt)}</div>
            <p class="result-message">${escapeHtml(message || "暂无状态说明")}</p>
          </div>
          <div class="result-actions">
            ${previewButton}
            ${download}
            ${sourceFileLinks}
            ${deleteButton}
          </div>
        </div>
        ${delivery}
        <div class="result-subgrid">
          <section>
            <h4>提单品类</h4>
            <div class="tag-list">${renderCategoryTags(categories)}</div>
            <div
              id="preview-${escapeHtml(job.task_id || taskLabel)}"
              class="output-preview-collapse ${previewOpen ? "open" : ""}"
            >
              <div class="output-preview-panel">
                <div class="output-preview-title">输出 Excel 预览</div>
                ${renderOutputPreview(job)}
              </div>
            </div>
          </section>
        </div>
      </article>
    `;
  }

  function renderResults() {
    const jobs = Array.from(completedJobs.values()).sort((left, right) => {
      return String(right.completed_at || right.updated_at || "").localeCompare(String(left.completed_at || left.updated_at || ""));
    });
    resultSummary.textContent = jobs.length ? `已保留 ${jobs.length} 个完成或失败任务` : "暂无完成或失败任务";
    if (!jobs.length) {
      resultList.className = "result-list empty";
      resultList.textContent = "暂无完成或失败任务";
      return;
    }
    resultList.className = "result-list";
    resultList.innerHTML = jobs.map(resultCardHtml).join("");
  }

  function normalizeQueueData(data) {
    if (Array.isArray(data)) {
      return {
        max_concurrency: 5,
        running_count: data.filter((job) => job.status === "running").length,
        queued_count: data.filter((job) => job.status === "queued").length,
        jobs: data.filter((job) => job.status === "queued" || job.status === "running"),
        finished_jobs: data.filter((job) => job.status === "succeeded" || job.status === "failed"),
      };
    }
    const jobs = Array.isArray(data.jobs) ? data.jobs : [];
    const unfinishedJobs = jobs.filter((job) => job.status === "queued" || job.status === "running");
    const finishedJobs = jobs.filter((job) => job.status === "succeeded" || job.status === "failed");
    return {
      max_concurrency: Number(data.max_concurrency) || 5,
      running_count: Number(data.running_count) || unfinishedJobs.filter((job) => job.status === "running").length,
      queued_count: Number(data.queued_count) || unfinishedJobs.filter((job) => job.status === "queued").length,
      jobs: unfinishedJobs,
      finished_jobs: finishedJobs,
    };
  }

  function mergeFinishedJobs(jobs) {
    jobs.forEach((job) => {
      if (job.task_id && (job.status === "succeeded" || job.status === "failed")) {
        completedJobs.set(job.task_id, job);
      }
    });
    renderResults();
  }

  function renderQueue(meta) {
    queueMeta = normalizeQueueData(meta);
    const jobs = queueMeta.jobs.filter((job) => !cancelledTaskIds.has(job.task_id));
    queueMeta.jobs = jobs;
    queueMeta.running_count = jobs.filter((job) => job.status === "running").length;
    queueMeta.queued_count = jobs.filter((job) => job.status === "queued").length;
    lastQueueJobs = jobs;
    queueSummary.textContent = `运行中 ${queueMeta.running_count}/${queueMeta.max_concurrency} · 排队 ${queueMeta.queued_count}`;

    if (jobs.length === 0) {
      queueList.className = "queue-list empty";
      queueList.textContent = "暂无排队或运行任务";
      setStatus("队列空闲", "done");
      return;
    }

    queueList.className = "queue-list";
    queueList.innerHTML = jobs.map(renderQueueCard).join("");
    setStatus(`运行 ${queueMeta.running_count}/${queueMeta.max_concurrency} · 排队 ${queueMeta.queued_count}`, "running");
  }

  function renderQueueCard(job) {
    const status = job.status || "unknown";
    const progress = formatProgress(job.progress);
    const phase = job.phase || "queued";
    const message = job.message || (status === "queued" ? "等待可用运行槽位" : "");
    const taskLabel = displayTaskNo(job);
    const isCanceling = Boolean(job.task_id && cancelingTaskIds.has(job.task_id));
    const cancelButton = job.task_id
      ? `
        <button
          type="button"
          class="cancel-job-button"
          data-action="cancel-job"
          data-task-id="${escapeHtml(job.task_id)}"
          ${isCanceling ? "disabled" : ""}
          aria-label="取消任务 ${escapeHtml(taskLabel)}"
        >
          ${isCanceling ? "取消中" : "取消任务"}
        </button>
      `
      : "";
    return `
      <article class="queue-card ${escapeHtml(status)}">
        <div class="queue-card-main">
          <div class="queue-title-row">
            <h3>任务 ${escapeHtml(taskLabel)}</h3>
            <span class="queue-status ${escapeHtml(status)}">${escapeHtml(statusText[status] || status)}</span>
          </div>
          <div class="queue-meta-row">
            <span>${escapeHtml(queuePositionText(job))}</span>
            <span>阶段 ${escapeHtml(phase)}</span>
            <span>耗时 ${escapeHtml(formatRuntime(job))}</span>
          </div>
          <p class="queue-message">${escapeHtml(message || "暂无状态说明")}</p>
        </div>
        <div class="progress-block" aria-label="任务 ${escapeHtml(taskLabel)} 进度 ${escapeHtml(progress)}%">
          <div class="progress-label">
            <span>真实进度</span>
            <strong>${escapeHtml(progress)}%</strong>
          </div>
          <div class="progress-track" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${escapeHtml(progress)}">
            <span style="width: ${escapeHtml(progress)}%"></span>
          </div>
          <div class="queue-actions-row">
            ${cancelButton}
          </div>
        </div>
      </article>
    `;
  }

  async function fetchJson(url, options) {
    const requestOptions = options ? { ...options } : {};
    requestOptions.headers = {
      ...authHeaders(),
      ...(requestOptions.headers || {}),
    };
    const response = await fetch(url, requestOptions);
    const contentType = response.headers.get("content-type") || "";
    const payload = contentType.includes("application/json") ? await response.json() : { message: await response.text() };
    if (!response.ok || (payload.code && payload.code !== 200)) {
      const error = new Error(payload.detail || payload.message || `请求失败：${response.status}`);
      error.status = response.status;
      if (response.status === 401) {
        persistAuthToken("");
        showLogin(error.message || "请先登录");
      }
      throw error;
    }
    return unwrapResponse(payload);
  }

  async function fetchJob(taskID) {
    return fetchJson(`${JOB_ENDPOINT}/${encodeURIComponent(taskID)}`);
  }

  async function login(username, password) {
    loginButton.disabled = true;
    loginMessage.textContent = "正在登录...";
    try {
      const data = await fetchJson(AUTH_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });
      persistAuthToken(data.token || "");
      showApp(data.user);
      loginForm.reset();
      setStatus("待上传", "");
      clearLocalQueueState();
      await refreshQueueWithFinished();
    } catch (error) {
      persistAuthToken("");
      loginMessage.textContent = networkErrorMessage(error, "登录失败");
    } finally {
      loginButton.disabled = false;
    }
  }

  async function restoreSession() {
    if (!authToken) {
      showLogin();
      return;
    }
    try {
      const user = await fetchJson(ME_ENDPOINT);
      showApp(user);
      clearLocalQueueState();
      await refreshQueueWithFinished();
    } catch (error) {
      persistAuthToken("");
      showLogin(networkErrorMessage(error, "登录已失效，请重新登录"));
    }
  }

  async function logout() {
    try {
      await fetchJson(LOGOUT_ENDPOINT, { method: "POST" });
    } catch {
      // Local cleanup is enough if the server token already expired.
    }
    persistAuthToken("");
    clearLocalQueueState();
    showLogin("已退出登录");
  }

  function accountPasswordInput(user) {
    if (user.role === "admin") {
      return `
        <label class="account-password-field">
          <span>当前 admin 密码</span>
          <input type="password" data-admin-current-password="${escapeHtml(user.user_id)}" autocomplete="current-password" />
        </label>
        <label class="account-password-field">
          <span>新密码</span>
          <input type="password" data-new-password="${escapeHtml(user.user_id)}" autocomplete="new-password" />
        </label>
      `;
    }
    return `
      <label class="account-password-field">
        <span>新密码</span>
        <input type="password" data-new-password="${escapeHtml(user.user_id)}" autocomplete="new-password" />
      </label>
    `;
  }

  function renderAccounts(users) {
    if (!users.length) {
      permissionList.className = "permission-list empty";
      permissionList.textContent = "暂无账号";
      return;
    }
    permissionList.className = "permission-list";
    permissionList.innerHTML = users.map((user) => {
      const checked = user.can_view_all_tasks ? "checked" : "";
      const permissionControl = user.role === "admin"
        ? '<span class="account-role-badge">管理员</span>'
        : `
          <label class="account-permission-toggle">
            <input
              type="checkbox"
              data-action="toggle-permission"
              data-user-id="${escapeHtml(user.user_id)}"
              ${checked}
            />
            <span>${user.can_view_all_tasks ? "可查看所有任务" : "只能查看自己的任务"}</span>
          </label>
        `;
      return `
        <article class="permission-row">
          <div class="account-main">
            <strong>${escapeHtml(user.display_name || user.user_id)}</strong>
            <small>${escapeHtml(user.user_id)} · ${escapeHtml(user.role === "admin" ? "admin" : "user")}</small>
          </div>
          <div class="account-controls">
            ${permissionControl}
            <div class="account-password-controls ${user.role === "admin" ? "admin-password-controls" : ""}">
              ${accountPasswordInput(user)}
              <button
                type="button"
                class="secondary compact"
                data-action="change-password"
                data-user-id="${escapeHtml(user.user_id)}"
              >
                修改密码
              </button>
            </div>
          </div>
        </article>
      `;
    }).join("");
  }

  async function openPermissionModal() {
    permissionModal.classList.remove("hidden");
    body.classList.add("modal-open");
    permissionMessage.textContent = "正在读取账号...";
    permissionList.innerHTML = "";
    try {
      const data = await fetchJson(ACCOUNT_ENDPOINT);
      renderAccounts(Array.isArray(data.users) ? data.users : []);
      permissionMessage.textContent = "新增账号名会自动递增。修改密码后，该账号已登录会话会立即失效。";
    } catch (error) {
      permissionMessage.textContent = networkErrorMessage(error, "账号读取失败");
    }
  }

  function closePermissionModal() {
    permissionModal.classList.add("hidden");
    body.classList.remove("modal-open");
  }

  async function refreshAccounts(message = "") {
    const data = await fetchJson(ACCOUNT_ENDPOINT);
    renderAccounts(Array.isArray(data.users) ? data.users : []);
    if (message) {
      permissionMessage.textContent = message;
    }
  }

  async function createAccount() {
    const passwordInput = document.querySelector("#newAccountPassword");
    const password = passwordInput ? String(passwordInput.value || "") : "";
    permissionMessage.textContent = "正在新增账号...";
    try {
      const user = await fetchJson(ACCOUNT_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password }),
      });
      if (passwordInput) passwordInput.value = "";
      await refreshAccounts(`已新增 ${user.user_id}。`);
    } catch (error) {
      permissionMessage.textContent = networkErrorMessage(error, "新增账号失败");
    }
  }

  async function updatePermission(userID, canViewAllTasks) {
    permissionMessage.textContent = "正在更新权限...";
    try {
      await fetchJson(`${PERMISSION_ENDPOINT}/${encodeURIComponent(userID)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ can_view_all_tasks: canViewAllTasks }),
      });
      await refreshAccounts("权限已更新。");
    } catch (error) {
      permissionMessage.textContent = networkErrorMessage(error, "权限更新失败");
      refreshAccounts().catch(() => {});
    }
  }

  async function changeAccountPassword(userID, button) {
    const newPasswordInput = permissionModal.querySelector(`[data-new-password="${CSS.escape(userID)}"]`);
    const currentPasswordInput = permissionModal.querySelector(`[data-admin-current-password="${CSS.escape(userID)}"]`);
    const newPassword = newPasswordInput ? String(newPasswordInput.value || "") : "";
    const currentAdminPassword = currentPasswordInput ? String(currentPasswordInput.value || "") : "";
    if (!newPassword) {
      permissionMessage.textContent = "新密码不能为空。";
      return;
    }
    button.disabled = true;
    permissionMessage.textContent = "正在修改密码...";
    try {
      const data = await fetchJson(`${ACCOUNT_ENDPOINT}/${encodeURIComponent(userID)}/password`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          new_password: newPassword,
          current_admin_password: currentAdminPassword,
        }),
      });
      if (data.current_user_logged_out) {
        persistAuthToken("");
        clearLocalQueueState();
        closePermissionModal();
        showLogin("admin 密码已更新，请用新密码重新登录");
        return;
      }
      await refreshAccounts(`${userID} 密码已更新。`);
    } catch (error) {
      permissionMessage.textContent = networkErrorMessage(error, "密码修改失败");
    } finally {
      button.disabled = false;
    }
  }

  async function ensurePreviewRows(taskID) {
    if (!taskID || previewLoadingTaskIds.has(taskID)) return;
    const current = completedJobs.get(taskID);
    if (current && Array.isArray(current.output_rows) && current.output_rows.length > 0) return;
    previewLoadingTaskIds.add(taskID);
    previewErrorByTaskId.delete(taskID);
    renderResults();
    try {
      const job = await fetchJob(taskID);
      if (job && job.task_id) {
        completedJobs.set(job.task_id, job);
      }
      if (!Array.isArray(job.output_rows) || job.output_rows.length === 0) {
        previewErrorByTaskId.set(taskID, "该任务记录没有可预览的输出行，请直接下载 Excel 查看。");
      }
    } catch (error) {
      previewErrorByTaskId.set(taskID, networkErrorMessage(error, "预览加载失败"));
    } finally {
      previewLoadingTaskIds.delete(taskID);
      renderResults();
    }
  }

  async function loadCompletedDetails(previousIds, nextJobs) {
    const nextIds = new Set(nextJobs.map((job) => job.task_id).filter(Boolean));
    const completedIds = Array.from(previousIds).filter((id) => {
      return !nextIds.has(id) && !cancelledTaskIds.has(id) && !deletedResultTaskIds.has(id);
    });
    for (const id of completedIds) {
      try {
        const job = await fetchJob(id);
        if (job.status === "succeeded" || job.status === "failed") {
          completedJobs.set(job.task_id, job);
        }
      } catch (error) {
        if (error && error.status === 404) {
          cancelledTaskIds.add(id);
          completedJobs.delete(id);
          continue;
        }
        completedJobs.set(id, {
          task_id: id,
          status: "failed",
          message: networkErrorMessage(error, `任务 ${id} 状态获取失败`),
          error: networkErrorMessage(error, `任务 ${id} 状态获取失败`),
          updated_at: new Date().toISOString(),
        });
      }
    }
    renderResults();
  }

  async function cancelJob(taskID) {
    if (!taskID || cancelingTaskIds.has(taskID)) return;
    cancelingTaskIds.add(taskID);
    renderQueue(queueMeta);
    let refreshed = false;
    try {
      await fetchJson(`${JOB_ENDPOINT}/${encodeURIComponent(taskID)}`, {
        method: "DELETE",
      });
      cancelledTaskIds.add(taskID);
      knownActiveTaskIds.delete(taskID);
      completedJobs.delete(taskID);
      lastQueueJobs = lastQueueJobs.filter((job) => job.task_id !== taskID);
      setStatus("任务已取消", "done");
      await refreshQueue({ detectCompleted: false, keepPolling: true });
      refreshed = true;
    } catch (error) {
      setStatus("取消失败", "error");
      queueSummary.textContent = networkErrorMessage(error, `任务 ${taskID} 取消失败`);
    } finally {
      cancelingTaskIds.delete(taskID);
      if (!refreshed) {
        renderQueue(queueMeta);
        scheduleQueuePoll();
      }
    }
  }

  async function deleteResultRecord(taskID) {
    if (!taskID || deletingResultTaskIds.has(taskID)) return;
    deletingResultTaskIds.add(taskID);
    renderResults();
    try {
      await fetchJson(`${JOB_ENDPOINT}/${encodeURIComponent(taskID)}`, {
        method: "DELETE",
      });
      completedJobs.delete(taskID);
      deletedResultTaskIds.add(taskID);
      setStatus("记录已删除", "done");
      renderResults();
    } catch (error) {
      setStatus("删除失败", "error");
      resultSummary.textContent = networkErrorMessage(error, `任务 ${taskID} 记录删除失败`);
      renderResults();
    } finally {
      deletingResultTaskIds.delete(taskID);
      renderResults();
    }
  }

  async function refreshQueue(options = {}) {
    const { detectCompleted = true, keepPolling = true } = options;
    if (!requireLoggedIn()) return;
    const previousIds = new Set(knownActiveTaskIds);
    try {
      const data = await fetchJson(JOB_ENDPOINT);
      const meta = normalizeQueueData(data);
      renderQueue(meta);
      mergeFinishedJobs(meta.finished_jobs);
      if (detectCompleted) {
        await loadCompletedDetails(previousIds, meta.jobs);
      }
      knownActiveTaskIds = new Set(meta.jobs.map((job) => job.task_id).filter(Boolean));
      if (keepPolling) {
        scheduleQueuePoll();
      }
    } catch (error) {
      clearPollTimer();
      setStatus("队列获取失败", "error");
      queueSummary.textContent = networkErrorMessage(error, "任务队列获取失败");
      submitButton.disabled = false;
    }
  }

  async function refreshQueueWithFinished() {
    if (!requireLoggedIn()) return;
    const data = await fetchJson(`${JOB_ENDPOINT}?include_finished=true&limit=20`);
    const meta = normalizeQueueData(data);
    renderQueue(meta);
    mergeFinishedJobs(meta.finished_jobs);
    knownActiveTaskIds = new Set(meta.jobs.map((job) => job.task_id).filter(Boolean));
    scheduleQueuePoll();
  }

  function buildPayload() {
    const manifest = manifestInput.files[0];
    const bill = billInput.files[0];
    const displayTaskNo = String(taskNoInput.value || "").trim();
    const targetTaxAmount = Number(targetTaxInput.value);
    const targetItemCount = Number(targetCountInput.value);
    const payload = new FormData();

    payload.append("manifest", manifest);
    payload.append("manifest_file", manifest);
    payload.append("file", manifest);
    payload.append("bill", bill);
    payload.append("bill_file", bill);
    payload.append("bl_file", bill);
    payload.append("display_task_no", displayTaskNo);
    payload.append("task_no", displayTaskNo);
    payload.append("target_tax_amount", String(targetTaxAmount));
    payload.append("target_item_count", String(targetItemCount));
    payload.append("profile", "auto");
    payload.append("mode", "auto");

    return payload;
  }

  async function submitForm(event) {
    event.preventDefault();

    if (!requireLoggedIn()) {
      showLogin("请先登录");
      return;
    }
    if (!manifestInput.files[0] || !billInput.files[0]) {
      setStatus("缺少文件", "error");
      return;
    }
    const displayTaskNo = String(taskNoInput.value || "").trim();
    if (!displayTaskNo) {
      setStatus("任务编号必填", "error");
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
    setStatus("任务启动中", "running");

    try {
      const data = await fetchJson(JOB_ENDPOINT, {
        method: "POST",
        body: buildPayload(),
      });
      if (data.task_id) {
        knownActiveTaskIds.add(data.task_id);
      }
      await refreshQueue({ detectCompleted: true, keepPolling: true });
    } catch (error) {
      setStatus("启动失败", "error");
      completedJobs.set(`submit_error_${Date.now()}`, {
        task_id: "提交失败",
        display_task_no: "提交失败",
        status: "failed",
        message: networkErrorMessage(error, "任务启动失败"),
        error: networkErrorMessage(error, "任务启动失败"),
        updated_at: new Date().toISOString(),
      });
      renderResults();
    } finally {
      submitButton.disabled = false;
    }
  }

  function clearLocalQueueState() {
    clearPollTimer();
    knownActiveTaskIds = new Set();
    cancelledTaskIds = new Set();
    deletedResultTaskIds = new Set();
    cancelingTaskIds = new Set();
    deletingResultTaskIds = new Set();
    previewTaskIds = new Set();
    previewLoadingTaskIds = new Set();
    previewErrorByTaskId = new Map();
    completedJobs = new Map();
    lastQueueJobs = [];
    renderQueue({
      max_concurrency: queueMeta.max_concurrency || 5,
      running_count: 0,
      queued_count: 0,
      jobs: [],
    });
  }

  loginForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const username = String(loginUsername.value || "").trim();
    const password = String(loginPassword.value || "");
    if (!username || !password) {
      loginMessage.textContent = "请输入账号和密码";
      return;
    }
    login(username, password);
  });
  logoutButton.addEventListener("click", logout);
  permissionButton.addEventListener("click", openPermissionModal);
  closePermissionButton.addEventListener("click", () => {
    closePermissionModal();
  });
  permissionModal.addEventListener("click", (event) => {
    if (event.target === permissionModal) {
      closePermissionModal();
      return;
    }
    const target = event.target instanceof Element ? event.target : event.target.parentElement;
    const createButton = target ? target.closest("[data-action='create-account']") : null;
    if (createButton) {
      createButton.disabled = true;
      createAccount().finally(() => {
        createButton.disabled = false;
      });
      return;
    }
    const passwordButton = target ? target.closest("[data-action='change-password']") : null;
    if (passwordButton) {
      changeAccountPassword(passwordButton.dataset.userId, passwordButton);
      return;
    }
    const checkbox = target ? target.closest("[data-action='toggle-permission']") : null;
    if (!checkbox) return;
    checkbox.disabled = true;
    updatePermission(checkbox.dataset.userId, checkbox.checked).finally(() => {
      checkbox.disabled = false;
    });
  });
  manifestInput.addEventListener("change", updateFileNames);
  billInput.addEventListener("change", updateFileNames);
  form.addEventListener("submit", submitForm);
  queueList.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target : event.target.parentElement;
    const button = target ? target.closest("[data-action='cancel-job']") : null;
    if (!button) return;
    cancelJob(button.dataset.taskId);
  });
  resultList.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target : event.target.parentElement;
    const previewButton = target ? target.closest("[data-action='toggle-preview']") : null;
    if (previewButton) {
      const taskId = previewButton.dataset.taskId;
      if (previewTaskIds.has(taskId)) {
        previewTaskIds.delete(taskId);
      } else {
        previewTaskIds.add(taskId);
        ensurePreviewRows(taskId);
      }
      renderResults();
      return;
    }
    const button = target ? target.closest("[data-action='delete-result']") : null;
    if (!button) return;
    deleteResultRecord(button.dataset.taskId);
  });
  refreshButton.addEventListener("click", () => {
    refreshQueue({ detectCompleted: true, keepPolling: true });
  });
  resetButton.addEventListener("click", () => {
    form.reset();
    updateFileNames();
    taskNoInput.value = "";
    renderQueue(queueMeta);
  });

  updateFileNames();
  renderResults();
  clearLocalQueueState();
  restoreSession().catch((error) => {
    setStatus("队列获取失败", "error");
    queueSummary.textContent = networkErrorMessage(error, "登录状态恢复失败");
    showLogin(networkErrorMessage(error, "登录状态恢复失败"));
  });
})();
