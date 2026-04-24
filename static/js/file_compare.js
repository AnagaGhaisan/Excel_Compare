const uploadForm = document.getElementById("uploadForm");
const loadingScreen = document.getElementById("loadingScreen");
const loadingProgressContainer = document.getElementById("loadingProgressContainer");
const loadingProgressBar = document.getElementById("loadingProgressBar");
const loadingProgressText = document.getElementById("loadingProgressText");
const loadingStatusText = document.getElementById("loadingStatusText");
const submitButton = uploadForm.querySelector('button[type="submit"]');
const accountFormulaModalElement = document.getElementById("accountFormulaModal");
const accountFormulaModal = new bootstrap.Modal(accountFormulaModalElement);
const accountFormulaTableBody = document.getElementById("accountFormulaTableBody");
const confirmAccountFormulaBtn = document.getElementById("confirmAccountFormulaBtn");
const accountFormulaDescriptionText = document.getElementById("accountFormulaDescriptionText");

let progressEventSource = null;
let pendingUploadFormData = null;
let pendingAccounts = [];
let pendingGlUploadToken = null;
let currentDisplayedProgress = 0;
let compareFlowProgress = 0;

const formulaOptions = [
  { value: "credit_minus_debit", label: "-debit + credit" },
  { value: "debit_minus_credit", label: "debit - credit" },
  { value: "negative_debit_minus_credit", label: "-(debit - credit)" },
];

function setLoadingProgress(value, options = {}) {
  const { allowDecrease = false } = options;
  const normalizedValue = Math.max(0, Math.min(100, Number(value) || 0));
  const nextValue = allowDecrease
    ? normalizedValue
    : Math.max(currentDisplayedProgress, normalizedValue);
  currentDisplayedProgress = nextValue;
  compareFlowProgress = nextValue;
  const percentageText = `${nextValue}%`;

  loadingProgressBar.style.width = percentageText;
  loadingProgressBar.textContent = "";
  loadingProgressText.textContent = percentageText;
  loadingProgressContainer.setAttribute("aria-valuenow", nextValue);
}

function showLoading(statusMessage, initialProgress = 0) {
  loadingStatusText.textContent = statusMessage;
  currentDisplayedProgress = 0;
  setLoadingProgress(initialProgress, { allowDecrease: true });
  loadingScreen.style.display = "flex";
}

function resumeLoading(statusMessage) {
  loadingStatusText.textContent = statusMessage;
  currentDisplayedProgress = compareFlowProgress;
  setLoadingProgress(compareFlowProgress, { allowDecrease: true });
  loadingScreen.style.display = "flex";
}

function hideLoading() {
  loadingScreen.style.display = "none";
}

function closeProgressStream() {
  if (progressEventSource) {
    progressEventSource.close();
    progressEventSource = null;
  }
}

function resetCompareProgressState() {
  closeProgressStream();
  pendingGlUploadToken = null;
  compareFlowProgress = 0;
  currentDisplayedProgress = 0;
  setLoadingProgress(0, { allowDecrease: true });
}

function setSubmittingState(isSubmitting) {
  submitButton.disabled = isSubmitting;
  submitButton.textContent = isSubmitting ? "Processing..." : "Upload & Compare";
  confirmAccountFormulaBtn.disabled = isSubmitting;
}

function updateFormulaModalDescription(accounts) {
  if (!accountFormulaDescriptionText) {
    return;
  }

  if (!accounts.length) {
    accountFormulaDescriptionText.textContent =
      "Tidak ada akun unik yang terbaca dari file GL. Silakan cek kembali file yang diunggah.";
    return;
  }

  const accountLabel = accounts.length === 1 ? "akun" : "akun";
  accountFormulaDescriptionText.textContent =
    `Sistem sudah membaca ${accounts.length} ${accountLabel} unik dari file GL. ` +
    "Silakan review rumus tiap akun sebelum proses compare dijalankan.";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function renderAccountFormulaRows(accounts) {
  accountFormulaTableBody.innerHTML = "";

  if (!accounts.length) {
    accountFormulaTableBody.innerHTML = `
      <tr>
        <td colspan="3" class="text-center text-muted py-4">Tidak ada akun yang ditemukan di file GL.</td>
      </tr>
    `;
    return;
  }

  const rowsHtml = accounts
    .map((account, index) => {
      const optionsHtml = formulaOptions
        .map(
          (option) => `
            <option value="${option.value}" ${option.value === account.default_formula ? "selected" : ""}>
              ${option.label}
            </option>
          `,
        )
        .join("");

      return `
        <tr>
          <td>${escapeHtml(account.account_name)}</td>
          <td>${escapeHtml(account.direction || "-")}</td>
          <td>
            <select class="form-select form-select-sm account-formula-select" data-account-index="${index}">
              ${optionsHtml}
            </select>
          </td>
        </tr>
      `;
    })
    .join("");

  accountFormulaTableBody.innerHTML = rowsHtml;
}

async function fetchAccountOptions() {
  const accountFormData = new FormData();
  const glFile = document.getElementById("k3_file").files[0];

  if (!glFile) {
    throw new Error("GL file wajib dipilih.");
  }

  showLoading("Langkah 1 dari 3: membaca akun dari file GL...", 0);
  setLoadingProgress(4);
  accountFormData.append("k3_file", glFile);

  const response = await fetch("/upload/accounts", {
    method: "POST",
    body: accountFormData,
  });

  setLoadingProgress(8);

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || "Gagal membaca akun dari file GL.");
  }

  loadingStatusText.textContent = "Menyiapkan pilihan rumus per akun...";
  setLoadingProgress(15);
  pendingGlUploadToken = payload.gl_upload_token || null;
  return Array.isArray(payload.accounts) ? payload.accounts : [];
}

async function submitFinalUpload() {
  if (!pendingUploadFormData) {
    throw new Error("Upload form belum siap diproses.");
  }

  const formulaMap = {};
  const formulaSelects = accountFormulaTableBody.querySelectorAll(".account-formula-select");
  formulaSelects.forEach((selectElement) => {
    const accountIndex = Number(selectElement.dataset.accountIndex);
    const account = pendingAccounts[accountIndex];
    if (account && account.account_name) {
      formulaMap[account.account_name] = selectElement.value;
    }
  });

  pendingUploadFormData.set("account_formulas", JSON.stringify(formulaMap));
  if (pendingGlUploadToken) {
    pendingUploadFormData.delete("k3_file");
    pendingUploadFormData.set("gl_upload_token", pendingGlUploadToken);
  }

  resumeLoading("Langkah 3 dari 3: mengirim file compare dan menyiapkan proses...");
  setLoadingProgress(20);
  await startUploadWithProgress(pendingUploadFormData);
}

function handleProgressMessage(payload) {
  if (!payload) return;

  if (typeof payload.progress !== "undefined") {
    setLoadingProgress(payload.progress);
  }

  if (payload.message) {
    loadingStatusText.textContent = payload.message;
  }

  if (payload.status === "done") {
    loadingStatusText.textContent = "Comparison completed. Redirecting...";
    setLoadingProgress(100);
    closeProgressStream();

    if (payload.redirect_url) {
      setTimeout(() => {
        window.location.href = payload.redirect_url;
      }, 250);
    }
  } else if (payload.status === "error") {
    closeProgressStream();
    hideLoading();
    setSubmittingState(false);
    loadingStatusText.textContent = payload.error || "Comparison failed. Please try again.";
  }
}

async function startUploadWithProgress(formData) {
  setLoadingProgress(25);
  const response = await fetch("/upload/start", {
    method: "POST",
    body: formData,
  });

  let payload = {};
  try {
    payload = await response.json();
  } catch (error) {
    payload = {};
  }

  if (!response.ok || !payload.job_id) {
    throw new Error(payload.error || "Failed to start upload process.");
  }

  loadingStatusText.textContent = "Job compare berhasil dibuat. Menghubungkan progress monitor...";
  setLoadingProgress(30);

  const progressUrl = `/upload/progress/${encodeURIComponent(payload.job_id)}`;
  progressEventSource = new EventSource(progressUrl);

  progressEventSource.onmessage = function (event) {
    try {
      const progressPayload = JSON.parse(event.data);
      handleProgressMessage(progressPayload);
    } catch (error) {
      loadingStatusText.textContent = "Unable to read progress update.";
    }
  };

  progressEventSource.onerror = function () {
    closeProgressStream();
    hideLoading();
    setSubmittingState(false);
    loadingStatusText.textContent = "Connection lost while tracking progress.";
  };
}

uploadForm.addEventListener("submit", async function (event) {
  event.preventDefault();
  if (!uploadForm.reportValidity()) {
    return;
  }

  setSubmittingState(true);

  try {
    resetCompareProgressState();
    pendingUploadFormData = new FormData(uploadForm);
    pendingAccounts = await fetchAccountOptions();
    renderAccountFormulaRows(pendingAccounts);
    updateFormulaModalDescription(pendingAccounts);
    hideLoading();
    setSubmittingState(false);
    accountFormulaModal.show();
  } catch (error) {
    hideLoading();
    setSubmittingState(false);
    loadingStatusText.textContent = error.message || "Failed to upload files.";
  }
});

confirmAccountFormulaBtn.addEventListener("click", async function () {
  setSubmittingState(true);
  accountFormulaModal.hide();

  try {
    await submitFinalUpload();
  } catch (error) {
    closeProgressStream();
    hideLoading();
    setSubmittingState(false);
    loadingStatusText.textContent = error.message || "Failed to upload files.";
  }
});

accountFormulaModalElement.addEventListener("hidden.bs.modal", function () {
  if (loadingScreen.style.display !== "flex") {
    setSubmittingState(false);
  }
});

window.addEventListener("beforeunload", function () {
  closeProgressStream();
});
