const uploadForm = document.getElementById("uploadForm");
const loadingScreen = document.getElementById("loadingScreen");
const loadingProgressContainer = document.getElementById("loadingProgressContainer");
const loadingProgressBar = document.getElementById("loadingProgressBar");
const loadingProgressText = document.getElementById("loadingProgressText");
const loadingStatusText = document.getElementById("loadingStatusText");
const submitButton = uploadForm.querySelector('button[type="submit"]');

let progressEventSource = null;

function setLoadingProgress(value) {
  const normalizedValue = Math.max(0, Math.min(100, Number(value) || 0));
  const percentageText = `${normalizedValue}%`;

  loadingProgressBar.style.width = percentageText;
  loadingProgressBar.textContent = "";
  loadingProgressText.textContent = percentageText;
  loadingProgressContainer.setAttribute("aria-valuenow", normalizedValue);
}

function showLoading(statusMessage) {
  loadingStatusText.textContent = statusMessage;
  setLoadingProgress(0);
  loadingScreen.style.display = "flex";
}

function closeProgressStream() {
  if (progressEventSource) {
    progressEventSource.close();
    progressEventSource = null;
  }
}

function setSubmittingState(isSubmitting) {
  submitButton.disabled = isSubmitting;
  submitButton.textContent = isSubmitting ? "Memproses..." : "Process & Recapitulation";
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
    loadingStatusText.textContent = "Recap selesai. Mengalihkan...";
    setLoadingProgress(100);
    closeProgressStream();

    if (payload.redirect_url) {
      setTimeout(() => {
        window.location.href = payload.redirect_url;
      }, 250);
    }
  } else if (payload.status === "error") {
    closeProgressStream();
    setSubmittingState(false);
    loadingStatusText.textContent = payload.error || "Proses recap gagal. Silakan coba lagi.";
  }
}

async function startRecapWithProgress(formData) {
  const response = await fetch("/upload_recap/start", {
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
    throw new Error(payload.error || "Gagal memulai proses recap.");
  }

  const progressUrl = `/upload_recap/progress/${encodeURIComponent(payload.job_id)}`;
  progressEventSource = new EventSource(progressUrl);

  progressEventSource.onmessage = function (event) {
    try {
      const progressPayload = JSON.parse(event.data);
      handleProgressMessage(progressPayload);
    } catch (error) {
      loadingStatusText.textContent = "Tidak dapat membaca status progres.";
    }
  };

  progressEventSource.onerror = function () {
    closeProgressStream();
    setSubmittingState(false);
    loadingStatusText.textContent = "Koneksi terputus saat memantau progres.";
  };
}

uploadForm.addEventListener("submit", async function (event) {
  event.preventDefault();

  setSubmittingState(true);
  showLoading("Mengunggah file...");
  loadingStatusText.textContent = "Mengunggah file dan memulai proses recap...";

  const formData = new FormData(uploadForm);

  try {
    await startRecapWithProgress(formData);
  } catch (error) {
    setSubmittingState(false);
    loadingStatusText.textContent = error.message || "Gagal mengunggah file.";
  }
});

window.addEventListener("beforeunload", function () {
  closeProgressStream();
});
