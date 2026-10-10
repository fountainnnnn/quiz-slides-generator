// Backend origin
const BACKEND_BASE_URL =
  new URLSearchParams(location.search).get("api") ||
  "http://localhost:7860";

const form = document.getElementById("gen-form");
const statusAlert = document.getElementById("statusAlert");
const dlEl = document.getElementById("download");
const submitBtn = document.getElementById("submit-btn");

const qTotalRange = document.getElementById("q-total");
const qTotalNum = document.getElementById("q-total-num");
const customBox = document.getElementById("custom-box");
const customSum = document.getElementById("custom-sum");

const progressWrap = document.getElementById("progressWrap");
const progressBar = document.getElementById("progressBar");

// sync totals
function syncTotals(fromRange) {
  const val = parseInt(fromRange ? qTotalRange.value : qTotalNum.value, 10) || 1;
  const clamped = Math.max(1, Math.min(50, val));
  qTotalRange.value = clamped;
  qTotalNum.value = clamped;
  if (!customBox.classList.contains("d-none")) updateCustomSum();
}
qTotalRange.addEventListener("input", () => syncTotals(true));
qTotalNum.addEventListener("input", () => syncTotals(false));

// mix mode toggle
form.addEventListener("change", (e) => {
  if (e.target.name === "mix_mode") {
    const isCustom = e.target.value === "custom";
    customBox.classList.toggle("d-none", !isCustom);
    updateCustomSum();
  }
});

function updateCustomSum() {
  if (customBox.classList.contains("d-none")) return;
  const mcq = +form.mcq_n.value || 0;
  const th = +form.theory_n.value || 0;
  const cf = +form.codefill_n.value || 0;
  const fb = +form.fillblank_n.value || 0;
  const total = +qTotalNum.value || 0;
  customSum.textContent = `Sum: ${mcq + th + cf + fb} / ${total}`;
}
customBox.querySelectorAll("input").forEach((input) => {
  input.addEventListener("input", updateCustomSum);
});

// progress bar
let timer = null;
function startProgress() {
  progressWrap.classList.remove("d-none");
  progressBar.style.width = "2%";
  progressBar.classList.add("progress-bar-animated");
  let pct = 2;
  timer = setInterval(() => {
    pct = Math.min(90, pct + Math.random() * 6);
    progressBar.style.width = pct + "%";
  }, 250);
}
function finishProgress(success = true) {
  if (timer) clearInterval(timer);
  progressBar.classList.remove("progress-bar-animated");
  progressBar.style.width = "100%";
  progressBar.classList.toggle("bg-success", success);
  progressBar.classList.toggle("bg-danger", !success);
  setTimeout(() => {
    progressWrap.classList.add("d-none");
    progressBar.style.width = "0%";
    progressBar.classList.remove("bg-success", "bg-danger");
  }, 1200);
}

// status
function showStatus(message, type = "info") {
  statusAlert.className = `alert alert-${type}`;
  statusAlert.textContent = message;
}

// form submit
form.addEventListener("submit", async (e) => {
  e.preventDefault();
  dlEl.innerHTML = "";
  showStatus("Uploading and generating…", "info");
  submitBtn.disabled = true;
  startProgress();

  const fd = new FormData();
  const file = form.querySelector('input[type="file"]').files[0];
  if (!file) {
    finishProgress(false);
    showStatus("Please choose a file.", "warning");
    submitBtn.disabled = false;
    return;
  }
  fd.append("file", file);

  fd.append("total_questions", qTotalNum.value);
  fd.append("mix_mode", form.mix_mode.value);
  fd.append("difficulty", form.querySelector('input[name="difficulty"]:checked').value);
  fd.append("include_explanations", form.include_explanations.checked ? "true" : "false");
  fd.append("include_thumbs", form.include_explanations.checked ? "true" : "false");

  if (form.mix_mode.value === "custom") {
    const customTotal = (+form.mcq_n.value || 0) + (+form.theory_n.value || 0) + (+form.codefill_n.value || 0) + (+form.fillblank_n.value || 0);
    if (customTotal !== (+qTotalNum.value || 0)) {
      finishProgress(false);
      showStatus("Custom question counts must add up to the total.", "warning");
      submitBtn.disabled = false;
      return;
    }
    fd.append("mcq_n", form.mcq_n.value || "0");
    fd.append("theory_n", form.theory_n.value || "0");
    fd.append("codefill_n", form.codefill_n.value || "0");
    fd.append("fillblank_n", form.fillblank_n.value || "0");
  }

  if (form.OpenAI_api_key && form.OpenAI_api_key.value) {
    fd.append("openai_api_key", form.OpenAI_api_key.value);
  }

  try {
    const res = await fetch(`${BACKEND_BASE_URL}/generate`, { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok || data.status !== "ok") throw new Error(data.detail || data.message || "Generation failed");

    showStatus("Done! Your deck is ready.", "success");
    const a = document.createElement("a");
    a.href = data.url;
    a.textContent = "Download PPTX";
    a.className = "btn btn-outline-accent";
    a.download = data.filename;
    dlEl.innerHTML = "";
    dlEl.appendChild(a);
    finishProgress(true);
  } catch (err) {
    console.error(err);
    showStatus(`Error: ${err.message}`, "danger");
    finishProgress(false);
  } finally {
    submitBtn.disabled = false;
  }
});

document.addEventListener("DOMContentLoaded", () => {
  const el = document.getElementById("hero-typer");
  if (el) el.textContent = "Quiz Slide Generator";
});

// init
updateCustomSum();
