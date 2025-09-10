// Backend origin (override with ?api=https://your-api.com)
const BACKEND_BASE_URL =
  new URLSearchParams(location.search).get("api") ||
  "https://crystallizedcrust-quiz-generator.hf.space";

// ...

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

  if (form.mix_mode.value === "custom") {
    fd.append("mcq_n", form.mcq_n.value || "0");
    fd.append("theory_n", form.theory_n.value || "0");
    fd.append("codefill_n", form.codefill_n.value || "0");
    fd.append("fillblank_n", form.fillblank_n.value || "0");
  }

  // 🔄 changed from gemini_api_key → openai_api_key
  if (form.openai_api_key && form.openai_api_key.value) {
    fd.append("openai_api_key", form.openai_api_key.value);
  }

  try {
    const res = await fetch(`${BACKEND_BASE_URL}/generate`, { method: "POST", body: fd });
    const data = await res.json();

    if (!res.ok || data.status !== "ok") {
      throw new Error(data.detail || data.message || "Generation failed");
    }

    showStatus("Done! Your deck is ready.", "success");
    const a = document.createElement("a");
    a.href = data.url;
    a.textContent = "Download PPTX";
    a.className = "btn btn-outline-primary";
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

// initialize
updateCustomSum();
