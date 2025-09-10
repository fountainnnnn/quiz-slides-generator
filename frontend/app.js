// Backend origin (override with ?api=https://your-api.com)
const BACKEND_BASE_URL = (new URLSearchParams(location.search).get("api")) || "http://localhost:8000";

const form = document.getElementById("gen-form");
const statusEl = document.getElementById("status");
const dlEl = document.getElementById("download");
const submitBtn = document.getElementById("submit-btn");

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  statusEl.textContent = "Uploading and generating...";
  dlEl.innerHTML = "";
  submitBtn.disabled = true;

  const formData = new FormData(form);
  try {
    const res = await fetch(`${BACKEND_BASE_URL}/generate`, { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok || data.status !== "ok") {
      throw new Error(data.detail || data.message || "Generation failed");
    }
    statusEl.textContent = "Done! Your deck is ready.";
    const a = document.createElement("a");
    a.href = data.url;  // absolute URL returned by backend
    a.textContent = "Download PPTX";
    a.download = data.filename;
    dlEl.innerHTML = "";
    dlEl.appendChild(a);
  } catch (err) {
    statusEl.textContent = `Error: ${err.message}`;
  } finally {
    submitBtn.disabled = false;
  }
});
