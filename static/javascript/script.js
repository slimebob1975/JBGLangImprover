// Load saved API key from localStorage on page load
document.addEventListener("DOMContentLoaded", () => {
    const savedKey = localStorage.getItem("openai_api_key");
    if (savedKey) {
      document.getElementById("apiKey").value = savedKey;
    }
  });

document.getElementById("uploadForm").addEventListener("submit", async (e) => {
e.preventDefault();

const fileInput = document.getElementById("documentFile");
const file = fileInput.files[0];
const apiKey = document.getElementById("apiKey").value.trim();
const model = document.getElementById("model").value;
const customPrompt = document.getElementById("customPrompt").value.trim();

const status = document.getElementById("status");
const spinner = document.getElementById("spinner-container");

if (!file || !apiKey) {
    status.className = "status-error";
    status.textContent = "❌ Du måste välja en fil och ange din API-nyckel.";
    return;
}

status.className = "status-info";
status.textContent = "🔄 Bearbetar dokument...";
spinner.style.display = "block";

const formData = new FormData();
formData.append("file", file);
formData.append("api_key", apiKey);
formData.append("model", model);
formData.append("custom_prompt", customPrompt);

// Save the API key for future visits
localStorage.setItem("openai_api_key", apiKey);

try {
    const response = await fetch("/upload/", {
    method: "POST",
    body: formData
    });

    if (!response.ok) throw new Error("Upload failed");

    const blob = await response.blob();
    const downloadUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = downloadUrl;
    a.download = file.name.replace(/\.[^.]+$/, "_klarspråkad" + file.name.slice(file.name.lastIndexOf(".")));
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(downloadUrl);

    status.className = "status-success";
    status.textContent = "✅ Färdig! Filen laddades ner.";
} catch (err) {
    console.error("Upload error:", err);
    status.className = "status-error";
    status.textContent = "❌ Tekniskt fel vid överföring.";
} finally {
    spinner.style.display = "none";
}
});
  