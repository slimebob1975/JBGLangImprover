document.getElementById("uploadForm").addEventListener("submit", async (e) => {
e.preventDefault();
const fileInput = document.getElementById("documentFile");
const file = fileInput.files[0];
if (!file) return;

const status = document.getElementById("status");
const spinner = document.getElementById("spinner-container");

// Show spinner and info message
status.className = "status-info";
status.textContent = "Bearbetar dokument...";
spinner.style.display = "block";

const formData = new FormData();
formData.append("file", file);

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
