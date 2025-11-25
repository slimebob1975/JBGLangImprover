fetch("/config")
  .then(res => res.json())
  .then(data => {
    document.title = data.title;
    document.getElementById("app-title").innerText = data.title;
  });

// Load saved API key from localStorage on page load
document.addEventListener("DOMContentLoaded", () => {
    const savedKey = localStorage.getItem("openai_api_key");
    if (savedKey) {
      document.getElementById("apiKey").value = savedKey;
    }

    const temperatureSlider = document.getElementById("temperature");
    const tempDisplay = document.getElementById("tempDisplay");
    const modelSelect = document.getElementById("model");

    // Reagera när modellen ändras
    if (modelSelect) {
        modelSelect.addEventListener("change", applyTemperaturePolicy);
        applyTemperaturePolicy();
    }

    // Initiera rätt läge vid sidladdning
    applyTemperaturePolicy(modelSelect, temperatureSlider, tempDisplay);

    if (temperatureSlider && tempDisplay) {
        temperatureSlider.addEventListener("input", function () {
            tempDisplay.textContent = this.value;
        });
    }

    // Fetch editable prompt
    fetch('/get_editable_prompt/')
    .then(response => response.json())
    .then(data => {
        document.getElementById("editablePrompt").value = data.editable_prompt;
    })
    .catch(error => {
        console.error("Failed to fetch editable prompt:", error);
    });
  });

function applyTemperaturePolicy() {
    console.log('applytemperaturePolicy')

     // Get fresh references to get rid of scope problems
    const modelSelect       = document.getElementById("model");
    const temperatureSlider = document.getElementById("temperature");
    const tempDisplay       = document.getElementById("tempDisplay");

    if (!modelSelect || !temperatureSlider || !tempDisplay) return;
    console.log('Not fast return')

    const isGPT5 = modelSelect.value.startsWith("gpt-5");
    if (isGPT5) {
        console.log('model choice was gpt-5-based')
        // Lås till 1 för GPT-5-modeller
        temperatureSlider.value = "1";
        tempDisplay.textContent = "1";
        temperatureSlider.disabled = true;
        temperatureSlider.setAttribute("aria-disabled", "true");
        temperatureSlider.title = "Låst till 1 för GPT-5-modeller";
    } else {
        console.log('model choice was not gpt-5-based')
        // Återställ för övriga modeller
        temperatureSlider.disabled = false;
        temperatureSlider.removeAttribute("aria-disabled");
        temperatureSlider.title = "";
        temperatureSlider.value = "0.7";
        tempDisplay.textContent = "0.7";
    }
}

document.getElementById("uploadForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const fileInput = document.getElementById("documentFile");
    const file = fileInput.files[0];
    const apiKey = document.getElementById("apiKey").value.trim();
    const model = document.getElementById("model").value;
    const editablePrompt = document.getElementById("editablePrompt").value.trim();
    const temperature = parseFloat(document.getElementById("temperature").value);
    const includeMotivations = document.getElementById("includeMotivations").checked;
    const docxMode = document.querySelector('input[name="docxMode"]:checked').value;

    if (!file || !apiKey) {
        updateStatus("❌ Du måste välja en fil och ange din API-nyckel.", "status-error");
        return;
    }

    lockUI();
    updateStatus("🔄 Bearbetar dokument...", "status-info");

    const formData = new FormData();
    formData.append("file", file);
    formData.append("api_key", apiKey);
    formData.append("model", model);
    formData.append("editable_prompt", editablePrompt);
    formData.append("temperature", temperature);
    formData.append("include_motivations", includeMotivations);
    formData.append("docx_mode", docxMode);

    try {
        const res = await fetch("/upload/", {
            method: "POST",
            body: formData
        });

        const { job_id, original_filename } = await res.json();
        pollForResult(job_id, original_filename);

    } catch (err) {
        console.error(err);
        updateStatus("❌ Tekniskt fel vid överföring.", "status-error");
    }
});

async function pollForResult(jobId, originalFilename) {
    const spinner = document.getElementById("spinner-container");
    spinner.style.display = "block";

    const interval = setInterval(async () => {
        console.log(`Kollar status för job ID: ${jobId}...`);
        try {
            const res = await fetch(`/status/${jobId}`);
            const data = await res.json();

            // Show status message if present
            if (data.status) {
                updateStatus(data.status, "status-info");
            }

            // Error from server
            if (data.error) {
                clearInterval(interval);
                console.warn(`Jobb ${jobId} rapporterade fel:`, data.error);
                updateStatus(
                    data.status || "Ett fel uppstod vid språkgranskningen.",
                    "status-error"
                );
                spinner.style.display = "none";
                return;
            }

            // Not done yet
            if (data.done === false) {
                console.log(`Jobb ${jobId} är fortfarande under bearbetning...`);
                return; // keep polling
            }

            // Done
            if (data.done === true) {
                console.log(`Jobb ${jobId} är klart. Startar nedladdning.`);
                clearInterval(interval);
                await downloadResult(jobId, originalFilename);
                return;
            }

            // Fallback if response shape is unexpected
            console.warn("Oväntat svar från /status:", data);
        } catch (err) {
            console.warn(`Misslyckades med att hämta status för ${jobId}:`, err);
        }
    }, 5000);  // Poll every 5 seconds (adjust if you like)
}

async function downloadResult(jobId, originalFilename) {
    try {
        const response = await fetch(`/download/${jobId}`);
        const blob = await response.blob();
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);

        const extension = originalFilename.slice(originalFilename.lastIndexOf("."));
        const base = originalFilename.slice(0, originalFilename.lastIndexOf("."));
        a.download = `${base}_klarspråkad${extension}`;

        a.click();

        updateStatus("✅ Färdig! Filen laddades ner.", "status-success");
    } catch (err) {
        console.error("Download failed:", err);
        updateStatus("❌ Kunde inte hämta resultatfil.", "status-error");
    } finally {
        document.getElementById("spinner-container").style.display = "none";
    }
}

function updateStatus(message, className) {
    const status = document.getElementById("status");
    status.className = className;
    status.textContent = message;
}

function lockUI() {
    const ids = ["documentFile", "apiKey", "model", "editablePrompt", "temperature", "includeMotivations", "simpleMarking", "trackedChanges", "button"];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = true;
    });
}


document.getElementById("uploadForm").addEventListener("submit_old", async (e) => {
    e.preventDefault();

    const fileInput = document.getElementById("documentFile");
    const file = fileInput.files[0];
    const apiKey = document.getElementById("apiKey").value.trim();
    const model = document.getElementById("model").value;
    const editablePrompt = document.getElementById("editablePrompt").value.trim();
    const temperature = parseFloat(document.getElementById("temperature").value);
    const includeMotivations = document.getElementById("includeMotivations").checked;
    const docxMode = document.querySelector('input[name="docxMode"]:checked').value;

    // Lock all elements on page
    document.getElementById("documentFile").disabled = true;
    document.getElementById("apiKey").disabled = true;
    document.getElementById("model").disabled = true;
    document.getElementById("editablePrompt").disabled = true;
    document.getElementById("temperature").disabled = true;
    document.getElementById("includeMotivations").disabled = true;
    document.getElementById("simpleMarking").disabled = true;
    document.getElementById("trackedChanges").disabled = true;
    document.getElementById("button").disabled = true;


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
    formData.append("editable_prompt", editablePrompt);
    formData.append("temperature", temperature);
    formData.append("include_motivations", includeMotivations);
    formData.append("docx_mode", docxMode);

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
    