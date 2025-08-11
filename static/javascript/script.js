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
  