<script>
    document.getElementById("uploadForm").addEventListener("submit", async (e) => {
      e.preventDefault();
      const fileInput = document.getElementById("documentFile");
      const file = fileInput.files[0];
      if (!file) return;
  
      document.getElementById("status").textContent = "🔄 Bearbetar dokument...";
  
      const formData = new FormData();
      formData.append("file", file);
  
      const res = await fetch("/upload/", {
        method: "POST",
        body: formData
      });
  
      if (res.ok) {
        const blob = await res.blob();
        const downloadUrl = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = downloadUrl;
        a.download = file.name.replace(/\.[^.]+$/, "_improved" + file.name.slice(file.name.lastIndexOf(".")));
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(downloadUrl);
        document.getElementById("status").textContent = "✅ Färdig! Filen laddades ner.";
      } else {
        document.getElementById("status").textContent = "❌ Något gick fel.";
      }
    });
  </script>