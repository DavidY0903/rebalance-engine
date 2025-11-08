const fileElem = document.getElementById("fileElem");
const dropArea = document.getElementById("drop-area");
const fileName = document.getElementById("file-name");
const rebalanceBtn = document.getElementById("rebalance-btn");
const statusMsg = document.getElementById("status-msg");
const loader = document.getElementById("loader");
const successBox = document.getElementById("success-box");
const downloadBtn = document.getElementById("download-btn");

// ✅ Dark Mode
const darkToggle = document.getElementById("darkToggle");
darkToggle.addEventListener("change", () => {
    document.body.classList.toggle("dark", darkToggle.checked);
});

// ✅ Drag & Drop Area
dropArea.addEventListener("click", () => fileElem.click());
['dragenter','dragover','dragleave','drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, e => e.preventDefault());
});

dropArea.addEventListener("drop", e => {
    fileElem.files = e.dataTransfer.files;
    handleFiles();
});

fileElem.addEventListener("change", handleFiles);

// ✅ File Selected
function handleFiles() {
    const file = fileElem.files[0];
    if (!file) return;

    fileName.textContent = file.name;
    statusMsg.textContent = "✅ File loaded（檔案已載入）";
    rebalanceBtn.disabled = false;
}

// ✅ Start Rebalance
rebalanceBtn.addEventListener("click", async () => {
    if (!fileElem.files.length) return;

    loader.classList.remove("hidden");
    successBox.classList.add("hidden");
    statusMsg.textContent = "🔄 Running…（計算中…）";

    const formData = new FormData();
    formData.append("file", fileElem.files[0]);

    try {
        const res = await fetch("/api/rebalance", {
            method: "POST",
            body: formData
        });

        if (!res.ok) throw new Error("Rebalance failed");

        // ✅ Receive output file from FastAPI
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);

        // ✅ Get filename from response header
        const disposition = res.headers.get("Content-Disposition");
        let filename = "rebalance.xlsx";

        if (disposition) {
            const match = disposition.match(/filename\*?=(?:UTF-8''|")?([^\";]+)/);
            if (match && match[1]) {
                try { filename = decodeURIComponent(match[1]); }
                catch { filename = match[1]; }
            }
        }

        // ✅ Enable download
        downloadBtn.href = url;
        downloadBtn.download = filename;

        loader.classList.add("hidden");
        successBox.classList.remove("hidden");
        statusMsg.textContent = "✅ Completed!（已完成）";

    } catch (err) {
        loader.classList.add("hidden");
        statusMsg.textContent = "❌ Error: " + err.message;
    }
});
