(function () {
  const params = new URLSearchParams(window.location.search);
  const drtId = params.get("id");

  const createdDateEl = document.getElementById("loadDrtCreatedDate");
  if (createdDateEl) {
    createdDateEl.value = "Mar 18, 2025 10:30 AM PST";
  }

  const targetDateEl = document.getElementById("loadDrtTargetDate");
  if (targetDateEl) {
    const y = 2025;
    const m = String(Math.floor(Math.random() * 12) + 1).padStart(2, "0");
    const d = String(Math.floor(Math.random() * 28) + 1).padStart(2, "0");
    targetDateEl.value = `${m}/${d}/${y}`;
  }

  const loadDrtIdEl = document.getElementById("loadDrtId");
  if (loadDrtIdEl) {
    loadDrtIdEl.textContent = drtId || "DRT-" + String(Math.floor(Math.random() * 90000) + 10000);
  }

  const assignedToInput = document.getElementById("loadDrtAssignedTo");
  if (assignedToInput) {
    const queueData = window.DRT_QUEUE_DATA || [];
    const row = drtId ? queueData.find((r) => r.id === drtId) : null;
    const names = ["Maya Chen", "James Wilson", "Wilson Peddity", "Sarah Lee", "William Peddity", "Noor Patel", "Ishan Rao"];
    const placeholderName = (row && row.owner) ? row.owner : names[Math.floor(Math.random() * names.length)];
    assignedToInput.placeholder = placeholderName;
    assignedToInput.addEventListener("focus", () => { assignedToInput.placeholder = ""; });
    assignedToInput.addEventListener("blur", () => {
      if (!assignedToInput.value.trim()) assignedToInput.placeholder = placeholderName;
    });
  }

  const prioritySelect = document.getElementById("loadDrtPriority");
  if (prioritySelect) {
    function updatePriorityClass() {
      prioritySelect.classList.remove("priority-high", "priority-medium", "priority-low");
      const v = prioritySelect.value;
      if (v === "high") prioritySelect.classList.add("priority-high");
      else if (v === "medium") prioritySelect.classList.add("priority-medium");
      else if (v === "low") prioritySelect.classList.add("priority-low");
    }
    updatePriorityClass();
    prioritySelect.addEventListener("change", updatePriorityClass);
  }

  const resultSelect = document.getElementById("loadDrtResult");
  const closeBtn = document.getElementById("loadDrtCloseBtn");

  function updateResultClass() {
    if (!resultSelect) return;
    resultSelect.classList.remove("result-won", "result-lost");
    const v = resultSelect.value;
    if (v === "won") resultSelect.classList.add("result-won");
    else if (v === "lost") resultSelect.classList.add("result-lost");
  }

  function updateCloseButton() {
    if (!closeBtn || !resultSelect) return;
    closeBtn.disabled = !resultSelect.value;
  }

  if (resultSelect) {
    updateResultClass();
    updateCloseButton();
    resultSelect.addEventListener("change", () => {
      updateResultClass();
      updateCloseButton();
    });
  }

  if (closeBtn) {
    closeBtn.addEventListener("click", () => {
      if (closeBtn.disabled) return;
      window.location.href = "index.html";
    });
  }
})();
