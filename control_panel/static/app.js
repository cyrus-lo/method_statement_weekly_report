(() => {
  const selectAll = document.querySelector("[data-select-all-sources]");
  const sourceBoxes = [...document.querySelectorAll("[data-source-sheet]")];
  if (selectAll) {
    const syncSelectAll = () => {
      const selected = sourceBoxes.filter((box) => box.checked).length;
      selectAll.checked = selected === sourceBoxes.length;
      selectAll.indeterminate = selected > 0 && selected < sourceBoxes.length;
    };
    selectAll.addEventListener("change", () => {
      sourceBoxes.forEach((box) => { box.checked = selectAll.checked; });
      syncSelectAll();
    });
    sourceBoxes.forEach((box) => box.addEventListener("change", syncSelectAll));
    syncSelectAll();
  }

  const tabRoot = document.querySelector("[data-tabs]");
  if (tabRoot) {
    const tabs = [...tabRoot.querySelectorAll('[role="tab"]')];

    function selectTab(tab) {
      tabs.forEach((item) => {
        const selected = item === tab;
        item.setAttribute("aria-selected", String(selected));
        item.tabIndex = selected ? 0 : -1;
        const panel = document.getElementById(item.getAttribute("aria-controls"));
        if (panel) panel.hidden = !selected;
      });
    }

    tabs.forEach((tab, index) => {
      tab.addEventListener("click", () => selectTab(tab));
      tab.addEventListener("keydown", (event) => {
        if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
        event.preventDefault();
        const direction = ["ArrowUp", "ArrowLeft"].includes(event.key) ? -1 : ["ArrowDown", "ArrowRight"].includes(event.key) ? 1 : 0;
        const nextIndex = event.key === "Home" ? 0 : event.key === "End" ? tabs.length - 1 : (index + direction + tabs.length) % tabs.length;
        tabs[nextIndex].focus();
        selectTab(tabs[nextIndex]);
      });
    });
  }

  const detail = document.querySelector("[data-run-id]");
  if (!detail || !["queued", "running"].includes(detail.dataset.runState)) return;

  const runId = detail.dataset.runId;
  const stateElement = document.getElementById("run-state");
  const startedAt = document.getElementById("started-at");
  const finishedAt = document.getElementById("finished-at");
  const exitCode = document.getElementById("exit-code");
  const logElement = document.getElementById("run-log");
  const errorElement = document.getElementById("run-error");
  const artifactsElement = document.getElementById("artifact-list");
  const refreshElement = document.getElementById("log-refresh");

  function replaceArtifacts(artifacts) {
    artifactsElement.replaceChildren();
    if (!artifacts.length) {
      const item = document.createElement("li");
      item.className = "empty-artifacts";
      item.textContent = "Artifacts appear after a successful run.";
      artifactsElement.append(item);
      return;
    }
    artifacts.forEach((artifact) => {
      const item = document.createElement("li");
      const link = document.createElement("a");
      link.href = artifact.download_url;
      link.textContent = `${artifact.name} — ${artifact.file_name}`;
      item.append(link);
      artifactsElement.append(item);
    });
  }

  async function refresh() {
    try {
      const response = await fetch(`/runs/${runId}/status`, { cache: "no-store" });
      if (!response.ok) throw new Error("Could not refresh this run.");
      const run = await response.json();
      stateElement.textContent = run.state;
      stateElement.className = `status status-${run.state}`;
      startedAt.textContent = run.started_at || "Waiting to start";
      finishedAt.textContent = run.finished_at || "In progress";
      exitCode.textContent = run.exit_code ?? "—";
      logElement.textContent = run.log;
      logElement.scrollTop = logElement.scrollHeight;
      replaceArtifacts(run.artifacts);
      if (run.error_summary) {
        errorElement.classList.remove("hidden");
        errorElement.textContent = `Run failed. ${run.error_summary}`;
      }
      if (!["queued", "running"].includes(run.state)) {
        refreshElement.textContent = "Complete";
        window.clearInterval(interval);
      }
    } catch (error) {
      refreshElement.textContent = error.message;
      window.clearInterval(interval);
    }
  }

  const interval = window.setInterval(refresh, 1500);
  refresh();
})();
