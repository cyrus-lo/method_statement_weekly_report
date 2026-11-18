(() => {
  const bypassDialog = document.getElementById("step-2-bypass-dialog");
  if (bypassDialog) {
    let pendingStep2Form = null;
    const confirmButton = bypassDialog.querySelector("[data-step-2-confirm]");
    const cancelButton = bypassDialog.querySelector("[data-step-2-cancel]");

    document.querySelectorAll('form[action$="/runs/ecm-step-2"]').forEach((form) => {
      form.addEventListener("submit", (event) => {
        const periodRow = form.closest("tr");
        if (
          periodRow?.dataset.stepOneBState !== "not-started" ||
          form.dataset.step1bBypassConfirmed === "true"
        ) return;

        event.preventDefault();
        pendingStep2Form = form;
        bypassDialog.showModal();
      });
    });

    confirmButton?.addEventListener("click", () => {
      const form = pendingStep2Form;
      if (!form) return;

      let confirmation = form.querySelector('[name="confirm_step_1b_bypass"]');
      if (!confirmation) {
        confirmation = document.createElement("input");
        confirmation.type = "hidden";
        confirmation.name = "confirm_step_1b_bypass";
        form.append(confirmation);
      }
      confirmation.value = "true";
      form.dataset.step1bBypassConfirmed = "true";
      pendingStep2Form = null;
      bypassDialog.close();
      form.requestSubmit();
    });

    cancelButton?.addEventListener("click", () => {
      pendingStep2Form = null;
      bypassDialog.close();
    });
    bypassDialog.addEventListener("cancel", () => {
      pendingStep2Form = null;
    });
    bypassDialog.addEventListener("close", () => {
      pendingStep2Form = null;
    });
  }

  const detail = document.querySelector("[data-run-id]");
  if (!detail || !["queued", "running"].includes(detail.dataset.runState)) return;

  const runId = detail.dataset.runId;
  const stateElement = document.getElementById("run-state");
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
      item.textContent = "No saved artifacts.";
      artifactsElement.append(item);
      return;
    }
    artifacts.forEach((artifact) => {
      const item = document.createElement("li");
      if (artifact.view_url) {
        const viewLink = document.createElement("a");
        viewLink.href = artifact.view_url;
        viewLink.target = "_blank";
        viewLink.rel = "noopener";
        viewLink.textContent = `Open CSV · ${artifact.file_name}`;
        const separator = document.createTextNode(" · ");
        const downloadLink = document.createElement("a");
        downloadLink.href = artifact.download_url;
        downloadLink.textContent = "Download";
        item.append(viewLink, separator, downloadLink);
      } else {
        const link = document.createElement("a");
        link.href = artifact.download_url;
        link.textContent = `${artifact.name} · ${artifact.file_name}`;
        item.append(link);
      }
      artifactsElement.append(item);
    });
  }

  let delay = 1500;
  let timer;
  async function refresh() {
    try {
      const response = await fetch(`/runs/${runId}/status`, { cache: "no-store" });
      if (!response.ok) throw new Error("Could not refresh this run.");
      const run = await response.json();
      delay = 1500;
      stateElement.textContent = run.state;
      stateElement.className = `status status-${run.state}`;
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
        return;
      }
      refreshElement.textContent = "Refreshing…";
    } catch (error) {
      delay = Math.min(delay * 2, 10000);
      refreshElement.textContent = "Refresh temporarily unavailable; retrying…";
    }
    timer = window.setTimeout(refresh, delay);
  }

  refresh();
})();
