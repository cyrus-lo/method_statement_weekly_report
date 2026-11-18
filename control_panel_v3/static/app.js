(() => {
  const bypassDialog = document.querySelector("#bypass-dialog");
  let bypassForm = null;
  document.querySelectorAll("[data-open-bypass]").forEach((button) => {
    button.addEventListener("click", () => {
      bypassForm = document.getElementById(button.dataset.form);
      bypassDialog?.showModal();
    });
  });
  document.querySelector("[data-cancel-bypass]")?.addEventListener("click", () => bypassDialog?.close());
  document.querySelector("[data-confirm-bypass]")?.addEventListener("click", () => {
    bypassDialog?.close();
    bypassForm?.requestSubmit();
  });

  document.querySelectorAll(".mode-form").forEach((form) => {
    form.addEventListener("submit", (event) => {
      const modeButton = form.querySelector("button");
      if (modeButton?.textContent.includes("switch to live") && !window.confirm("Switch this V3 process to LIVE mode? Live workflows may modify V3-owned workbooks and launch browser automation.")) {
        event.preventDefault();
      }
    });
  });

  const statusCard = document.querySelector("[data-run-status-url][data-active='true']");
  if (!statusCard) return;
  const stateNode = document.querySelector("#run-state");
  const finishedNode = document.querySelector("#run-finished");
  const logNode = document.querySelector("#run-log");
  const artifactsNode = document.querySelector("#run-artifacts");
  let delay = 1000;
  let retryCount = 0;
  const maxRetries = 8;
  const maxDelay = 15000;
  const renderArtifacts = (items) => {
    if (!artifactsNode || !Array.isArray(items)) return;
    artifactsNode.replaceChildren();
    if (items.length === 0) {
      const empty = document.createElement("li");
      empty.textContent = "No artifacts.";
      artifactsNode.append(empty);
      return;
    }
    for (const item of items) {
      const li = document.createElement("li");
      const download = document.createElement("a");
      download.href = `/runs/${encodeURIComponent(statusCard.dataset.runStatusUrl.split("/").at(-2))}/artifacts/${encodeURIComponent(item.id)}`;
      download.textContent = `${item.name} · ${item.filename}`;
      li.append(download);
      if (item.is_csv) {
        const separator = document.createTextNode(" · ");
        const view = document.createElement("a");
        view.href = `${download.href}/view`;
        view.textContent = "Open CSV";
        li.append(separator, view);
      }
      artifactsNode.append(li);
    }
  };
  const poll = async () => {
    try {
      const response = await fetch(statusCard.dataset.runStatusUrl, { cache: "no-store", headers: { Accept: "application/json" } });
      if (!response.ok) throw new Error(`Status request returned ${response.status}`);
      const payload = await response.json();
      retryCount = 0;
      delay = 1000;
      if (stateNode) stateNode.textContent = payload.state;
      if (finishedNode && payload.finished_at) finishedNode.textContent = payload.finished_at;
      if (logNode && typeof payload.log === "string") logNode.textContent = payload.log;
      renderArtifacts(payload.artifacts);
      if (payload.error_summary && payload.state === "failed") {
        const alert = document.createElement("p");
        alert.className = "alert";
        alert.textContent = payload.error_summary;
        statusCard.append(alert);
      }
      if (["succeeded", "failed"].includes(payload.state)) {
        statusCard.dataset.active = "false";
        window.location.reload();
        return;
      }
    } catch (_error) {
      retryCount += 1;
      delay = Math.min(maxDelay, 1000 * (2 ** Math.min(retryCount, 4)));
      if (retryCount > maxRetries) return;
    }
    window.setTimeout(poll, delay);
  };
  window.setTimeout(poll, delay);
})();
