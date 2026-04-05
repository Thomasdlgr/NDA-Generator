(function () {
  "use strict";

  const f = document.getElementById("f");
  const go = document.getElementById("go");
  const panel = document.getElementById("progress-panel");
  const barWrap = document.getElementById("bar-wrap");
  const bar = document.getElementById("progress-bar");
  const lbl = document.getElementById("progress-label");
  const listEl = document.getElementById("issues-checklist");
  const logEl = document.getElementById("log");
  const dl = document.getElementById("dl");
  const previewPanel = document.getElementById("preview-panel");
  const previewFrame = document.getElementById("preview-frame");
  const inputNda = document.getElementById("file-nda");
  const inputPb = document.getElementById("file-playbook");
  const nameNda = document.getElementById("nda-name");
  const namePb = document.getElementById("playbook-name");
  const cardNda = document.getElementById("card-nda");
  const cardPb = document.getElementById("card-playbook");
  const issueDetail = document.getElementById("issue-detail");
  const issueDetailTitle = document.getElementById("issue-detail-title");
  const issueDetailPreferred = document.getElementById("issue-detail-preferred");
  const issueDetailFallback = document.getElementById("issue-detail-fallback");
  const issueDetailWording = document.getElementById("issue-detail-wording");
  const reportPanel = document.getElementById("report-panel");
  const reportBody = document.getElementById("report-body");
  const previewIssueSelect = document.getElementById("preview-issue-select");

  let es = null;
  /** @type {Record<number, number[]>} */
  let paragraphIndicesByIssue = {};

  function resetPreviewIssueSelect(issueTitles) {
    if (!previewIssueSelect) return;
    previewIssueSelect.innerHTML =
      '<option value="">— Toutes (aucun surlignage) —</option>';
    (issueTitles || []).forEach(function (title, i) {
      const opt = document.createElement("option");
      opt.value = String(i + 1);
      opt.textContent = title;
      previewIssueSelect.appendChild(opt);
    });
    previewIssueSelect.disabled = true;
    previewIssueSelect.value = "";
  }

  function postHighlightToPreview(indices) {
    const fr = previewFrame;
    if (!fr || !fr.contentWindow) return;
    fr.contentWindow.postMessage(
      { type: "nda-highlight", indices: indices || [] },
      "*"
    );
  }

  function postHighlightFromSelect() {
    if (!previewIssueSelect || !previewFrame || !previewFrame.contentWindow) return;
    const v = previewIssueSelect.value;
    const idx = v ? parseInt(v, 10) : NaN;
    const indices =
      !v || isNaN(idx) ? [] : paragraphIndicesByIssue[idx] || [];
    postHighlightToPreview(indices);
  }

  function clearIssueDetail() {
    if (!issueDetail) return;
    issueDetail.hidden = true;
    if (issueDetailTitle) issueDetailTitle.textContent = "";
    if (issueDetailPreferred) issueDetailPreferred.textContent = "";
    if (issueDetailFallback) issueDetailFallback.textContent = "";
    if (issueDetailWording) issueDetailWording.textContent = "";
  }

  function fillIssueDetail(data) {
    if (!issueDetail) return;
    function place(el, text) {
      if (!el) return;
      var s = text != null ? String(text).trim() : "";
      el.textContent = s || "—";
    }
    if (issueDetailTitle) issueDetailTitle.textContent = data.title || "";
    place(issueDetailPreferred, data.preferred_position);
    place(issueDetailFallback, data.fallback_position);
    place(issueDetailWording, data.preferred_wording);
    issueDetail.hidden = false;
  }

  function setFileName(span, file) {
    span.textContent = file ? file.name : "";
  }

  inputNda.addEventListener("change", function () {
    setFileName(nameNda, inputNda.files[0]);
  });
  inputPb.addEventListener("change", function () {
    setFileName(namePb, inputPb.files[0]);
  });

  function wireDrop(card, input) {
    ["dragenter", "dragover"].forEach(function (ev) {
      card.addEventListener(ev, function (e) {
        e.preventDefault();
        e.stopPropagation();
        card.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      card.addEventListener(ev, function (e) {
        e.preventDefault();
        e.stopPropagation();
        card.classList.remove("dragover");
      });
    });
    card.addEventListener("drop", function (e) {
      const files = e.dataTransfer && e.dataTransfer.files;
      if (!files || !files.length) return;
      const file = files[0];
      const okDocx =
        file.name.toLowerCase().endsWith(".docx") ||
        file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      const okXlsx =
        file.name.toLowerCase().endsWith(".xlsx") ||
        file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      if (input === inputNda && !okDocx) return;
      if (input === inputPb && !okXlsx) return;
      var dt = new DataTransfer();
      dt.items.add(file);
      input.files = dt.files;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }

  wireDrop(cardNda, inputNda);
  wireDrop(cardPb, inputPb);

  if (previewIssueSelect) {
    previewIssueSelect.addEventListener("change", function () {
      postHighlightFromSelect();
    });
  }
  if (previewFrame) {
    previewFrame.addEventListener("load", function () {
      if (previewIssueSelect) previewIssueSelect.disabled = false;
      postHighlightFromSelect();
    });
  }

  function setPct(p) {
    const n = Math.max(0, Math.min(100, Number(p) || 0));
    bar.style.width = n + "%";
    if (barWrap) {
      barWrap.setAttribute("aria-valuenow", String(Math.round(n)));
    }
    const done = document.querySelectorAll("#issues-checklist li.done").length;
    const total = listEl.children.length;
    lbl.textContent =
      n +
      " % — " +
      (total ? done + " / " + total + " issue(s) traitée(s)" : "préparation…");
  }

  f.addEventListener("submit", async function (e) {
    e.preventDefault();
    if (es) {
      es.close();
      es = null;
    }
    dl.style.display = "none";
    if (previewPanel) previewPanel.hidden = true;
    if (previewFrame) previewFrame.removeAttribute("src");
    logEl.textContent = "";
    clearIssueDetail();
    if (reportBody) reportBody.innerHTML = "";
    if (reportPanel) reportPanel.hidden = true;
    paragraphIndicesByIssue = {};
    resetPreviewIssueSelect([]);
    panel.hidden = false;
    listEl.innerHTML = "";
    setPct(0);
    lbl.textContent = "0 % — démarrage…";
    go.disabled = true;

    const fd = new FormData(f);
    if (!fd.get("strict_ops")) fd.delete("strict_ops");
    if (!fd.get("verbose")) fd.delete("verbose");

    try {
      const res = await fetch("/api/jobs", { method: "POST", body: fd });
      const j = await res.json();
      if (!res.ok) {
        logEl.textContent =
          typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail || j, null, 2);
        panel.hidden = true;
        go.disabled = false;
        return;
      }
      const jobId = j.job_id;
      resetPreviewIssueSelect(j.issues || []);
      (j.issues || []).forEach(function (title, i) {
        const li = document.createElement("li");
        li.dataset.index = String(i + 1);
        li.innerHTML =
          '<span class="chk pending" aria-hidden="true">…</span><span class="ttl"></span>';
        li.querySelector(".ttl").textContent = title;
        listEl.appendChild(li);
      });
      lbl.textContent = "0 % — " + listEl.children.length + " issue(s) à traiter";

      es = new EventSource("/api/jobs/" + encodeURIComponent(jobId) + "/events");
      es.onmessage = function (ev) {
        let data;
        try {
          data = JSON.parse(ev.data);
        } catch (_) {
          return;
        }
        if (data.kind === "init") setPct(data.percent ?? 0);
        if (data.kind === "issue_begin") {
          setPct(data.percent ?? 0);
          fillIssueDetail(data);
          document.querySelectorAll("#issues-checklist li").forEach(function (li) {
            li.classList.remove("current");
          });
          const li = listEl.querySelector('li[data-index="' + data.index + '"]');
          if (li) li.classList.add("current");
        }
        if (data.kind === "issue_end") {
          setPct(data.percent ?? 0);
          clearIssueDetail();
          if (typeof data.index === "number") {
            paragraphIndicesByIssue[data.index] = Array.isArray(
              data.paragraph_indices
            )
              ? data.paragraph_indices
              : [];
          }
          if (data.summary_html && reportBody && reportPanel) {
            reportPanel.hidden = false;
            reportBody.insertAdjacentHTML("beforeend", data.summary_html);
          }
          const li = listEl.querySelector('li[data-index="' + data.index + '"]');
          if (li) {
            li.classList.remove("current");
            li.classList.add("done");
            const st = data.status || "ok";
            const mark = st === "ok" ? "✓" : st === "no_ops" ? "○" : "!";
            const chk = li.querySelector(".chk");
            chk.textContent = mark;
            chk.classList.remove("pending");
            li.title = st;
          }
        }
        if (data.kind === "complete") {
          es.close();
          es = null;
          setPct(data.percent ?? (data.success ? 100 : 0));
          lbl.textContent =
            (data.percent ?? 0) +
            " % — " +
            (data.success ? "terminé" : "terminé avec erreurs");
          go.disabled = false;
          if (data.success) {
            const ndaFile = inputNda.files[0];
            const base =
              ndaFile && ndaFile.name
                ? ndaFile.name.replace(/\.docx$/i, "")
                : "NDA";
            if (previewPanel && previewFrame) {
              previewPanel.hidden = false;
              previewFrame.src =
                "/api/jobs/" + encodeURIComponent(jobId) + "/preview";
            }
            dl.href =
              "/api/jobs/" + encodeURIComponent(jobId) + "/download";
            dl.download = base + "_revu.docx";
            dl.textContent = "Télécharger le document revu";
            dl.style.display = "inline-block";
          } else {
            fetch("/api/jobs/" + encodeURIComponent(jobId) + "/log")
              .then(function (r) {
                return r.text();
              })
              .then(function (t) {
                logEl.textContent = t || "Échec — voir les logs serveur.";
              })
              .catch(function () {});
          }
        }
      };
      es.onerror = function () {
        if (es) es.close();
        es = null;
        go.disabled = false;
        logEl.textContent =
          (logEl.textContent || "") + "\nConnexion aux événements interrompue.";
      };
    } catch (err) {
      logEl.textContent = "Erreur : " + err;
      panel.hidden = true;
      go.disabled = false;
    }
  });
})();
