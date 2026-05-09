/* global Office, Word */
"use strict";

let styleData = {};

// ── Office init ────────────────────────────────────────────────────────────

Office.onReady(async function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("btnLoadUrl").addEventListener("click", onLoadUrl);
    await loadStyleData();
    setupSelectionListener();
  }
});

// ── Style data loading ─────────────────────────────────────────────────────

async function loadStyleData() {
  try {
    const url = await readStyleHelpUrl();
    if (url) {
      document.getElementById("urlInput").value = url;
      await fetchStyleData(url);
    } else {
      showStatus("Set a URL below to load style definitions");
    }
  } catch (err) {
    console.error("StyleHelp: loadStyleData →", err);
    showStatus("Error reading document properties");
  }
}

async function readStyleHelpUrl() {
  return Word.run(async (context) => {
    const props = context.document.properties.customProperties;
    props.load("items");
    await context.sync();
    const match = props.items.find((p) => p.key === "StyleHelpURL");
    return match ? match.value : null;
  });
}

async function fetchStyleData(url) {
  showStatus("Loading…");
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  styleData = await response.json();
  hideUrlInput();
  showStatus(url.split("/").pop() || "Loaded");
  await updateDisplay();
}

// ── Selection listener ─────────────────────────────────────────────────────

function setupSelectionListener() {
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    debounce(updateDisplay, 150)
  );
  updateDisplay();
}

async function updateDisplay() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const firstPara = selection.paragraphs.getFirst();
      firstPara.load("style");
      await context.sync();
      renderStyle(firstPara.style);
    });
  } catch (err) {
    console.error("StyleHelp: updateDisplay →", err);
  }
}

// ── Render ─────────────────────────────────────────────────────────────────

function renderStyle(styleName) {
  document.getElementById("styleName").textContent = styleName || "";

  const data = styleData[styleName] || {};
  document.getElementById("titleField").value = data.title || "";
  document.getElementById("descriptionField").value = data.description || "";

  const swatch = document.getElementById("colorSwatch");
  swatch.style.backgroundColor = data.color || "transparent";
}

// ── URL input ──────────────────────────────────────────────────────────────

async function onLoadUrl() {
  const url = document.getElementById("urlInput").value.trim();
  if (!url) return;
  try {
    await fetchStyleData(url);
  } catch (err) {
    showStatus(`Error: ${err.message}`);
  }
}

function showUrlInput() {
  document.getElementById("urlSection").style.display = "block";
}

function hideUrlInput() {
  document.getElementById("urlSection").style.display = "none";
}

// ── UI helpers ─────────────────────────────────────────────────────────────

function showStatus(msg) {
  document.getElementById("statusText").textContent = msg;
}

function debounce(fn, ms) {
  let timer;
  return function (...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), ms);
  };
}
