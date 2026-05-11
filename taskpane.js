/* global Office, Word */
"use strict";

let styleData = {};
let scenarioList = [];
let scenarioBaseDir = "";

// ── Office init ────────────────────────────────────────────────────────────

Office.onReady(async function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("scenarioSelect").addEventListener("change", onScenarioChange);
    await init();
    setupSelectionListener();
  }
});

// ── Init ───────────────────────────────────────────────────────────────────

async function init() {
  try {
    const { baseUrl, preferredScenario } = await readDocProps();
    if (baseUrl) {
      await loadIndex(baseUrl, preferredScenario);
    } else {
      showStatus("No styleBaseUrl set in document properties");
    }
  } catch (err) {
    console.error("StyleHelp: init →", err);
    showStatus("Error loading style definitions");
  }
}

async function readDocProps() {
  return Word.run(async (context) => {
    const props = context.document.properties.customProperties;
    props.load("items");
    await context.sync();
    const find = (key) => {
      const p = props.items.find((p) => p.key === key);
      return p ? p.value : null;
    };
    return {
      baseUrl: find("styleBaseUrl"),
      preferredScenario: find("preferredScenario"),
    };
  });
}

// ── Index loading ──────────────────────────────────────────────────────────

async function loadIndex(indexUrl, preferredScenario) {
  showStatus("Loading…");
  scenarioBaseDir = indexUrl.substring(0, indexUrl.lastIndexOf("/") + 1);

  const response = await fetch(indexUrl);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  scenarioList = await response.json();

  populateDropdown(scenarioList, preferredScenario);

  const selectedId = document.getElementById("scenarioSelect").value;
  const scenario = scenarioList.find((s) => s.id === selectedId) || scenarioList[0];
  if (scenario) await loadScenario(scenario);
}

function populateDropdown(list, preferredId) {
  const select = document.getElementById("scenarioSelect");
  select.innerHTML = "";
  list.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s.id;
    opt.textContent = s.name;
    select.appendChild(opt);
  });
  if (preferredId && list.find((s) => s.id === preferredId)) {
    select.value = preferredId;
  }
}

// ── Scenario loading ───────────────────────────────────────────────────────

async function loadScenario(scenario) {
  const url = scenarioBaseDir + scenario.file;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const json = await response.json();
  styleData = json.styles || json;
  showStatus(scenario.name);
  await updateDisplay();
}

async function onScenarioChange() {
  const select = document.getElementById("scenarioSelect");
  const scenario = scenarioList.find((s) => s.id === select.value);
  if (!scenario) return;
  try {
    showStatus("Loading…");
    await loadScenario(scenario);
    await savePreferredScenario(scenario.id);
  } catch (err) {
    showStatus(`Error: ${err.message}`);
  }
}

async function savePreferredScenario(scenarioId) {
  try {
    await Word.run(async (context) => {
      context.document.properties.customProperties.add("preferredScenario", scenarioId);
      await context.sync();
    });
  } catch (err) {
    console.error("StyleHelp: savePreferredScenario →", err);
  }
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
