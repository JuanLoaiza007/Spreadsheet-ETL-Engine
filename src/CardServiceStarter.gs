/**
 * @fileoverview
 * @version 0.0.1
 * @author JuanLoaiza007
 * @license MIT
 * @description Self-Healing Infrastructure for the Spreadsheet ETL Engine
 */

function onOpen() {
  getConfigSheet();
  SpreadsheetApp.getUi()
    .createMenu("Utilidades")
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Configuración")
        .addItem("1. Cambiar Fuente", "setSource")
        .addItem("2. Cambiar Mapa", "setMap")
        .addItem("3. Cambiar Salida", "setOutput")
        .addSeparator()
        .addItem("Ver Configuración Actual", "showCurrentConfig"),
    )
    .addSeparator()
    .addItem("Ejecutar ETL", "executeETL")
    .addToUi();
}

/** --- INFRASTRUCTURE MANAGEMENT --- **/
function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("_Config");

  if (!sheet) {
    sheet = ss.insertSheet("_Config");
    sheet.hideSheet();
    sheet.getRange("A1:B3").setValues([
      ["source", ""],
      ["map", ""],
      ["output", "Output"],
    ]);
  }
  return sheet;
}

function updateSingleConfig(key, value) {
  const config = getSavedConfig();
  config[key] = value;
  getConfigSheet()
    .getRange("A1:B3")
    .setValues([
      ["source", config.source],
      ["map", config.map],
      ["output", config.output],
    ]);
}

function getSavedConfig() {
  const data = getConfigSheet().getRange("A1:B3").getValues();
  return {
    source: data[0][1],
    map: data[1][1],
    output: data[2][1],
  };
}

/** --- UI AND EVENTS --- **/
function setSource() {
  const config = getSavedConfig();
  const selected = promptForSheetWithMarker(
    "Fuente",
    "Seleccione la hoja origen:",
    config.source,
  );
  if (selected) updateSingleConfig("source", selected);
}

function setMap() {
  const config = getSavedConfig();
  const selected = promptForSheetWithMarker(
    "Mapa",
    "Seleccione la hoja de mapeo:",
    config.map,
  );
  if (selected) updateSingleConfig("map", selected);
}

function setOutput() {
  const ui = SpreadsheetApp.getUi();
  const config = getSavedConfig();
  const res = ui.prompt(
    "Salida",
    `Actual: ${config.output}\n\nNuevo nombre:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() === ui.Button.OK) {
    updateSingleConfig("output", res.getResponseText().trim() || "Output");
  }
}

function promptForSheetWithMarker(title, message, currentSelection) {
  const ui = SpreadsheetApp.getUi();
  const sheets = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map((s) => s.getName())
    .filter((name) => name !== "_Config");

  const listText = sheets
    .map(
      (name, i) =>
        `${i + 1}. ${name}${name === currentSelection ? " [*]" : ""}`,
    )
    .join("\n");
  const res = ui.prompt(
    title,
    `${message}\n\n${listText}\n\nNúmero:`,
    ui.ButtonSet.OK_CANCEL,
  );

  if (res.getSelectedButton() !== ui.Button.OK) return null;
  const index = parseInt(res.getResponseText().trim(), 10) - 1;
  return index >= 0 && index < sheets.length
    ? sheets[index]
    : (ui.alert("Error de selección"), null);
}

function executeETL() {
  const config = getSavedConfig();
  if (!config.source || !config.map) {
    SpreadsheetApp.getUi().alert(
      "Error",
      "Configuración incompleta.",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
    return;
  }

  try {
    runMapping(config);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "Error",
      e.message,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

function showCurrentConfig() {
  const c = getSavedConfig();
  SpreadsheetApp.getUi().alert(
    "Configuración",
    `Fuente: ${c.source}\nMapa: ${c.map}\nSalida: ${c.output}`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
