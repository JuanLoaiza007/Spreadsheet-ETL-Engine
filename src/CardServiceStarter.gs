function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Utilidades")
    .addItem("Configuración", "showForm")
    .addToUi();
}

function showForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map((s) => s.getName());
  const defaultMap = sheets.find((n) => n.includes("Map_")) || "";

  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="p-5 antialiased">
      <div class="space-y-4">
        
        <div>
          <label class="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1">Fuente</label>
          <select id="source" class="w-full p-2 border border-gray-300 rounded-md bg-white text-sm focus:ring-2 focus:ring-blue-500 outline-none">
            ${sheets.map((name) => `<option value="${name}">${name}</option>`).join("")}
          </select>
        </div>

        <div>
          <label class="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1">Mapa</label>
          <select id="map" class="w-full p-2 border border-gray-300 rounded-md bg-white text-sm focus:ring-2 focus:ring-blue-500 outline-none">
            ${sheets.map((name) => `<option value="${name}" ${name === defaultMap ? "selected" : ""}>${name}</option>`).join("")}
          </select>
        </div>

        <div>
          <label class="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1">Nombre de Salida</label>
          <input type="text" id="dest" value="Output" class="w-full p-2 border border-gray-300 rounded-md bg-white text-sm focus:ring-2 focus:ring-blue-500 outline-none">
        </div>

        <button onclick="ejecutar()" 
          class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-4 rounded-md transition duration-200 shadow-sm active:transform active:scale-95 mt-2">
          EJECUTAR
        </button>

      </div>

      <script>
        function ejecutar() {
          const data = {
            src: document.getElementById('source').value,
            map: document.getElementById('map').value,
            dst: document.getElementById('dest').value
          };
          
          const resumen = "Configuración:\\n\\n- Fuente: " + data.src + "\\n- Mapa: " + data.map + "\\n- Salida: " + data.dst;
          alert(resumen);
          google.script.host.close();
        }
      </script>
    </body>
    </html>
  `;

  const ui = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(ui, "Configuración del Proyecto");
}
