/* global Office, Excel */

let sheetCache = [];
let activeSheetId = null;
let handlersRegistered = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Open this add-in in Excel.", true);
    return;
  }

  document
    .getElementById("refresh")
    .addEventListener("click", loadSheets);

  document
    .getElementById("search")
    .addEventListener("input", renderSheets);

  document
    .getElementById("showHidden")
    .addEventListener("change", renderSheets);

  loadSheets();
});


async function loadSheets() {

  setStatus("Loading worksheets...");

  try {

    await Excel.run(async (context) => {

      const sheets = context.workbook.worksheets;
      const activeSheet = sheets.getActiveWorksheet();

      sheets.load(
        "items/id,items/name,items/visibility,items/tabColor"
      );

      activeSheet.load("id");

      await context.sync();

      sheetCache = sheets.items.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        visibility: sheet.visibility,
        tabColor: sheet.tabColor || "#bbb"
      }));

      activeSheetId = activeSheet.id;

      /*
       * Register these only once while the task pane is open.
       * This keeps the navigator synchronized with Excel.
       */
      if (!handlersRegistered) {

        sheets.onActivated.add(onWorksheetActivated);
        sheets.onAdded.add(loadSheets);
        sheets.onDeleted.add(loadSheets);

        handlersRegistered = true;

        await context.sync();
      }

    });

    setStatus("Ready");

    renderSheets();

  } catch (error) {

    console.error(error);

    setStatus(
      error.message || "Unable to load worksheets.",
      true
    );
  }
}


function onWorksheetActivated(event) {

  activeSheetId = event.worksheetId;

  renderSheets();
}


async function activateSheet(sheetId) {

  try {

    await Excel.run(async (context) => {

      const sheet =
        context.workbook.worksheets.getItem(sheetId);

      sheet.activate();

      await context.sync();

    });

    activeSheetId = sheetId;

    renderSheets();

  } catch (error) {

    console.error(error);

    setStatus(
      error.message || "Unable to activate worksheet.",
      true
    );
  }
}


function renderSheets() {

  const searchText =
    document
      .getElementById("search")
      .value
      .trim()
      .toLowerCase();

  const showHidden =
    document
      .getElementById("showHidden")
      .checked;


  const filteredSheets =
    sheetCache.filter((sheet) => {

      const matchesSearch =
        sheet.name
          .toLowerCase()
          .includes(searchText);

      const matchesVisibility =
        showHidden ||
        sheet.visibility === "Visible";

      return matchesSearch && matchesVisibility;

    });


  const list =
    document.getElementById("sheetList");

  list.replaceChildren();


  filteredSheets.forEach((sheet) => {

    const listItem =
      document.createElement("li");

    const button =
      document.createElement("button");


    button.className = "sheet";


    if (sheet.id === activeSheetId) {
      button.classList.add("active");
    }


    if (sheet.visibility !== "Visible") {
      button.classList.add("hidden");
    }


    button.title = sheet.name;


    button.addEventListener(
      "click",
      () => activateSheet(sheet.id)
    );


    /*
     * Excel sheet tab color
     */

    const colorDot =
      document.createElement("span");

    colorDot.className = "dot";

    colorDot.style.background =
      sheet.tabColor || "#bbb";


    /*
     * Worksheet name
     */

    const sheetName =
      document.createElement("span");

    sheetName.className = "name";

    sheetName.textContent =
      sheet.name;


    button.append(
      colorDot,
      sheetName
    );


    /*
     * Hidden worksheet indicator
     */

    if (sheet.visibility !== "Visible") {

      const badge =
        document.createElement("span");

      badge.className = "badge";

      badge.textContent =
        sheet.visibility;

      button.append(badge);
    }


    listItem.append(button);

    list.append(listItem);

  });


  /*
   * Worksheet count
   */

  document
    .getElementById("count")
    .textContent =
      `${filteredSheets.length} / ${sheetCache.length}`;


  /*
   * No search results message
   */

  document
    .getElementById("empty")
    .hidden =
      filteredSheets.length !== 0;
}


function setStatus(message, isError = false) {

  const status =
    document.getElementById("status");

  status.textContent = message;

  status.style.color =
    isError ? "#ffdddd" : "";
}