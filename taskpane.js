/* global Office, Excel */

const KEYS = {
  status: "SheetNavigator.Status",
  statusBy: "SheetNavigator.StatusBy",
  statusAt: "SheetNavigator.StatusAt",
  changed: "SheetNavigator.ChangedAfterReview",
  history: "SheetNavigator.ChangeHistory"
};

const STATUS_LABELS = {
  none: "",
  prepared: "Prepared",
  reviewed: "Reviewed",
  openItems: "Open Items",
  signerReview: "Signer Review"
};

let sheetCache = [];
let activeSheetId = null;
let selectedSheetId = null;
let draggedSheetId = null;
let handlersRegistered = false;
let collapsed = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Open this add-in in Excel.", true);
    return;
  }

  document.getElementById("refresh")
    .addEventListener("click", loadSheets);

  document.getElementById("search")
    .addEventListener("input", renderSheets);

  document.getElementById("statusFilter")
    .addEventListener("change", renderSheets);

  document.getElementById("showHidden")
    .addEventListener("change", renderSheets);

  document.getElementById("collapseToggle")
    .addEventListener("click", toggleCollapsed);

  document.getElementById("closeHistory")
    .addEventListener("click", closeHistory);

  document.getElementById("historyOverlay")
    .addEventListener("click", (event) => {
      if (event.target.id === "historyOverlay") {
        closeHistory();
      }
    });

  document.addEventListener("click", handleOutsideClick);
  document.addEventListener("keydown", handleKeyboard);

  setupMenuActions();
  setupColorMenu();

  loadSheets();
});


/* =========================================================
   LOAD WORKSHEETS AND METADATA
   ========================================================= */

async function loadSheets() {
  setStatus("Loading...");

  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const activeSheet = sheets.getActiveWorksheet();

      sheets.load(
        "items/id,items/name,items/visibility," +
        "items/tabColor,items/position"
      );

      activeSheet.load("id");

      await context.sync();

      const loaded = [];

      for (const sheet of sheets.items) {
        const props = sheet.customProperties;
        props.load("items/key,items/value");

        loaded.push({
          sheet,
          props
        });
      }

      await context.sync();

      sheetCache = loaded.map(({ sheet, props }) => {
        const metadata = {};

        for (const property of props.items) {
          metadata[property.key] = property.value;
        }

        return {
          id: sheet.id,
          name: sheet.name,
          visibility: sheet.visibility,
          tabColor: sheet.tabColor || "",
          position: sheet.position,

          workflowStatus:
            metadata[KEYS.status] || "none",

          statusBy:
            metadata[KEYS.statusBy] || "",

          statusAt:
            metadata[KEYS.statusAt] || "",

          changedAfterReview:
            metadata[KEYS.changed] === "true",

          changeHistory:
            parseHistory(metadata[KEYS.history])
        };
      });

      sheetCache.sort(
        (left, right) =>
          left.position - right.position
      );

      activeSheetId = activeSheet.id;

      if (!handlersRegistered) {
        sheets.onActivated.add(onWorksheetActivated);
        sheets.onAdded.add(loadSheets);
        sheets.onDeleted.add(loadSheets);
        sheets.onChanged.add(onWorksheetChanged);

        handlersRegistered = true;

        await context.sync();
      }
    });

    setStatus("Ready");
    renderSheets();
    scrollActiveIntoView();

  } catch (error) {
    showError(error);
  }
}


function parseHistory(value) {
  if (!value) {
    return [];
  }

  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}


/* =========================================================
   WORKSHEET EVENTS
   ========================================================= */

function onWorksheetActivated(event) {
  activeSheetId = event.worksheetId;

  renderSheets();
  scrollActiveIntoView();
}


async function onWorksheetChanged(event) {
  const sheet = sheetCache.find(
    item => item.id === event.worksheetId
  );

  if (!sheet) {
    return;
  }

  /*
   * Prepared, Open Items, and No Status are intentionally
   * unaffected by later edits.
   */
  if (
    sheet.workflowStatus !== "reviewed" &&
    sheet.workflowStatus !== "signerReview"
  ) {
    return;
  }

  /*
   * Worksheet.onChanged captures edits to worksheet data.
   * Formula recalculation is not handled here. We do not
   * subscribe to onCalculated, so linked-formula result
   * changes do not create warnings.
   */
  const historyItem = buildHistoryItem(event);

  const newHistory = [
    historyItem,
    ...sheet.changeHistory
  ].slice(0, 100);

  try {
    await Excel.run(async (context) => {
      const worksheet =
        context.workbook.worksheets
          .getItem(event.worksheetId);

      await setCustomProperty(
        context,
        worksheet,
        KEYS.changed,
        "true"
      );

      await setCustomProperty(
        context,
        worksheet,
        KEYS.history,
        JSON.stringify(newHistory)
      );

      await context.sync();
    });

    sheet.changedAfterReview = true;
    sheet.changeHistory = newHistory;

    renderSheets();

  } catch (error) {
    console.error(
      "Could not record worksheet change:",
      error
    );
  }
}


function buildHistoryItem(event) {
  const item = {
    timestamp: new Date().toISOString(),
    address: event.address || "Unknown range",
    changeType:
      String(event.changeType || "Range edited"),
    before: "",
    after: ""
  };

  /*
   * Detailed old/new information is generally available
   * only for single-cell changes.
   */
  if (event.details) {
    item.before =
      formatChangeValue(
        event.details.valueBefore
      );

    item.after =
      formatChangeValue(
        event.details.valueAfter
      );
  }

  return item;
}


function formatChangeValue(value) {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  if (typeof value === "object") {
    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }

  return String(value);
}


/* =========================================================
   RENDERING
   ========================================================= */

function renderSheets() {
  const searchText =
    document.getElementById("search")
      .value
      .trim()
      .toLowerCase();

  const statusFilter =
    document.getElementById("statusFilter")
      .value;

  const showHidden =
    document.getElementById("showHidden")
      .checked;

  const filtered = sheetCache.filter((sheet) => {
    const matchesName =
      sheet.name.toLowerCase()
        .includes(searchText);

    const matchesVisibility =
      showHidden ||
      sheet.visibility === "Visible";

    const matchesStatus =
      statusMatches(sheet, statusFilter);

    return (
      matchesName &&
      matchesVisibility &&
      matchesStatus
    );
  });

  const list =
    document.getElementById("sheetList");

  list.replaceChildren();

  for (const sheet of filtered) {
    list.append(
      createSheetRow(sheet)
    );
  }

  document.getElementById("count")
    .textContent =
      `(${filtered.length}/${sheetCache.length})`;

  document.getElementById("empty")
    .hidden =
      filtered.length !== 0;
}


function createSheetRow(sheet) {
  const listItem =
    document.createElement("li");

  const row =
    document.createElement("div");

  row.className = "sheet-row";
  row.dataset.sheetId = sheet.id;
  row.draggable = true;

  if (sheet.id === activeSheetId) {
    row.classList.add("active");
  }

  if (sheet.visibility !== "Visible") {
    row.classList.add("hidden-sheet");
  }

  if (
    sheet.workflowStatus &&
    sheet.workflowStatus !== "none"
  ) {
    row.classList.add(
      `status-${sheet.workflowStatus}`
    );
  }

  row.addEventListener(
    "contextmenu",
    (event) => {
      event.preventDefault();

      openSheetMenuAtPoint(
        event.clientX,
        event.clientY,
        sheet.id
      );
    }
  );

  row.addEventListener(
    "dragstart",
    handleDragStart
  );

  row.addEventListener(
    "dragover",
    handleDragOver
  );

  row.addEventListener(
    "dragleave",
    handleDragLeave
  );

  row.addEventListener(
    "drop",
    handleDrop
  );

  row.addEventListener(
    "dragend",
    handleDragEnd
  );


  const dragHandle =
    document.createElement("span");

  dragHandle.className = "drag-handle";
  dragHandle.textContent = "⋮⋮";
  dragHandle.title = "Drag to reorder";


  const sheetButton =
    document.createElement("button");

  sheetButton.className = "sheet-button";
  sheetButton.title = sheet.name;

  sheetButton.addEventListener(
    "click",
    () => activateSheet(sheet.id)
  );

  sheetButton.addEventListener(
    "dblclick",
    (event) => {
      event.stopPropagation();
      renameSheet(sheet);
    }
  );


  const tabColor =
    document.createElement("span");

  tabColor.className = "tab-color";

  /*
   * This is the worksheet's actual Excel tab color.
   */
  tabColor.style.background =
    sheet.tabColor || "#c8c6c4";


  const sheetName =
    document.createElement("span");

  sheetName.className = "sheet-name";
  sheetName.textContent = sheet.name;

  sheetButton.append(
    tabColor,
    sheetName
  );

  if (sheet.visibility !== "Visible") {
    const hidden =
      document.createElement("span");

    hidden.className = "hidden-label";
    hidden.textContent = "Hidden";

    sheetButton.append(hidden);
  }


  const workflow =
    document.createElement("span");

  workflow.className =
    "workflow-label";

  workflow.textContent =
    STATUS_LABELS[
      sheet.workflowStatus
    ] || "";


  const warning =
    document.createElement("button");

  warning.className =
    "change-warning";

  if (sheet.changedAfterReview) {
    warning.textContent = "⚠";
    warning.title =
      "Changed after review. Click to view history.";

    warning.addEventListener(
      "click",
      (event) => {
        event.stopPropagation();
        openHistory(sheet);
      }
    );
  } else {
    warning.classList.add("empty");
    warning.textContent = "⚠";
  }

  row.append(
    dragHandle,
    sheetButton,
    workflow,
    warning
  );

  listItem.append(row);

  return listItem;
}


function statusMatches(
  sheet,
  filter
) {
  if (filter === "all") {
    return true;
  }

  if (filter === "changed") {
    return sheet.changedAfterReview;
  }

  if (filter === "none") {
    return (
      !sheet.workflowStatus ||
      sheet.workflowStatus === "none"
    );
  }

  return sheet.workflowStatus === filter;
}


/* =========================================================
   ACTIVATE AND COMPACT MODE
   ========================================================= */

async function activateSheet(sheetId) {
  try {
    await Excel.run(async (context) => {
      const sheet =
        context.workbook.worksheets
          .getItem(sheetId);

      sheet.activate();

      await context.sync();
    });

    activeSheetId = sheetId;

    renderSheets();
    scrollActiveIntoView();

  } catch (error) {
    showError(error);
  }
}


function toggleCollapsed() {
  collapsed = !collapsed;

  const app =
    document.querySelector(".app");

  const button =
    document.getElementById(
      "collapseToggle"
    );

  app.classList.toggle(
    "collapsed",
    collapsed
  );

  button.textContent =
    collapsed ? "›" : "‹";

  button.title =
    collapsed
      ? "Expand navigator"
      : "Collapse navigator";

  setStatus(
    collapsed ? "Collapsed" : "Ready"
  );
}


/* =========================================================
   RIGHT-CLICK MENU
   ========================================================= */

function openSheetMenuAtPoint(
  x,
  y,
  sheetId
) {
  selectedSheetId = sheetId;

  const menu =
    document.getElementById("sheetMenu");

  const sheet =
    getSelectedSheet();

  if (!sheet) {
    return;
  }

  const visibilityButton =
    menu.querySelector(
      '[data-action="visibility"]'
    );

  visibilityButton.textContent =
    sheet.visibility === "Visible"
      ? "Hide Sheet"
      : "Unhide Sheet";

  menu.hidden = false;

  const menuWidth = 190;
  const menuHeight =
    Math.min(menu.offsetHeight, 520);

  menu.style.left =
    `${Math.min(
      x,
      window.innerWidth - menuWidth - 5
    )}px`;

  menu.style.top =
    `${Math.min(
      y,
      window.innerHeight - menuHeight - 5
    )}px`;

  document.getElementById(
    "colorMenu"
  ).hidden = true;
}


function setupMenuActions() {
  document.querySelectorAll(
    "#sheetMenu button"
  ).forEach((button) => {
    button.addEventListener(
      "click",
      async (event) => {
        event.stopPropagation();

        const action =
          button.dataset.action;

        const sheet =
          getSelectedSheet();

        if (!sheet) {
          return;
        }
setStatus(
  `Clicked: ${action} ${button.dataset.status || ""}`
);

        if (action === "color") {
          openColorMenu();
          return;
        }

        closeMenus();

        switch (action) {
          case "activate":
            await activateSheet(sheet.id);
            break;

          case "insert":
            await insertSheet(sheet);
            break;

          case "duplicate":
            await duplicateSheet(sheet);
            break;

          case "delete":
            await deleteSheet(sheet);
            break;

          case "rename":
            await renameSheet(sheet);
            break;

          case "visibility":
            await toggleVisibility(sheet);
            break;

          case "workflow":
            await setWorkflowStatus(
              sheet,
              button.dataset.status
            );
            break;

          case "history":
            openHistory(sheet);
            break;
        }
      }
    );
  });
}


function getSelectedSheet() {
  return sheetCache.find(
    sheet =>
      sheet.id === selectedSheetId
  );
}


/* =========================================================
   SHEET OPERATIONS
   ========================================================= */

async function insertSheet(relativeSheet) {
  const proposedName =
    window.prompt(
      "New worksheet name:",
      "New Sheet"
    );

  if (proposedName === null) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheets =
        context.workbook.worksheets;

      const relative =
        sheets.getItem(relativeSheet.id);

      relative.load("position");

      await context.sync();

      const added =
        proposedName.trim()
          ? sheets.add(proposedName.trim())
          : sheets.add();

      added.position =
        relative.position + 1;

      added.activate();

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


async function duplicateSheet(sheet) {
  try {
    await Excel.run(async (context) => {
      const worksheet =
        context.workbook.worksheets
          .getItem(sheet.id);

      const copy =
        worksheet.copy(
          Excel.WorksheetPositionType.after,
          worksheet
        );

      copy.activate();

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


async function deleteSheet(sheet) {
  if (sheetCache.length <= 1) {
    window.alert(
      "Excel must contain at least one worksheet."
    );
    return;
  }

  const confirmed =
    window.confirm(
      `Delete worksheet "${sheet.name}"?\n\n` +
      "This action may not be reversible."
    );

  if (!confirmed) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets
        .getItem(sheet.id)
        .delete();

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


async function renameSheet(sheet) {
  const newName =
    window.prompt(
      "Worksheet name:",
      sheet.name
    );

  if (newName === null) {
    return;
  }

  const cleaned =
    newName.trim();

  if (!cleaned) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      const worksheet =
        context.workbook.worksheets
          .getItem(sheet.id);

      worksheet.name = cleaned;

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


async function toggleVisibility(sheet) {
  const visibleCount =
    sheetCache.filter(
      item =>
        item.visibility === "Visible"
    ).length;

  if (
    sheet.visibility === "Visible" &&
    visibleCount <= 1
  ) {
    window.alert(
      "Excel must have at least one visible worksheet."
    );
    return;
  }

  try {
    await Excel.run(async (context) => {
      const worksheet =
        context.workbook.worksheets
          .getItem(sheet.id);

      worksheet.visibility =
        sheet.visibility === "Visible"
          ? "Hidden"
          : "Visible";

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


/* =========================================================
   WORKFLOW STATUS
   ========================================================= */

async function setWorkflowStatus(
  sheet,
  workflowStatus
) {
  const timestamp =
    workflowStatus === "none"
      ? ""
      : new Date().toISOString();

  try {
    await Excel.run(async (context) => {

      const worksheet =
        context.workbook.worksheets
          .getItem(sheet.id);

      await setCustomProperty(
        context,
        worksheet,
        KEYS.status,
        workflowStatus
      );

      await setCustomProperty(
        context,
        worksheet,
        KEYS.statusBy,
        ""
      );

      await setCustomProperty(
        context,
        worksheet,
        KEYS.statusAt,
        timestamp
      );

      /*
       * Re-marking Reviewed or Signer Review confirms
       * that outstanding changes have been addressed.
       */
      if (
        workflowStatus === "reviewed" ||
        workflowStatus === "signerReview"
      ) {
        await setCustomProperty(
          context,
          worksheet,
          KEYS.changed,
          "false"
        );

        await setCustomProperty(
          context,
          worksheet,
          KEYS.history,
          "[]"
        );
      }

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


/* =========================================================
   CUSTOM PROPERTIES
   ========================================================= */

async function setCustomProperty(
  context,
  worksheet,
  key,
  value
) {
  const properties =
    worksheet.customProperties;

  const existing =
    properties.getItemOrNullObject(key);

  existing.load("isNullObject");

  await context.sync();

  if (!existing.isNullObject) {
    existing.delete();
    await context.sync();
  }

  properties.add(
    key,
    String(value)
  );
}


/* =========================================================
   TAB COLORS
   ========================================================= */

function setupColorMenu() {
  document.querySelectorAll(
    ".color-choice"
  ).forEach((button) => {
    button.addEventListener(
      "click",
      async (event) => {
        event.stopPropagation();

        const color =
          button.dataset.color;

        closeMenus();

        await setTabColor(
          selectedSheetId,
          color
        );
      }
    );
  });
}


function openColorMenu() {
  const menu =
    document.getElementById("sheetMenu");

  const colorMenu =
    document.getElementById(
      "colorMenu"
    );

  const rect =
    menu.getBoundingClientRect();

  colorMenu.hidden = false;

  colorMenu.style.left =
    `${Math.max(
      5,
      rect.left - colorMenu.offsetWidth - 6
    )}px`;

  colorMenu.style.top =
    `${Math.min(
      rect.top,
      window.innerHeight -
      colorMenu.offsetHeight - 5
    )}px`;
}


async function setTabColor(
  sheetId,
  color
) {
  try {
    await Excel.run(async (context) => {
      const worksheet =
        context.workbook.worksheets
          .getItem(sheetId);

      worksheet.tabColor = color;

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


/* =========================================================
   DRAG AND DROP
   ========================================================= */

function handleDragStart(event) {
  draggedSheetId =
    event.currentTarget.dataset.sheetId;

  event.currentTarget
    .classList.add("dragging");

  event.dataTransfer.effectAllowed =
    "move";
}


function handleDragOver(event) {
  event.preventDefault();

  event.dataTransfer.dropEffect =
    "move";

  event.currentTarget
    .classList.add("drag-over");
}


function handleDragLeave(event) {
  event.currentTarget
    .classList.remove("drag-over");
}


async function handleDrop(event) {
  event.preventDefault();

  const targetId =
    event.currentTarget.dataset.sheetId;

  event.currentTarget
    .classList.remove("drag-over");

  if (
    !draggedSheetId ||
    draggedSheetId === targetId
  ) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheets =
        context.workbook.worksheets;

      const source =
        sheets.getItem(draggedSheetId);

      const target =
        sheets.getItem(targetId);

      target.load("position");

      await context.sync();

      source.position = target.position;

      await context.sync();
    });

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


function handleDragEnd(event) {
  event.currentTarget
    .classList.remove("dragging");

  document.querySelectorAll(
    ".drag-over"
  ).forEach((element) => {
    element.classList.remove(
      "drag-over"
    );
  });

  draggedSheetId = null;
}


/* =========================================================
   CHANGE HISTORY
   ========================================================= */

function openHistory(sheet) {
  closeMenus();

  document.getElementById(
    "historySheetName"
  ).textContent =
    sheet.name;

  const reviewInfo =
    document.getElementById(
      "historyReviewInfo"
    );

  const statusLabel =
    STATUS_LABELS[
      sheet.workflowStatus
    ] || "No Status";

  const signedAt =
    formatDateTime(sheet.statusAt);

  reviewInfo.textContent =
    `${statusLabel}` +
    `${sheet.statusBy
      ? ` by ${sheet.statusBy}`
      : ""}` +
    `${signedAt
      ? ` • ${signedAt}`
      : ""}`;


  const list =
    document.getElementById(
      "historyList"
    );

  list.replaceChildren();

  for (
    const item of sheet.changeHistory
  ) {
    const entry =
      document.createElement("div");

    entry.className =
      "history-item";

    const time =
      document.createElement("div");

    time.className =
      "history-time";

    time.textContent =
      formatDateTime(item.timestamp);


    const address =
      document.createElement("div");

    address.className =
      "history-address";

    address.textContent =
      item.address ||
      "Unknown range";


    const values =
      document.createElement("div");

    values.className =
      "history-values";

    if (
      item.before !== "" ||
      item.after !== ""
    ) {
      values.textContent =
        `${item.before || "(blank)"} → ` +
        `${item.after || "(blank)"}`;
    } else {
      values.textContent =
        item.changeType ||
        "Range edited";
    }

    entry.append(
      time,
      address,
      values
    );

    list.append(entry);
  }


  document.getElementById(
    "historyEmpty"
  ).hidden =
    sheet.changeHistory.length !== 0;


  document.getElementById(
    "historyOverlay"
  ).hidden = false;
}


function closeHistory() {
  document.getElementById(
    "historyOverlay"
  ).hidden = true;
}


function formatDateTime(value) {
  if (!value) {
    return "";
  }

  const date =
    new Date(value);

  if (
    Number.isNaN(date.getTime())
  ) {
    return value;
  }

  return date.toLocaleString();
}


/* =========================================================
   GENERAL HELPERS
   ========================================================= */

function handleOutsideClick(event) {
  if (
    !event.target.closest("#sheetMenu") &&
    !event.target.closest("#colorMenu")
  ) {
    closeMenus();
  }
}


function handleKeyboard(event) {
  if (event.key === "Escape") {
    closeMenus();
    closeHistory();
  }
}


function closeMenus() {
  document.getElementById(
    "sheetMenu"
  ).hidden = true;

  document.getElementById(
    "colorMenu"
  ).hidden = true;
}


function scrollActiveIntoView() {
  requestAnimationFrame(() => {
    const row =
      document.querySelector(
        `.sheet-row[data-sheet-id="${activeSheetId}"]`
      );

    if (row) {
      row.scrollIntoView({
        block: "nearest",
        behavior: "smooth"
      });
    }
  });
}


function showError(error) {
  console.error(error);

  setStatus(
    error.message ||
    "Excel operation failed.",
    true
  );
}


function setStatus(
  message,
  isError = false
) {
  const status =
    document.getElementById("status");

  status.textContent = message;

  status.style.color =
    isError ? "#c50f1f" : "";
}