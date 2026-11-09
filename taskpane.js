/* global Office, Excel */


/* =========================================
   STATE
   ========================================= */

let sheetCache = [];
let activeSheetId = null;
let selectedSheetId = null;
let draggedSheetId = null;
let handlersRegistered = false;
let dialogResolver = null;
let navigatorCollapsed = false;


/* =========================================
   WORKFLOW STATUS
   ========================================= */

const STATUS_KEY =
  "SheetNavigator.Status";


const STATUS_LABELS = {
  none: "",
  prepared: "Prepared",
  openItems: "Open Items",
  pointsPosted: "Points Posted",
  pointsCleared: "Points Cleared",
  reviewed: "Reviewed",
  signerReview: "Signer Review"
};


/* =========================================
   TAB COLOR PALETTES
   ========================================= */

const THEME_COLORS = [
  "#FFFFFF",
  "#000000",
  "#E7E6E1",
  "#203864",
  "#4472C4",
  "#A64D45",
  "#8DAE52",
  "#6B5B95",
  "#4F9DB8",
  "#E69138",

  "#F2F2F2",
  "#7F7F7F",
  "#D9D2B6",
  "#BDD7EE",
  "#DDEBF7",
  "#E6B8AF",
  "#E2EFD9",
  "#D9D2E9",
  "#DDEBF7",
  "#FCE4D6",

  "#D9D9D9",
  "#595959",
  "#C4B78B",
  "#8EA9DB",
  "#B4C6E7",
  "#D5A6A1",
  "#C6E0B4",
  "#B4A7D6",
  "#A9D8E8",
  "#F9CB9C",

  "#BFBFBF",
  "#3F3F3F",
  "#948A54",
  "#5B9BD5",
  "#8FAADC",
  "#C98982",
  "#A9D18E",
  "#9E8BBE",
  "#79B4CB",
  "#F4B183",

  "#A6A6A6",
  "#262626",
  "#4C4524",
  "#244062",
  "#4472C4",
  "#943634",
  "#6B8E23",
  "#60497A",
  "#3F8496",
  "#D2691E"
];


const STANDARD_COLORS = [
  "#C00000",
  "#FF3300",
  "#FFC000",
  "#FFFF00",
  "#70AD47",
  "#44B3E1",
  "#2E75B6",
  "#002060",
  "#7030A0",
  "#A5A5A5"
];


/* =========================================
   STARTUP
   ========================================= */

Office.onReady((info) => {

  if (
    info.host !==
    Office.HostType.Excel
  ) {

    setStatus(
      "Open this add-in in Excel.",
      true
    );

    return;
  }


  document
    .getElementById("refresh")
    .addEventListener(
      "click",
      loadSheets
    );


  document
    .getElementById("search")
    .addEventListener(
      "input",
      renderSheets
    );


  document
    .getElementById("showHidden")
    .addEventListener(
      "change",
      renderSheets
    );


  document
    .getElementById("statusFilter")
    .addEventListener(
      "change",
      renderSheets
    );


  document
    .getElementById("collapseToggle")
    .addEventListener(
      "click",
      toggleNavigator
    );


  document.addEventListener(
    "click",
    handleOutsideClick
  );


  document.addEventListener(
    "keydown",
    handleKeyboard
  );


  setupMenuActions();

  setupColorPalette();

  setupDialog();

  loadSheets();

});


/* =========================================
   LOAD WORKSHEETS
   ========================================= */

async function loadSheets() {

  setStatus("Loading...");


  try {

    await Excel.run(
      async (context) => {

        const sheets =
          context.workbook.worksheets;


        const activeSheet =
          sheets.getActiveWorksheet();


        sheets.load(
          "items/id,items/name,items/visibility," +
          "items/tabColor,items/position"
        );


        activeSheet.load("id");


        await context.sync();


        const loadedSheets = [];


        for (
          const sheet of sheets.items
        ) {

          const properties =
            sheet.customProperties;


          properties.load(
            "items/key,items/value"
          );


          loadedSheets.push({
            sheet,
            properties
          });

        }


        await context.sync();


        sheetCache =
          loadedSheets.map(
            ({
              sheet,
              properties
            }) => {

              const metadata = {};


              for (
                const property of
                properties.items
              ) {

                metadata[property.key] =
                  property.value;

              }


              const storedStatus =
                metadata[STATUS_KEY] ||
                "none";


              const validStatus =
                Object.prototype
                  .hasOwnProperty
                  .call(
                    STATUS_LABELS,
                    storedStatus
                  )

                  ? storedStatus

                  : "none";


              return {

                id:
                  sheet.id,

                name:
                  sheet.name,

                visibility:
                  sheet.visibility,

                tabColor:
                  sheet.tabColor || "",

                position:
                  sheet.position,

                workflowStatus:
                  validStatus

              };

            }
          );


        sheetCache.sort(
          (left, right) =>
            left.position -
            right.position
        );


        activeSheetId =
          activeSheet.id;


        if (!handlersRegistered) {

          sheets.onActivated.add(
            onWorksheetActivated
          );


          sheets.onAdded.add(
            loadSheets
          );


          sheets.onDeleted.add(
            loadSheets
          );


          handlersRegistered =
            true;


          await context.sync();

        }

      }
    );


    setStatus("Ready");


    renderSheets();


    scrollActiveIntoView();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   ACTIVATE WORKSHEET
   ========================================= */

function onWorksheetActivated(
  event
) {

  activeSheetId =
    event.worksheetId;


  renderSheets();


  scrollActiveIntoView();

}


async function activateSheet(
  sheetId
) {

  try {

    await Excel.run(
      async (context) => {

        const worksheet =
          context.workbook
            .worksheets
            .getItem(sheetId);


        worksheet.activate();


        await context.sync();

      }
    );


    activeSheetId =
      sheetId;


    renderSheets();


    scrollActiveIntoView();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   RENDER WORKSHEETS
   ========================================= */

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


  const statusFilter =
    document
      .getElementById("statusFilter")
      .value;


  const filteredSheets =
    sheetCache.filter(
      (sheet) => {

        const matchesSearch =
          sheet.name
            .toLowerCase()
            .includes(searchText);


        const matchesVisibility =
          showHidden ||
          sheet.visibility ===
            "Visible";


        const matchesStatus =
          statusFilter === "all"

          ||

          sheet.workflowStatus ===
            statusFilter;


        return (
          matchesSearch &&
          matchesVisibility &&
          matchesStatus
        );

      }
    );


  const list =
    document.getElementById(
      "sheetList"
    );


  list.replaceChildren();


  for (
    const sheet of filteredSheets
  ) {

    const listItem =
      document.createElement(
        "li"
      );


    const row =
      document.createElement(
        "div"
      );


    row.className =
      "sheet-row";


    row.dataset.sheetId =
      sheet.id;


    row.draggable =
      true;


    /* WORKFLOW ROW COLOR */

    if (
      sheet.workflowStatus !==
      "none"
    ) {

      row.classList.add(
        `status-${sheet.workflowStatus}`
      );

    }


    /* ACTIVE WORKSHEET */

    if (
      sheet.id ===
      activeSheetId
    ) {

      row.classList.add(
        "active"
      );

    }


    /* HIDDEN WORKSHEET */

    if (
      sheet.visibility !==
      "Visible"
    ) {

      row.classList.add(
        "hidden-sheet"
      );

    }


    /* RIGHT-CLICK MENU */

    row.addEventListener(
      "contextmenu",
      (event) => {

        event.preventDefault();

        event.stopPropagation();


        openSheetMenuAtPoint(
          event.clientX,
          event.clientY,
          sheet.id
        );

      }
    );


    /* DRAG HANDLE */

    const dragHandle =
      document.createElement(
        "span"
      );


    dragHandle.className =
      "drag-handle";


    dragHandle.textContent =
      "⋮⋮";


    dragHandle.title =
      "Drag to reorder";


    /* WORKSHEET BUTTON */

    const sheetButton =
      document.createElement(
        "button"
      );


    sheetButton.className =
      "sheet-button";


    sheetButton.title =
      buildSheetTooltip(
        sheet
      );


    sheetButton.addEventListener(
      "click",
      () =>
        activateSheet(
          sheet.id
        )
    );


    sheetButton.addEventListener(
      "dblclick",
      async (event) => {

        event.stopPropagation();


        await renameSheet(
          sheet
        );

      }
    );


    /* ACTUAL EXCEL TAB COLOR */

    const tabColor =
      document.createElement(
        "span"
      );


    tabColor.className =
      "tab-color";


    tabColor.style.background =
      sheet.tabColor ||
      "#c8c6c4";


    /* WORKSHEET NAME */

    const sheetName =
      document.createElement(
        "span"
      );


    sheetName.className =
      "sheet-name";


    sheetName.textContent =
      sheet.name;


    sheetButton.append(
      tabColor,
      sheetName
    );


    /* HIDDEN INDICATOR */

    if (
      sheet.visibility !==
      "Visible"
    ) {

      const hiddenLabel =
        document.createElement(
          "span"
        );


      hiddenLabel.className =
        "hidden-label";


      hiddenLabel.textContent =
        "H";


      sheetButton.append(
        hiddenLabel
      );

    }


    row.append(
      dragHandle,
      sheetButton
    );


    /* DRAG EVENTS */

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


    listItem.append(
      row
    );


    list.append(
      listItem
    );

  }


  document
    .getElementById("count")
    .textContent =
      `${filteredSheets.length}/` +
      `${sheetCache.length}`;


  document
    .getElementById("empty")
    .hidden =
      filteredSheets.length !== 0;

}


/* =========================================
   WORKSHEET TOOLTIP
   ========================================= */

function buildSheetTooltip(
  sheet
) {

  const status =
    STATUS_LABELS[
      sheet.workflowStatus
    ] || "No Status";


  return (
    `${sheet.name}\n` +
    `Workflow: ${status}`
  );

}


/* =========================================
   WORKFLOW STATUS
   ========================================= */

async function setWorkflowStatus(
  sheet,
  newStatus
) {

  if (
    !Object.prototype
      .hasOwnProperty
      .call(
        STATUS_LABELS,
        newStatus
      )
  ) {

    return;

  }


  try {

    await Excel.run(
      async (context) => {

        const worksheet =
          context.workbook
            .worksheets
            .getItem(
              sheet.id
            );


        worksheet
          .customProperties
          .add(
            STATUS_KEY,
            newStatus
          );


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   COLLAPSE / EXPAND
   ========================================= */

function toggleNavigator() {

  navigatorCollapsed =
    !navigatorCollapsed;


  const app =
    document.getElementById(
      "app"
    );


  const toggle =
    document.getElementById(
      "collapseToggle"
    );


  app.classList.toggle(
    "collapsed",
    navigatorCollapsed
  );


  toggle.textContent =
    navigatorCollapsed
      ? "›"
      : "‹";


  toggle.title =
    navigatorCollapsed

      ? "Expand Sheet Navigator"

      : "Collapse Sheet Navigator";


  toggle.setAttribute(
    "aria-label",
    toggle.title
  );

}


/* =========================================
   DRAG / DROP
   ========================================= */

function handleDragStart(
  event
) {

  draggedSheetId =
    event.currentTarget
      .dataset.sheetId;


  event.currentTarget
    .classList.add(
      "dragging"
    );


  event.dataTransfer
    .effectAllowed =
      "move";

}


function handleDragOver(
  event
) {

  event.preventDefault();


  event.dataTransfer
    .dropEffect =
      "move";


  event.currentTarget
    .classList.add(
      "drag-over"
    );

}


function handleDragLeave(
  event
) {

  event.currentTarget
    .classList.remove(
      "drag-over"
    );

}


async function handleDrop(
  event
) {

  event.preventDefault();


  const targetSheetId =
    event.currentTarget
      .dataset.sheetId;


  event.currentTarget
    .classList.remove(
      "drag-over"
    );


  if (
    !draggedSheetId ||
    draggedSheetId ===
      targetSheetId
  ) {

    return;

  }


  await moveSheet(
    draggedSheetId,
    targetSheetId
  );

}


function handleDragEnd(
  event
) {

  event.currentTarget
    .classList.remove(
      "dragging"
    );


  document
    .querySelectorAll(
      ".drag-over"
    )
    .forEach(
      (element) => {

        element
          .classList
          .remove(
            "drag-over"
          );

      }
    );


  draggedSheetId =
    null;

}


async function moveSheet(
  sourceId,
  targetId
) {

  try {

    await Excel.run(
      async (context) => {

        const sheets =
          context.workbook
            .worksheets;


        const sourceSheet =
          sheets.getItem(
            sourceId
          );


        const targetSheet =
          sheets.getItem(
            targetId
          );


        targetSheet.load(
          "position"
        );


        await context.sync();


        sourceSheet.position =
          targetSheet.position;


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   WORKSHEET MENU
   ========================================= */

function prepareSheetMenu(
  sheetId
) {

  selectedSheetId =
    sheetId;


  const menu =
    document.getElementById(
      "sheetMenu"
    );


  const sheet =
    sheetCache.find(
      (item) =>
        item.id === sheetId
    );


  if (!sheet) {

    return null;

  }


  const visibilityButton =
    menu.querySelector(
      '[data-action="visibility"]'
    );


  visibilityButton.textContent =
    sheet.visibility ===
    "Visible"

      ? "Hide Sheet"

      : "Unhide Sheet";


  updateWorkflowMenuSelection(
    sheet.workflowStatus
  );


  document
    .getElementById(
      "colorMenu"
    )
    .hidden =
      true;


  return menu;

}


/* =========================================
   SHOW CURRENT STATUS IN MENU
   ========================================= */

function updateWorkflowMenuSelection(
  workflowStatus
) {

  document
    .querySelectorAll(
      '#sheetMenu ' +
      '[data-action="workflow"]'
    )
    .forEach(
      (button) => {

        const isCurrent =
          button.dataset.status ===
          workflowStatus;


        button.classList.toggle(
          "current-workflow-status",
          isCurrent
        );


        button.setAttribute(
          "aria-checked",
          String(isCurrent)
        );

      }
    );

}


/* =========================================
   OPEN WORKSHEET MENU
   ========================================= */

function openSheetMenuAtPoint(
  x,
  y,
  sheetId
) {

  const menu =
    prepareSheetMenu(
      sheetId
    );


  if (!menu) {

    return;

  }


  menu.style.left =
    "5px";


  menu.style.top =
    "5px";


  menu.hidden =
    false;


  const width =
    menu.offsetWidth;


  const height =
    menu.offsetHeight;


  const left =
    Math.max(
      5,
      Math.min(
        x,
        window.innerWidth -
          width -
          5
      )
    );


  const top =
    Math.max(
      5,
      Math.min(
        y,
        window.innerHeight -
          height -
          5
      )
    );


  menu.style.left =
    `${left}px`;


  menu.style.top =
    `${top}px`;

}


/* =========================================
   WORKSHEET MENU ACTIONS
   ========================================= */

function setupMenuActions() {

  document
    .querySelectorAll(
      "#sheetMenu button"
    )
    .forEach(
      (button) => {

        button.addEventListener(
          "click",
          async (event) => {

            event.stopPropagation();


            const sheet =
              sheetCache.find(
                (item) =>
                  item.id ===
                  selectedSheetId
              );


            if (!sheet) {

              return;

            }


            const action =
              button.dataset.action;


            /* WORKFLOW STATUS */

            if (
              action === "workflow"
            ) {

              closeMenus();


              await setWorkflowStatus(
                sheet,
                button.dataset.status
              );


              return;

            }


            /* TAB COLOR */

            if (
              action === "color"
            ) {

              openColorMenu();


              return;

            }


            closeMenus();


            switch (action) {

              case "activate":

                await activateSheet(
                  sheet.id
                );

                break;


              case "insert":

                await insertSheetAfter(
                  sheet
                );

                break;


              case "duplicate":

                await duplicateSheet(
                  sheet
                );

                break;


              case "delete":

                await deleteSheet(
                  sheet
                );

                break;


              case "rename":

                await renameSheet(
                  sheet
                );

                break;


              case "visibility":

                await toggleVisibility(
                  sheet
                );

                break;

            }

          }
        );

      }
    );

}


/* =========================================
   INSERT SHEET
   ========================================= */

async function insertSheetAfter(
  sheet
) {

  const name =
    await showTextDialog(
      "Insert Sheet",
      "Enter a name for the new worksheet:",
      "New Sheet"
    );


  if (name === null) {

    return;

  }


  try {

    await Excel.run(
      async (context) => {

        const sheets =
          context.workbook
            .worksheets;


        const relativeSheet =
          sheets.getItem(
            sheet.id
          );


        relativeSheet.load(
          "position"
        );


        await context.sync();


        const cleanedName =
          name.trim();


        const addedSheet =
          cleanedName

            ? sheets.add(
                cleanedName
              )

            : sheets.add();


        addedSheet.position =
          relativeSheet.position + 1;


        addedSheet.activate();


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   DUPLICATE SHEET
   ========================================= */

async function duplicateSheet(
  sheet
) {

  try {

    await Excel.run(
      async (context) => {

        const worksheet =
          context.workbook
            .worksheets
            .getItem(
              sheet.id
            );


        const copy =
          worksheet.copy(
            Excel
              .WorksheetPositionType
              .after,
            worksheet
          );


        copy.activate();


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   DELETE SHEET
   ========================================= */

async function deleteSheet(
  sheet
) {

  if (
    sheetCache.length <= 1
  ) {

    await showMessageDialog(
      "Cannot delete sheet",
      "Excel must contain at least one worksheet."
    );


    return;

  }


  const confirmed =
    await showConfirmDialog(
      "Delete Sheet",
      `Delete worksheet "${sheet.name}"?`,
      "Delete",
      true
    );


  if (!confirmed) {

    return;

  }


  try {

    await Excel.run(
      async (context) => {

        context.workbook
          .worksheets
          .getItem(
            sheet.id
          )
          .delete();


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   RENAME SHEET
   ========================================= */

async function renameSheet(
  sheet
) {

  const name =
    await showTextDialog(
      "Rename Sheet",
      "Worksheet name:",
      sheet.name
    );


  if (
    name === null ||
    !name.trim()
  ) {

    return;

  }


  try {

    await Excel.run(
      async (context) => {

        context.workbook
          .worksheets
          .getItem(
            sheet.id
          )
          .name =
            name.trim();


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   HIDE / UNHIDE
   ========================================= */

async function toggleVisibility(
  sheet
) {

  const visibleCount =
    sheetCache.filter(
      (item) =>
        item.visibility ===
        "Visible"
    ).length;


  if (
    sheet.visibility ===
      "Visible" &&
    visibleCount <= 1
  ) {

    await showMessageDialog(
      "Cannot hide sheet",
      "Excel must have at least one visible worksheet."
    );


    return;

  }


  try {

    await Excel.run(
      async (context) => {

        const worksheet =
          context.workbook
            .worksheets
            .getItem(
              sheet.id
            );


        worksheet.visibility =
          sheet.visibility ===
          "Visible"

            ? "Hidden"

            : "Visible";


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   TAB COLOR PALETTE
   ========================================= */

function setupColorPalette() {

  renderPalette(
    "themeColors",
    THEME_COLORS
  );


  renderPalette(
    "standardColors",
    STANDARD_COLORS
  );


  renderRecentColors();


  document
    .getElementById(
      "noColor"
    )
    .addEventListener(
      "click",
      async (event) => {

        event.stopPropagation();


        closeMenus();


        await setTabColor(
          selectedSheetId,
          ""
        );

      }
    );


  const picker =
    document.getElementById(
      "moreColorsInput"
    );


  document
    .getElementById(
      "moreColors"
    )
    .addEventListener(
      "click",
      (event) => {

        event.stopPropagation();


        picker.click();

      }
    );


  picker.addEventListener(
    "input",
    async () => {

      const color =
        picker.value;


      rememberRecentColor(
        color
      );


      closeMenus();


      await setTabColor(
        selectedSheetId,
        color
      );

    }
  );

}


/* =========================================
   RENDER COLOR PALETTE
   ========================================= */

function renderPalette(
  elementId,
  colors
) {

  const host =
    document.getElementById(
      elementId
    );


  host.replaceChildren();


  for (
    const color of colors
  ) {

    const colorButton =
      document.createElement(
        "button"
      );


    colorButton.className =
      "palette-swatch";


    colorButton.style.background =
      color;


    colorButton.title =
      color;


    colorButton.addEventListener(
      "click",
      async (event) => {

        event.stopPropagation();


        rememberRecentColor(
          color
        );


        closeMenus();


        await setTabColor(
          selectedSheetId,
          color
        );

      }
    );


    host.append(
      colorButton
    );

  }

}


/* =========================================
   RECENT COLORS
   ========================================= */

function getRecentColors() {

  try {

    return JSON.parse(
      localStorage.getItem(
        "SheetNavigator.RecentColors"
      ) || "[]"
    );


  } catch {

    return [];

  }

}


function rememberRecentColor(
  color
) {

  if (!color) {

    return;

  }


  let recent =
    getRecentColors()
      .filter(
        (item) =>
          item.toLowerCase() !==
          color.toLowerCase()
      );


  recent.unshift(
    color
  );


  recent =
    recent.slice(
      0,
      8
    );


  localStorage.setItem(
    "SheetNavigator.RecentColors",
    JSON.stringify(
      recent
    )
  );


  renderRecentColors();

}


function renderRecentColors() {

  renderPalette(
    "recentColors",
    getRecentColors()
  );

}


/* =========================================
   OPEN COLOR PALETTE
   ========================================= */

function openColorMenu() {

  const parentMenu =
    document.getElementById(
      "sheetMenu"
    );


  const colorMenu =
    document.getElementById(
      "colorMenu"
    );


  const parentRect =
    parentMenu
      .getBoundingClientRect();


  colorMenu.hidden =
    false;


  const width =
    colorMenu.offsetWidth;


  const height =
    colorMenu.offsetHeight;


  let left =
    parentRect.left -
    width -
    7;


  if (left < 5) {

    left =
      Math.min(
        window.innerWidth -
          width -
          5,

        parentRect.right +
          7
      );

  }


  const top =
    Math.max(
      5,
      Math.min(
        parentRect.top,
        window.innerHeight -
          height -
          5
      )
    );


  colorMenu.style.left =
    `${Math.max(5, left)}px`;


  colorMenu.style.top =
    `${top}px`;

}


/* =========================================
   APPLY TAB COLOR
   ========================================= */

async function setTabColor(
  sheetId,
  color
) {

  try {

    await Excel.run(
      async (context) => {

        context.workbook
          .worksheets
          .getItem(sheetId)
          .tabColor =
            color;


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   CUSTOM DIALOG SETUP
   ========================================= */

function setupDialog() {

  document
    .getElementById(
      "dialogCancel"
    )
    .addEventListener(
      "click",
      () => {

        resolveDialog(
          null
        );

      }
    );


  document
    .getElementById(
      "dialogConfirm"
    )
    .addEventListener(
      "click",
      () => {

        const input =
          document.getElementById(
            "dialogInput"
          );


        resolveDialog(
          input.hidden
            ? true
            : input.value
        );

      }
    );


  document
    .getElementById(
      "dialogOverlay"
    )
    .addEventListener(
      "click",
      (event) => {

        if (
          event.target.id ===
          "dialogOverlay"
        ) {

          resolveDialog(
            null
          );

        }

      }
    );


  document
    .getElementById(
      "dialogInput"
    )
    .addEventListener(
      "keydown",
      (event) => {

        if (
          event.key === "Enter"
        ) {

          document
            .getElementById(
              "dialogConfirm"
            )
            .click();

        }


        if (
          event.key === "Escape"
        ) {

          resolveDialog(
            null
          );

        }

      }
    );

}


/* =========================================
   OPEN CUSTOM DIALOG
   ========================================= */

function openDialog(
  title,
  message,
  {
    input = false,
    value = "",
    confirmText = "OK",
    danger = false
  } = {}
) {

  if (dialogResolver) {

    resolveDialog(
      null
    );

  }


  document
    .getElementById(
      "dialogTitle"
    )
    .textContent =
      title;


  document
    .getElementById(
      "dialogMessage"
    )
    .textContent =
      message || "";


  const inputElement =
    document.getElementById(
      "dialogInput"
    );


  const confirmButton =
    document.getElementById(
      "dialogConfirm"
    );


  inputElement.hidden =
    !input;


  inputElement.value =
    value;


  confirmButton.textContent =
    confirmText;


  confirmButton.classList.toggle(
    "danger-button",
    danger
  );


  confirmButton.classList.toggle(
    "primary-button",
    !danger
  );


  document
    .getElementById(
      "dialogOverlay"
    )
    .hidden =
      false;


  if (input) {

    setTimeout(
      () => {

        inputElement.focus();

        inputElement.select();

      },
      0
    );

  }


  return new Promise(
    (resolve) => {

      dialogResolver =
        resolve;

    }
  );

}


/* =========================================
   RESOLVE DIALOG
   ========================================= */

function resolveDialog(
  value
) {

  document
    .getElementById(
      "dialogOverlay"
    )
    .hidden =
      true;


  if (dialogResolver) {

    const resolver =
      dialogResolver;


    dialogResolver =
      null;


    resolver(
      value
    );

  }

}


/* =========================================
   DIALOG HELPERS
   ========================================= */

function showTextDialog(
  title,
  message,
  value = ""
) {

  return openDialog(
    title,
    message,
    {
      input: true,
      value,
      confirmText: "OK"
    }
  );

}


function showConfirmDialog(
  title,
  message,
  confirmText = "OK",
  danger = false
) {

  return openDialog(
    title,
    message,
    {
      confirmText,
      danger
    }
  );

}


async function showMessageDialog(
  title,
  message
) {

  await openDialog(
    title,
    message,
    {
      confirmText: "OK"
    }
  );

}


/* =========================================
   MENU HELPERS
   ========================================= */

function handleOutsideClick(
  event
) {

  if (
    !event.target.closest(
      "#sheetMenu"
    )

    &&

    !event.target.closest(
      "#colorMenu"
    )
  ) {

    closeMenus();

  }

}


function handleKeyboard(
  event
) {

  if (
    event.key === "Escape"
  ) {

    closeMenus();

  }

}


function closeMenus() {

  document
    .getElementById(
      "sheetMenu"
    )
    .hidden =
      true;


  document
    .getElementById(
      "colorMenu"
    )
    .hidden =
      true;

}


/* =========================================
   AUTO-SCROLL ACTIVE WORKSHEET
   ========================================= */

function scrollActiveIntoView() {

  requestAnimationFrame(
    () => {

      const activeRow =
        document.querySelector(
          `.sheet-row[data-sheet-id="${activeSheetId}"]`
        );


      if (activeRow) {

        activeRow.scrollIntoView({
          block: "nearest",
          behavior: "smooth"
        });

      }

    }
  );

}


/* =========================================
   ERROR / STATUS
   ========================================= */

function showError(
  error
) {

  console.error(
    error
  );


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

  const statusElement =
    document.getElementById(
      "status"
    );


  statusElement.textContent =
    message;


  statusElement.style.color =
    isError
      ? "#ffd7d7"
      : "";

}