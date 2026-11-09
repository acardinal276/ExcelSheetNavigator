/* global Office, Excel */

let sheetCache = [];
let activeSheetId = null;
let selectedSheetId = null;
let draggedSheetId = null;
let handlersRegistered = false;
let dialogResolver = null;


/* =========================================
   WORKFLOW STATUS
   ========================================= */

const STATUS_KEY =
  "SheetNavigator.Status";


const STATUS_LABELS = {
  none: "",
  prepared: "Prepared",
  reviewed: "Reviewed",
  openItems: "Open Items",
  signerReview: "Signer Review"
};


/* =========================================
   COLOR PALETTES
   ========================================= */

const THEME_COLORS = [
  "#FFFFFF", "#000000", "#E7E6E1", "#203864", "#4472C4",
  "#A64D45", "#8DAE52", "#6B5B95", "#4F9DB8", "#E69138",

  "#F2F2F2", "#7F7F7F", "#D9D2B6", "#BDD7EE", "#DDEBF7",
  "#E6B8AF", "#E2EFD9", "#D9D2E9", "#DDEBF7", "#FCE4D6",

  "#D9D9D9", "#595959", "#C4B78B", "#8EA9DB", "#B4C6E7",
  "#D5A6A1", "#C6E0B4", "#B4A7D6", "#A9D8E8", "#F9CB9C",

  "#BFBFBF", "#3F3F3F", "#948A54", "#5B9BD5", "#8FAADC",
  "#C98982", "#A9D18E", "#9E8BBE", "#79B4CB", "#F4B183",

  "#A6A6A6", "#262626", "#4C4524", "#244062", "#4472C4",
  "#943634", "#6B8E23", "#60497A", "#3F8496", "#D2691E"
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


  document.addEventListener(
    "click",
    handleOutsideClick
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


        const active =
          sheets.getActiveWorksheet();


        sheets.load(
          "items/id,items/name,items/visibility," +
          "items/tabColor,items/position"
        );


        active.load("id");


        await context.sync();


        const loaded = [];


        for (
          const sheet of sheets.items
        ) {

          const props =
            sheet.customProperties;


          props.load(
            "items/key,items/value"
          );


          loaded.push({
            sheet,
            props
          });

        }


        await context.sync();


        sheetCache =
          loaded.map(
            ({ sheet, props }) => {

              const metadata = {};


              for (
                const prop of props.items
              ) {

                metadata[prop.key] =
                  prop.value;

              }


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
                  metadata[
                    STATUS_KEY
                  ] || "none"

              };

            }
          );


        sheetCache.sort(
          (a, b) =>
            a.position -
            b.position
        );


        activeSheetId =
          active.id;


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
  id
) {

  try {

    await Excel.run(
      async (context) => {

        context.workbook
          .worksheets
          .getItem(id)
          .activate();


        await context.sync();

      }
    );


    activeSheetId =
      id;


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

  const query =
    document
      .getElementById("search")
      .value
      .trim()
      .toLowerCase();


  const showHidden =
    document
      .getElementById("showHidden")
      .checked;


  const filter =
    document
      .getElementById("statusFilter")
      .value;


  const rows =
    sheetCache.filter(
      (sheet) => {

        const matchesSearch =
          sheet.name
            .toLowerCase()
            .includes(query);


        const matchesVisibility =
          showHidden ||
          sheet.visibility ===
            "Visible";


        const matchesStatus =
          filter === "all"

          ||

          (
            filter === "none" &&
            sheet.workflowStatus ===
              "none"
          )

          ||

          sheet.workflowStatus ===
            filter;


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
    const sheet of rows
  ) {

    const li =
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


    /* ACTIVE SHEET */

    if (
      sheet.id ===
      activeSheetId
    ) {

      row.classList.add(
        "active"
      );

    }


    /* HIDDEN SHEET */

    if (
      sheet.visibility !==
      "Visible"
    ) {

      row.classList.add(
        "hidden-sheet"
      );

    }


    /* RIGHT CLICK */

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

    const drag =
      document.createElement(
        "span"
      );


    drag.className =
      "drag-handle";


    drag.textContent =
      "⋮⋮";


    drag.title =
      "Drag to reorder";


    /* WORKSHEET BUTTON */

    const button =
      document.createElement(
        "button"
      );


    button.className =
      "sheet-button";


    button.title =
      sheet.name;


    button.addEventListener(
      "click",
      () =>
        activateSheet(
          sheet.id
        )
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

    const name =
      document.createElement(
        "span"
      );


    name.className =
      "sheet-name";


    name.textContent =
      sheet.name;


    button.append(
      tabColor,
      name
    );


    if (
      sheet.visibility !==
      "Visible"
    ) {

      const hidden =
        document.createElement(
          "span"
        );


      hidden.className =
        "hidden-label";


      hidden.textContent =
        "H";


      button.append(
        hidden
      );

    }


    /* STATUS LABEL */

    const workflow =
      document.createElement(
        "span"
      );


    workflow.className =
      "workflow-label";


    workflow.textContent =
      STATUS_LABELS[
        sheet.workflowStatus
      ] || "";


    row.append(
      drag,
      button,
      workflow
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


    li.append(row);


    list.append(li);

  }


  document
    .getElementById("count")
    .textContent =
      `${rows.length}/${sheetCache.length}`;


  document
    .getElementById("empty")
    .hidden =
      rows.length !== 0;

}


/* =========================================
   WORKFLOW STATUS
   ========================================= */

async function setWorkflowStatus(
  sheet,
  newStatus
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


  const target =
    event.currentTarget
      .dataset.sheetId;


  event.currentTarget
    .classList.remove(
      "drag-over"
    );


  if (
    !draggedSheetId ||
    draggedSheetId === target
  ) {

    return;

  }


  await moveSheet(
    draggedSheetId,
    target
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
      element =>
        element
          .classList
          .remove(
            "drag-over"
          )
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


        const source =
          sheets.getItem(
            sourceId
          );


        const target =
          sheets.getItem(
            targetId
          );


        target.load(
          "position"
        );


        await context.sync();


        source.position =
          target.position;


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
  id
) {

  selectedSheetId =
    id;


  const menu =
    document.getElementById(
      "sheetMenu"
    );


  const sheet =
    sheetCache.find(
      item =>
        item.id === id
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


  document
    .getElementById(
      "colorMenu"
    )
    .hidden =
      true;


  return menu;

}


function openSheetMenuAtPoint(
  x,
  y,
  id
) {

  const menu =
    prepareSheetMenu(id);


  if (!menu) {
    return;
  }


  menu.hidden =
    false;


  const width =
    menu.offsetWidth;


  const height =
    menu.offsetHeight;


  menu.style.left =
    `${
      Math.max(
        5,
        Math.min(
          x,
          window.innerWidth -
            width -
            5
        )
      )
    }px`;


  menu.style.top =
    `${
      Math.max(
        5,
        Math.min(
          y,
          window.innerHeight -
            height -
            5
        )
      )
    }px`;

}


/* =========================================
   MENU ACTIONS
   ========================================= */

function setupMenuActions() {

  document
    .querySelectorAll(
      "#sheetMenu button"
    )
    .forEach(
      button => {

        button.addEventListener(
          "click",
          async (event) => {

            event.stopPropagation();


            const sheet =
              sheetCache.find(
                item =>
                  item.id ===
                  selectedSheetId
              );


            if (!sheet) {
              return;
            }


            const action =
              button.dataset.action;


            /* WORKFLOW */

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
      async context => {

        const sheets =
          context.workbook
            .worksheets;


        const relative =
          sheets.getItem(
            sheet.id
          );


        relative.load(
          "position"
        );


        await context.sync();


        const cleaned =
          name.trim();


        const added =
          cleaned
            ? sheets.add(
                cleaned
              )
            : sheets.add();


        added.position =
          relative.position + 1;


        added.activate();


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
      async context => {

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
      async context => {

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
      async context => {

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

  try {

    await Excel.run(
      async context => {

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
      async event => {

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
      event => {

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


function renderPalette(
  id,
  colors
) {

  const host =
    document.getElementById(
      id
    );


  host.replaceChildren();


  for (
    const color of colors
  ) {

    const button =
      document.createElement(
        "button"
      );


    button.className =
      "palette-swatch";


    button.style.background =
      color;


    button.title =
      color;


    button.addEventListener(
      "click",
      async event => {

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
      button
    );

  }

}


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

  let recent =
    getRecentColors()
      .filter(
        item =>
          item !== color
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


function openColorMenu() {

  const parent =
    document.getElementById(
      "sheetMenu"
    );


  const menu =
    document.getElementById(
      "colorMenu"
    );


  const rect =
    parent
      .getBoundingClientRect();


  menu.hidden =
    false;


  const width =
    menu.offsetWidth;


  const height =
    menu.offsetHeight;


  let left =
    rect.left -
    width -
    7;


  if (left < 5) {

    left =
      rect.right +
      7;

  }


  menu.style.left =
    `${Math.max(5, left)}px`;


  menu.style.top =
    `${
      Math.max(
        5,
        Math.min(
          rect.top,
          window.innerHeight -
            height -
            5
        )
      )
    }px`;

}


async function setTabColor(
  id,
  color
) {

  try {

    await Excel.run(
      async context => {

        context.workbook
          .worksheets
          .getItem(id)
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
   CUSTOM DIALOG
   ========================================= */

function setupDialog() {

  document
    .getElementById(
      "dialogCancel"
    )
    .addEventListener(
      "click",
      () =>
        resolveDialog(
          null
        )
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

}


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
      message;


  const field =
    document.getElementById(
      "dialogInput"
    );


  const confirm =
    document.getElementById(
      "dialogConfirm"
    );


  field.hidden =
    !input;


  field.value =
    value;


  confirm.textContent =
    confirmText;


  confirm.classList.toggle(
    "danger-button",
    danger
  );


  confirm.classList.toggle(
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

        field.focus();

        field.select();

      },
      0
    );

  }


  return new Promise(
    resolve => {

      dialogResolver =
        resolve;

    }
  );

}


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


function showTextDialog(
  title,
  message,
  value
) {

  return openDialog(
    title,
    message,
    {
      input: true,
      value
    }
  );

}


function showConfirmDialog(
  title,
  message,
  confirmText,
  danger
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


/* =========================================
   HELPERS
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


function scrollActiveIntoView() {

  requestAnimationFrame(
    () => {

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

    }
  );

}


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

  const element =
    document.getElementById(
      "status"
    );


  element.textContent =
    message;


  element.style.color =
    isError
      ? "#ffd7d7"
      : "";

}