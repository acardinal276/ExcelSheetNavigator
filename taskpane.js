/* global Office, Excel */

let sheetCache = [];
let activeSheetId = null;
let selectedSheetId = null;
let draggedSheetId = null;
let handlersRegistered = false;
let dialogResolver = null;


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
  "#943634", "#6B8E23", "#60497A", "#3F8496", "#D2691E",

  "#7F7F7F", "#000000", "#1C1A0D", "#17365D", "#2F5597",
  "#632523", "#4F6228", "#403151", "#275D6B", "#984807"
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

  if (info.host !== Office.HostType.Excel) {
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


        const activeSheet =
          sheets.getActiveWorksheet();


        sheets.load(
          "items/id,items/name,items/visibility," +
          "items/tabColor,items/position"
        );


        activeSheet.load("id");


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


                preparedBy:
                  metadata[
                    "SheetNavigator.PreparedBy"
                  ] || "",


                preparedAt:
                  metadata[
                    "SheetNavigator.PreparedAt"
                  ] || "",


                reviewedBy:
                  metadata[
                    "SheetNavigator.ReviewedBy"
                  ] || "",


                reviewedAt:
                  metadata[
                    "SheetNavigator.ReviewedAt"
                  ] || "",


                signedBy:
                  metadata[
                    "SheetNavigator.SignedBy"
                  ] || "",


                signedAt:
                  metadata[
                    "SheetNavigator.SignedAt"
                  ] || ""

              };

            }
          );


        sheetCache.sort(
          (a, b) =>
            a.position - b.position
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


          handlersRegistered = true;


          await context.sync();

        }

      }
    );


    setStatus("Ready");


    renderSheets();


    scrollActiveIntoView();


  } catch (error) {

    console.error(error);


    setStatus(
      error.message ||
      "Unable to load worksheets.",
      true
    );

  }

}


/* =========================================
   WORKSHEET ACTIVATION
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
      (sheet) =>

        sheet.name
          .toLowerCase()
          .includes(query)

        &&

        (
          showHidden ||
          sheet.visibility === "Visible"
        )

        &&

        statusMatches(
          sheet,
          filter
        )

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


    if (
      sheet.id === activeSheetId
    ) {

      row.classList.add(
        "active"
      );

    }


    if (
      sheet.visibility !==
      "Visible"
    ) {

      row.classList.add(
        "hidden-sheet"
      );

    }


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


    /* SHEET BUTTON */

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


    /* TAB COLOR INDICATOR */

    const color =
      document.createElement(
        "span"
      );


    color.className =
      "tab-color";


    color.style.background =
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
      color,
      name
    );


    /* HIDDEN INDICATOR */

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


    /* EXISTING SIGN-OFF BUTTONS */

    const prepared =
      createSignoffButton(
        sheet,
        "prepared"
      );


    const reviewed =
      createSignoffButton(
        sheet,
        "reviewed"
      );


    const signed =
      createSignoffButton(
        sheet,
        "signed"
      );


    /* MORE BUTTON */

    const more =
      document.createElement(
        "button"
      );


    more.className =
      "more-button";


    more.textContent =
      "⋯";


    more.addEventListener(
      "click",
      (event) => {

        event.stopPropagation();


        openSheetMenuAtButton(
          event.currentTarget,
          sheet.id
        );

      }
    );


    row.append(
      drag,
      button,
      prepared,
      reviewed,
      signed,
      more
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
   EXISTING SIGN-OFFS
   ========================================= */

function createSignoffButton(
  sheet,
  type
) {

  const button =
    document.createElement(
      "button"
    );


  button.className =
    "signoff";


  const values = {

    prepared: {

      person:
        sheet.preparedBy,

      date:
        sheet.preparedAt,

      label:
        "Prepared"

    },


    reviewed: {

      person:
        sheet.reviewedBy,

      date:
        sheet.reviewedAt,

      label:
        "Reviewed"

    },


    signed: {

      person:
        sheet.signedBy,

      date:
        sheet.signedAt,

      label:
        "Signed"

    }

  };


  const data =
    values[type];


  if (data.person) {

    button.classList.add(
      "signed"
    );


    button.textContent =
      "✓";


    button.title =
      `${data.label} by ${data.person}` +
      `${data.date
        ? "\n" + data.date
        : ""}`;


  } else {

    button.textContent =
      "○";


    button.title =
      `Mark ${data.label}`;

  }


  button.addEventListener(
    "click",
    async (event) => {

      event.stopPropagation();


      await toggleSignoff(
        sheet,
        type
      );

    }
  );


  return button;

}


async function toggleSignoff(
  sheet,
  type
) {

  const mapping = {

    prepared: {

      by:
        "SheetNavigator.PreparedBy",

      at:
        "SheetNavigator.PreparedAt",

      current:
        sheet.preparedBy

    },


    reviewed: {

      by:
        "SheetNavigator.ReviewedBy",

      at:
        "SheetNavigator.ReviewedAt",

      current:
        sheet.reviewedBy

    },


    signed: {

      by:
        "SheetNavigator.SignedBy",

      at:
        "SheetNavigator.SignedAt",

      current:
        sheet.signedBy

    }

  };


  const config =
    mapping[type];


  try {

    if (config.current) {

      const remove =
        await showConfirmDialog(
          "Remove sign-off?",
          "Remove this sign-off?",
          "Remove",
          true
        );


      if (!remove) {
        return;
      }


      await Excel.run(
        async (context) => {

          const worksheet =
            context.workbook
              .worksheets
              .getItem(
                sheet.id
              );


          const properties =
            worksheet
              .customProperties;


          const personProperty =
            properties
              .getItemOrNullObject(
                config.by
              );


          const timeProperty =
            properties
              .getItemOrNullObject(
                config.at
              );


          personProperty.load(
            "isNullObject"
          );


          timeProperty.load(
            "isNullObject"
          );


          await context.sync();


          if (
            !personProperty
              .isNullObject
          ) {

            personProperty
              .delete();

          }


          if (
            !timeProperty
              .isNullObject
          ) {

            timeProperty
              .delete();

          }


          await context.sync();

        }
      );


    } else {

      const initials =
        await showTextDialog(
          "Add sign-off",
          "Enter your initials:",
          ""
        );


      if (!initials) {
        return;
      }


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
              config.by,
              initials
                .trim()
                .toUpperCase()
            );


          worksheet
            .customProperties
            .add(
              config.at,
              new Date()
                .toLocaleString()
            );


          await context.sync();

        }
      );

    }


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   STATUS FILTER
   ========================================= */

function statusMatches(
  sheet,
  filter
) {

  switch (filter) {

    case "notPrepared":

      return (
        !sheet.preparedBy
      );


    case "awaitingReview":

      return (
        !!sheet.preparedBy &&
        !sheet.reviewedBy
      );


    case "awaitingSigner":

      return (
        !!sheet.reviewedBy &&
        !sheet.signedBy
      );


    case "complete":

      return (
        !!sheet.preparedBy &&
        !!sheet.reviewedBy &&
        !!sheet.signedBy
      );


    default:

      return true;

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
      item =>
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
    sheet.visibility === "Visible"
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


/* MENU FROM MORE BUTTON */

function openSheetMenuAtButton(
  button,
  sheetId
) {

  const menu =
    prepareSheetMenu(
      sheetId
    );


  if (!menu) {
    return;
  }


  const rect =
    button
      .getBoundingClientRect();


  menu.style.left =
    "auto";


  menu.style.right =
    "5px";


  menu.style.top =
    `${rect.bottom + 2}px`;


  menu.hidden =
    false;

}


/* MENU FROM RIGHT CLICK */

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


  menu.style.right =
    "auto";


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
          width - 5
      )
    );


  const top =
    Math.max(
      5,
      Math.min(
        y,
        window.innerHeight -
          height - 5
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


            const action =
              button.dataset.action;


            const sheet =
              sheetCache.find(
                item =>
                  item.id ===
                  selectedSheetId
              );


            if (!sheet) {
              return;
            }


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

                await deleteSheetFromNavigator(
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


        const relative =
          sheets.getItem(
            sheet.id
          );


        relative.load(
          "position"
        );


        await context.sync();


        const cleanName =
          name.trim();


        const added =
          cleanName
            ? sheets.add(
                cleanName
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

async function deleteSheetFromNavigator(
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
      `Delete worksheet "${sheet.name}"? ` +
      "This action may not be reversible.",
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

  try {

    await Excel.run(
      async (context) => {

        const worksheet =
          context.workbook
            .worksheets
            .getItem(
              sheet.id
            );


        if (
          sheet.visibility ===
          "Visible"
        ) {

          const visibleCount =
            sheetCache.filter(
              item =>
                item.visibility ===
                  "Visible"
            ).length;


          if (
            visibleCount <= 1
          ) {

            await showMessageDialog(
              "Cannot hide sheet",
              "Excel must have at least one visible worksheet."
            );


            return;

          }


          worksheet.visibility =
            "Hidden";


        } else {

          worksheet.visibility =
            "Visible";

        }


        await context.sync();

      }
    );


    await loadSheets();


  } catch (error) {

    showError(error);

  }

}


/* =========================================
   COLOR PALETTE
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


  const input =
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


        input.click();

      }
    );


  input.addEventListener(
    "input",
    async () => {

      const color =
        input.value;


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


/* RENDER COLOR BUTTONS */

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
      button
    );

  }

}


/* RECENT COLORS */

function rememberRecentColor(
  color
) {

  if (!color) {
    return;
  }


  let recent =
    getRecentColors()
      .filter(
        item =>
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


function renderRecentColors() {

  renderPalette(
    "recentColors",
    getRecentColors()
  );

}


/* OPEN COLOR PALETTE */

function openColorMenu() {

  const sheetMenu =
    document.getElementById(
      "sheetMenu"
    );


  const colorMenu =
    document.getElementById(
      "colorMenu"
    );


  const rect =
    sheetMenu
      .getBoundingClientRect();


  colorMenu.hidden =
    false;


  const width =
    colorMenu.offsetWidth;


  const height =
    colorMenu.offsetHeight;


  let left =
    rect.left -
    width -
    7;


  /*
   * If there isn't room to the left,
   * open the palette to the right.
   */

  if (left < 5) {

    left =
      Math.min(
        window.innerWidth -
          width -
          5,
        rect.right + 7
      );

  }


  const top =
    Math.max(
      5,
      Math.min(
        rect.top,
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


/* APPLY EXCEL TAB COLOR */

async function setTabColor(
  id,
  color
) {

  try {

    await Excel.run(
      async (context) => {

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
   CUSTOM TASK-PANE DIALOGS
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


/* GENERIC DIALOG */

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
    (resolve) => {

      dialogResolver =
        resolve;

    }
  );

}


/* RESOLVE DIALOG */

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


/* TEXT ENTRY DIALOG */

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
      confirmText:
        "OK"
    }
  );

}


/* CONFIRMATION DIALOG */

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


/* MESSAGE DIALOG */

async function showMessageDialog(
  title,
  message
) {

  await openDialog(
    title,
    message,
    {
      confirmText:
        "OK"
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
    ) &&
    !event.target.closest(
      "#colorMenu"
    ) &&
    !event.target.closest(
      ".more-button"
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


/* =========================================
   KEEP ACTIVE SHEET VISIBLE
   ========================================= */

function scrollActiveIntoView() {

  requestAnimationFrame(
    () => {

      const row =
        document.querySelector(
          `.sheet-row[data-sheet-id="${activeSheetId}"]`
        );


      if (row) {

        row.scrollIntoView({
          block:
            "nearest",

          behavior:
            "smooth"
        });

      }

    }
  );

}


/* =========================================
   ERRORS / STATUS
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