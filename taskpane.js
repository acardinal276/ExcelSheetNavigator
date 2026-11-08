/* global Office, Excel */

let sheetCache = [];
let activeSheetId = null;
let selectedSheetId = null;
let draggedSheetId = null;
let handlersRegistered = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Open this add-in in Excel.", true);
    return;
  }

  document.getElementById("refresh")
    .addEventListener("click", loadSheets);

  document.getElementById("search")
    .addEventListener("input", renderSheets);

  document.getElementById("showHidden")
    .addEventListener("change", renderSheets);

  document.getElementById("statusFilter")
    .addEventListener("change", renderSheets);

  document.addEventListener("click", handleOutsideClick);

  setupMenuActions();
  setupColorMenu();

  loadSheets();
});


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

      const loadedSheets = [];

      for (const sheet of sheets.items) {

        const props = sheet.customProperties;

        props.load("items/key,items/value");

        loadedSheets.push({
          sheet,
          props
        });
      }

      await context.sync();

      sheetCache = loadedSheets.map(({ sheet, props }) => {

        const metadata = {};

        for (const prop of props.items) {
          metadata[prop.key] = prop.value;
        }

        return {
          id: sheet.id,
          name: sheet.name,
          visibility: sheet.visibility,
          tabColor: sheet.tabColor || "",
          position: sheet.position,

          preparedBy:
            metadata["SheetNavigator.PreparedBy"] || "",

          preparedAt:
            metadata["SheetNavigator.PreparedAt"] || "",

          reviewedBy:
            metadata["SheetNavigator.ReviewedBy"] || "",

          reviewedAt:
            metadata["SheetNavigator.ReviewedAt"] || "",

          signedBy:
            metadata["SheetNavigator.SignedBy"] || "",

          signedAt:
            metadata["SheetNavigator.SignedAt"] || ""
        };
      });

      sheetCache.sort(
        (a, b) => a.position - b.position
      );

      activeSheetId = activeSheet.id;

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

    });

    setStatus("Ready");

    renderSheets();

    scrollActiveIntoView();

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

  scrollActiveIntoView();
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

    scrollActiveIntoView();

  } catch (error) {
    showError(error);
  }
}


function renderSheets() {

  const searchText =
    document.getElementById("search")
      .value
      .trim()
      .toLowerCase();

  const showHidden =
    document.getElementById("showHidden")
      .checked;

  const statusFilter =
    document.getElementById("statusFilter")
      .value;


  const filteredSheets =
    sheetCache.filter((sheet) => {

      const matchesSearch =
        sheet.name
          .toLowerCase()
          .includes(searchText);

      const matchesVisibility =
        showHidden ||
        sheet.visibility === "Visible";

      const matchesStatus =
        statusMatches(
          sheet,
          statusFilter
        );

      return (
        matchesSearch &&
        matchesVisibility &&
        matchesStatus
      );
    });


  const list =
    document.getElementById("sheetList");

  list.replaceChildren();


  for (const sheet of filteredSheets) {

    const li =
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


    /* DRAG HANDLE */

    const dragHandle =
      document.createElement("span");

    dragHandle.className =
      "drag-handle";

    dragHandle.textContent =
      "⋮⋮";

    dragHandle.title =
      "Drag to reorder";


    /* SHEET NAME BUTTON */

    const sheetButton =
      document.createElement("button");

    sheetButton.className =
      "sheet-button";

    sheetButton.title =
      sheet.name;

    sheetButton.addEventListener(
      "click",
      () => activateSheet(sheet.id)
    );


    const tabColor =
      document.createElement("span");

    tabColor.className =
      "tab-color";

    tabColor.style.background =
      sheet.tabColor || "#c8c6c4";


    const sheetName =
      document.createElement("span");

    sheetName.className =
      "sheet-name";

    sheetName.textContent =
      sheet.name;


    sheetButton.append(
      tabColor,
      sheetName
    );


    if (sheet.visibility !== "Visible") {

      const hidden =
        document.createElement("span");

      hidden.className =
        "hidden-label";

      hidden.textContent =
        "H";

      sheetButton.append(hidden);
    }


    /* SIGN-OFFS */

    const prepare =
      createSignoffButton(
        sheet,
        "prepared"
      );

    const review =
      createSignoffButton(
        sheet,
        "reviewed"
      );

    const signer =
      createSignoffButton(
        sheet,
        "signed"
      );


    /* MORE MENU */

    const more =
      document.createElement("button");

    more.className =
      "more-button";

    more.textContent =
      "⋯";

    more.title =
      "Worksheet actions";

    more.addEventListener(
      "click",
      (event) => {

        event.stopPropagation();

        openSheetMenu(
          event.currentTarget,
          sheet.id
        );
      }
    );


    row.append(
      dragHandle,
      sheetButton,
      prepare,
      review,
      signer,
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


  document.getElementById("count")
    .textContent =
      `${filteredSheets.length}/${sheetCache.length}`;


  document.getElementById("empty")
    .hidden =
      filteredSheets.length !== 0;
}


function createSignoffButton(
  sheet,
  type
) {

  const button =
    document.createElement("button");

  button.className =
    "signoff";


  const values = {

    prepared: {
      person: sheet.preparedBy,
      date: sheet.preparedAt,
      label: "Prepared"
    },

    reviewed: {
      person: sheet.reviewedBy,
      date: sheet.reviewedAt,
      label: "Reviewed"
    },

    signed: {
      person: sheet.signedBy,
      date: sheet.signedAt,
      label: "Signed"
    }

  };


  const data =
    values[type];


  if (data.person) {

    button.classList.add("signed");

    button.textContent =
      "✓";

    button.title =
      `${data.label} by ${data.person}` +
      `${data.date ? "\n" + data.date : ""}`;

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
      by: "SheetNavigator.PreparedBy",
      at: "SheetNavigator.PreparedAt",
      current: sheet.preparedBy
    },

    reviewed: {
      by: "SheetNavigator.ReviewedBy",
      at: "SheetNavigator.ReviewedAt",
      current: sheet.reviewedBy
    },

    signed: {
      by: "SheetNavigator.SignedBy",
      at: "SheetNavigator.SignedAt",
      current: sheet.signedBy
    }

  };


  const config =
    mapping[type];


  try {

    if (config.current) {

      const remove =
        window.confirm(
          "Remove this sign-off?"
        );

      if (!remove) {
        return;
      }

      await Excel.run(
        async (context) => {

          const worksheet =
            context.workbook
              .worksheets
              .getItem(sheet.id);

          const properties =
            worksheet.customProperties;

          const byProperty =
            properties
              .getItemOrNullObject(
                config.by
              );

          const atProperty =
            properties
              .getItemOrNullObject(
                config.at
              );

          byProperty.load(
            "isNullObject"
          );

          atProperty.load(
            "isNullObject"
          );

          await context.sync();

          if (!byProperty.isNullObject) {
            byProperty.delete();
          }

          if (!atProperty.isNullObject) {
            atProperty.delete();
          }

          await context.sync();
        }
      );

    } else {

      let initials =
        window.prompt(
          "Enter your initials:"
        );

      if (!initials) {
        return;
      }

      initials =
        initials
          .trim()
          .toUpperCase();

      if (!initials) {
        return;
      }

      const timestamp =
        new Date()
          .toLocaleString();


      await Excel.run(
        async (context) => {

          const worksheet =
            context.workbook
              .worksheets
              .getItem(sheet.id);

          worksheet
            .customProperties
            .add(
              config.by,
              initials
            );

          worksheet
            .customProperties
            .add(
              config.at,
              timestamp
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


/* FILTERING */

function statusMatches(
  sheet,
  filter
) {

  switch (filter) {

    case "notPrepared":
      return !sheet.preparedBy;

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


/* DRAG / DROP */

function handleDragStart(event) {

  draggedSheetId =
    event.currentTarget
      .dataset.sheetId;

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

  const targetSheetId =
    event.currentTarget
      .dataset.sheetId;

  event.currentTarget
    .classList.remove("drag-over");


  if (
    !draggedSheetId ||
    draggedSheetId === targetSheetId
  ) {
    return;
  }


  await moveSheet(
    draggedSheetId,
    targetSheetId
  );
}


function handleDragEnd(event) {

  event.currentTarget
    .classList.remove("dragging");

  document
    .querySelectorAll(".drag-over")
    .forEach(
      element =>
        element.classList.remove(
          "drag-over"
        )
    );

  draggedSheetId = null;
}


async function moveSheet(
  sourceId,
  targetId
) {

  try {

    await Excel.run(
      async (context) => {

        const sheets =
          context.workbook.worksheets;

        const source =
          sheets.getItem(sourceId);

        const target =
          sheets.getItem(targetId);

        target.load("position");

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


/* SHEET ACTION MENU */

function openSheetMenu(
  button,
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


  const visibilityButton =
    menu.querySelector(
      '[data-action="visibility"]'
    );

  visibilityButton.textContent =
    sheet.visibility === "Visible"
      ? "Hide"
      : "Unhide";


  const rect =
    button.getBoundingClientRect();

  menu.style.top =
    `${rect.bottom + 2}px`;

  menu.style.right =
    "5px";

  menu.hidden = false;


  document.getElementById(
    "colorMenu"
  ).hidden = true;
}


function setupMenuActions() {

  document
    .querySelectorAll(
      "#sheetMenu button"
    )
    .forEach((button) => {

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


          if (action === "activate") {

            closeMenus();

            await activateSheet(
              sheet.id
            );

          }


          if (action === "rename") {

            closeMenus();

            await renameSheet(
              sheet
            );

          }


          if (action === "visibility") {

            closeMenus();

            await toggleVisibility(
              sheet
            );

          }


          if (action === "color") {

            openColorMenu();

          }

        }
      );
    });
}


async function renameSheet(sheet) {

  const newName =
    window.prompt(
      "Worksheet name:",
      sheet.name
    );

  if (!newName) {
    return;
  }

  const cleanedName =
    newName.trim();

  if (!cleanedName) {
    return;
  }


  try {

    await Excel.run(
      async (context) => {

        const worksheet =
          context.workbook
            .worksheets
            .getItem(sheet.id);

        worksheet.name =
          cleanedName;

        await context.sync();
      }
    );

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


async function toggleVisibility(sheet) {

  try {

    await Excel.run(
      async (context) => {

        const sheets =
          context.workbook.worksheets;

        const worksheet =
          sheets.getItem(sheet.id);


        if (
          sheet.visibility === "Visible"
        ) {

          const visibleCount =
            sheetCache.filter(
              item =>
                item.visibility ===
                "Visible"
            ).length;

          if (visibleCount <= 1) {

            window.alert(
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


/* COLOR MENU */

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

  colorMenu.style.top =
    `${rect.top}px`;

  colorMenu.style.right =
    `${rect.width + 10}px`;

  colorMenu.hidden = false;
}


function setupColorMenu() {

  document
    .querySelectorAll(
      ".color-choice"
    )
    .forEach((button) => {

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


async function setTabColor(
  sheetId,
  color
) {

  try {

    await Excel.run(
      async (context) => {

        const worksheet =
          context.workbook
            .worksheets
            .getItem(sheetId);

        worksheet.tabColor =
          color;

        await context.sync();
      }
    );

    await loadSheets();

  } catch (error) {
    showError(error);
  }
}


/* MENU HELPERS */

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

  document.getElementById(
    "sheetMenu"
  ).hidden = true;

  document.getElementById(
    "colorMenu"
  ).hidden = true;
}


/* KEEP ACTIVE SHEET VISIBLE */

function scrollActiveIntoView() {

  requestAnimationFrame(() => {

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

  });
}


/* ERRORS */

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
    document.getElementById(
      "status"
    );

  status.textContent =
    message;

  status.style.color =
    isError
      ? "#ffd7d7"
      : "";
}