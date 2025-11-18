export function initResizablePanes() {
  function makeResizable(resizer, left, right, isVertical = true, isRightSide = false) {
    let isResizing = false;

    resizer.addEventListener("mousedown", (e) => {
      isResizing = true;
      document.body.style.cursor = isVertical ? "col-resize" : "row-resize";
      document.body.style.userSelect = "none";
    });

    document.addEventListener("mousemove", (e) => {
      if (!isResizing) return;

      if (isVertical) {
        if (isRightSide) {
          const workbench = document.querySelector(".workbench");
          if (!workbench || !right) return;
          const workbenchRect = workbench.getBoundingClientRect();
          const rightEdge = workbenchRect.right;
          const newRightWidth = rightEdge - e.clientX;
          const minWidth = 200;
          const maxWidth = workbenchRect.width - 400;

          if (newRightWidth >= minWidth && newRightWidth <= maxWidth && newRightWidth > 0) {
            right.style.setProperty("width", `${newRightWidth}px`, "important");
            right.style.setProperty("min-width", "200px", "important");
            right.style.setProperty("flex", "0 0 auto", "important");
            right.style.setProperty("display", "flex", "important");
            left.style.flex = "1";
          }
        } else {
          const workbench = document.querySelector(".workbench");
          if (!workbench || !left || !right) return;
          const workbenchRect = workbench.getBoundingClientRect();
          const newLeftWidth = e.clientX - workbenchRect.left;
          const minWidth = 150;
          const maxWidth = workbenchRect.width - 400;

          if (newLeftWidth >= minWidth && newLeftWidth <= maxWidth) {
            left.style.setProperty("width", `${newLeftWidth}px`, "important");
            left.style.setProperty("flex", "0 0 auto", "important");
            left.style.setProperty("display", "flex", "important");
            right.style.flex = "1";
          }
        }
      } else {
        const container = left?.parentElement;
        if (!container) return;
        const containerRect = container.getBoundingClientRect();
        const newHeight = e.clientY - containerRect.top;
        const minHeight = 100;
        const maxHeight = containerRect.height - 100;
        if (newHeight > minHeight && newHeight < maxHeight) {
          left.style.setProperty("flex", `0 0 ${newHeight}px`, "important");
          left.style.setProperty("min-height", `${minHeight}px`, "important");
          right.style.setProperty("flex", "1 1 auto", "important");
          right.style.setProperty("min-height", `${minHeight}px`, "important");
        }
      }
    });

    document.addEventListener("mouseup", () => {
      isResizing = false;
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    });
  }

  const leftResizer = document.getElementById("leftResizer");
  const rightResizer = document.getElementById("rightResizer");
  const middleResizer = document.getElementById("middleResizer");
  const sidebar = document.querySelector(".sidebar");
  const middleSection = document.querySelector(".middle-section");
  const formulaEditor = document.querySelector(".formula-editor");
  const gridContainer = document.querySelector(".grid-container");
  const chatContainer = document.querySelector(".chat-container");

  if (sidebar) {
    const viewportWidth = window.innerWidth;
    const sidebarWidth = viewportWidth * 0.15;
    sidebar.style.width = `${sidebarWidth}px`;
    sidebar.style.display = "flex";
  }

  if (formulaEditor) {
    formulaEditor.style.setProperty("width", "400px", "important");
    formulaEditor.style.setProperty("min-width", "200px", "important");
    formulaEditor.style.setProperty("display", "flex", "important");
    formulaEditor.style.setProperty("flex", "0 0 400px", "important");

    setInterval(() => {
      const computedStyle = window.getComputedStyle(formulaEditor);
      const width = computedStyle.width;
      const display = computedStyle.display;
      if (width === "0px" || display === "none" || formulaEditor.offsetWidth === 0) {
        formulaEditor.style.setProperty("width", "400px", "important");
        formulaEditor.style.setProperty("display", "flex", "important");
        formulaEditor.style.setProperty("flex", "0 0 400px", "important");
      }
    }, 100);
  }

  if (leftResizer && sidebar && middleSection) {
    makeResizable(leftResizer, sidebar, middleSection, true, false);
  }

  if (rightResizer && formulaEditor && middleSection) {
    setTimeout(() => {
      makeResizable(rightResizer, middleSection, formulaEditor, true, true);
    }, 500);
  }

  if (middleResizer && gridContainer && chatContainer) {
    makeResizable(middleResizer, gridContainer, chatContainer, false);
  }
}
