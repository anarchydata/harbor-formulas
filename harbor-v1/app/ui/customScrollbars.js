export function initCustomScrollbars() {
        const gridWrapper = document.querySelector(".grid-wrapper");
        const gridInner = document.getElementById("gridWrapperInner");
        const scrollbarV = document.getElementById("customScrollbarV");
        const scrollbarH = document.getElementById("customScrollbarH");

        if (!gridWrapper || !gridInner || !scrollbarV || !scrollbarH) return;

        function updateScrollbars() {
          const scrollHeight = gridInner.scrollHeight;
          const scrollWidth = gridInner.scrollWidth;
          const clientHeight = gridInner.clientHeight;
          const clientWidth = gridInner.clientWidth;
          const scrollTop = gridInner.scrollTop;
          const scrollLeft = gridInner.scrollLeft;

          // Vertical scrollbar
          if (scrollHeight > clientHeight) {
            const thumbHeight = (clientHeight / scrollHeight) * clientHeight;
            const thumbTop = (scrollTop / (scrollHeight - clientHeight)) * (clientHeight - thumbHeight);
            scrollbarV.style.display = "block";
            let thumbV = scrollbarV.querySelector(".custom-scrollbar-thumb");
            if (!thumbV) {
              thumbV = document.createElement("div");
              thumbV.className = "custom-scrollbar-thumb";
              scrollbarV.appendChild(thumbV);
            }
            thumbV.style.height = `${thumbHeight}px`;
            thumbV.style.top = `${thumbTop}px`;
          } else {
            scrollbarV.style.display = "none";
          }

          // Horizontal scrollbar
          if (scrollWidth > clientWidth) {
            const thumbWidth = (clientWidth / scrollWidth) * clientWidth;
            const thumbLeft = (scrollLeft / (scrollWidth - clientWidth)) * (clientWidth - thumbWidth);
            scrollbarH.style.display = "block";
            let thumbH = scrollbarH.querySelector(".custom-scrollbar-thumb");
            if (!thumbH) {
              thumbH = document.createElement("div");
              thumbH.className = "custom-scrollbar-thumb";
              scrollbarH.appendChild(thumbH);
            }
            thumbH.style.width = `${thumbWidth}px`;
            thumbH.style.left = `${thumbLeft}px`;
          } else {
            scrollbarH.style.display = "none";
          }
        }

        // Scrollbar dragging
        let isDraggingV = false;
        let isDraggingH = false;
        let startY = 0;
        let startX = 0;
        let startScrollTop = 0;
        let startScrollLeft = 0;

        scrollbarV.addEventListener("mousedown", (e) => {
          if (e.target.classList.contains("custom-scrollbar-thumb")) {
            isDraggingV = true;
            startY = e.clientY;
            startScrollTop = gridInner.scrollTop;
            e.preventDefault();
          } else {
            // Click on track
            const rect = scrollbarV.getBoundingClientRect();
            const clickY = e.clientY - rect.top;
            const scrollHeight = gridInner.scrollHeight;
            const clientHeight = gridInner.clientHeight;
            const newScrollTop = (clickY / clientHeight) * scrollHeight;
            gridInner.scrollTop = newScrollTop;
          }
        });

        scrollbarH.addEventListener("mousedown", (e) => {
          if (e.target.classList.contains("custom-scrollbar-thumb")) {
            isDraggingH = true;
            startX = e.clientX;
            startScrollLeft = gridInner.scrollLeft;
            e.preventDefault();
          } else {
            // Click on track
            const rect = scrollbarH.getBoundingClientRect();
            const clickX = e.clientX - rect.left;
            const scrollWidth = gridInner.scrollWidth;
            const clientWidth = gridInner.clientWidth;
            const newScrollLeft = (clickX / clientWidth) * scrollWidth;
            gridInner.scrollLeft = newScrollLeft;
          }
        });

        document.addEventListener("mousemove", (e) => {
          if (isDraggingV) {
            const deltaY = e.clientY - startY;
            const scrollHeight = gridInner.scrollHeight;
            const clientHeight = gridInner.clientHeight;
            const scrollRatio = deltaY / clientHeight;
            gridInner.scrollTop = startScrollTop + (scrollRatio * scrollHeight);
          }
          if (isDraggingH) {
            const deltaX = e.clientX - startX;
            const scrollWidth = gridInner.scrollWidth;
            const clientWidth = gridInner.clientWidth;
            const scrollRatio = deltaX / clientWidth;
            gridInner.scrollLeft = startScrollLeft + (scrollRatio * scrollWidth);
          }
        });

        document.addEventListener("mouseup", () => {
          isDraggingV = false;
          isDraggingH = false;
        });

        gridInner.addEventListener("scroll", () => {
          updateScrollbars();
          const hasSelection = window.selectedCells && window.selectedCells.size > 0;
          if (hasSelection && typeof window.updateSelectionOverlay === "function") {
            window.updateSelectionOverlay();
          }
        });
        window.addEventListener("resize", updateScrollbars);
        updateScrollbars();
      }
