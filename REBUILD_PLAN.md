# Rebuild Plan: Strip Interactivity & Rebuild in TypeScript

## Strategy Overview

**Core Idea**: Kill all event listeners and interactive code in `bootstrap.js`, keep only core initialization. Then rebuild all interactivity from scratch in TypeScript with proper architecture.

**Key Principle**: Set up reusable infrastructure (TypeScript config, linting rules, base patterns) that will automatically apply to:
- Current grid functionality
- Future file explorer pane
- Future code editor pane (VS Code native)
- Future Clappy chat
- Any other features we add later

This means when we add new features, we just follow the established patterns - no need to reconfigure TypeScript, ESLint, or create new base classes.

**What We Keep** (core business logic, no events):
- `utils/` - helpers.js, diagnostics.js, autofill.js (pure business logic)
- `app/grid/` - handsontableAdapter.js, hyperformulaAdapter.js (grid configuration)
- `components/grid.js` - grid component logic
- `index.html` - HTML structure (just the containers/windows, no interactivity)

**What We Rebuild** (all interactivity/event handlers):
- **ALL** of `app/ui/` - customScrollbars.js, resizablePanes.js, clappyChat.js (delete - we'll rebuild when ready)
- All event listeners in `bootstrap.js` (19 found)
- File tree collapse/expand logic (will add file explorer pane later)
- All UI controllers and interactivity

**What We'll Add Later** (not in initial rebuild):
- File explorer pane (left side) - will be added when ready
- Code editor pane (right side) - will use native VS Code editor (not Monaco)
- Clappy chat - will be added back when ready

## Phase 1: Strip Down Bootstrap.js

### 1.1 Identify What to Keep (Core Initialization Only)

**KEEP** (minimal initialization):
- HyperFormula instance creation
- Monaco Editor setup (basic initialization, no event handlers)
- Handsontable grid creation (basic setup, no event handlers)
- Basic DOM element references (for services to use)
- Sheet creation (Sheet1)

**REMOVE** (all interactivity):
- All `addEventListener` calls (19 found)
- All event handler functions
- All interactive state management
- All user input handling
- All keyboard shortcuts
- All mouse interaction handlers
- All formula editor interactivity
- All cell selection logic
- All autofill handlers
- All sheet management UI handlers
- All pane collapse/expand handlers
- All button click handlers

### 1.2 Create Minimal Bootstrap

The stripped `bootstrap.js` should be ~200-300 lines:
```javascript
// Minimal initialization only
export async function initializeApp() {
  // 1. Initialize HyperFormula
  const { HyperFormula } = await import('hyperformula');
  window.hf = HyperFormula.buildEmpty({...});
  window.hf.addSheet('Sheet1');
  
  // 2. Get DOM references
  const domRefs = {
    monacoEditor: document.getElementById('monacoEditor'),
    handsontableRoot: document.getElementById('handsontableRoot'),
    // ... other essential refs
  };
  
  // 3. Initialize Monaco (basic setup, no handlers)
  // 4. Initialize Grid (basic setup, no handlers)
  
  // 5. Return services/refs for TypeScript layer
  return { hf: window.hf, domRefs };
}
```

## Phase 2: Scaffold TypeScript Project Structure

### 2.1 Project Setup (Reusable Infrastructure)

**Install TypeScript tooling** (applies to all future code):
```bash
# Install TypeScript tooling
npm install --save-dev typescript @types/node
npm install --save-dev @typescript-eslint/eslint-plugin @typescript-eslint/parser
npm install --save-dev eslint prettier eslint-config-prettier
npm install --save-dev @types/handsontable

# Note: @types/monaco-editor not needed if using VS Code native editor later
```

**Key Principle**: Set up infrastructure that will work for:
- Current grid functionality
- Future file explorer pane
- Future code editor pane (VS Code native)
- Future Clappy chat
- Any other features we add later

### 2.2 New Architecture Structure (Extensible Pattern)

```
app/
  core/
    services/
      HyperFormulaService.ts      - HyperFormula management
      GridService.ts               - Grid service
      SheetService.ts              - Sheet management
      BaseService.ts               - Abstract base class for all services (reusable pattern)
    controllers/
      BaseController.ts            - Abstract base class for all controllers (reusable pattern)
    types/
      dom-refs.ts                  - DOM element type definitions (reusable)
      grid-types.ts                - Grid-related types
      common-types.ts              - Common types for all features (reusable)
      service-types.ts             - Service interface types (reusable pattern)
  features/
    interactivity/
      CellSelection.ts             - Cell selection handlers
      FormulaEditor.ts             - Formula editor interactivity
      KeyboardShortcuts.ts         - Global keyboard handlers
      Autofill.ts                  - Autofill drag handlers
    ui/
      SheetTabsController.ts       - Sheet tab interactions
      PaneController.ts            - Bottom panes (result, messages, dependencies)
      ButtonHandlers.ts            - All button click handlers
  patterns/
    ServicePattern.ts              - Reusable service pattern/template
    ControllerPattern.ts           - Reusable controller pattern/template
    EventHandlerPattern.ts         - Reusable event handler pattern
  bootstrap.ts                     - Clean orchestration (~100-200 lines)
```

**Key Principle**: Create abstract patterns that can be reused when adding:
- File explorer pane (will use BaseController pattern)
- Code editor pane (will use BaseService + BaseController patterns)
- Clappy chat (will use BaseController pattern)
- Any future features

## Phase 3: Identify All Interactive Features

### 3.1 Event Listeners Found (19 total)

1. **Folder collapse** - `projectFolder.addEventListener("click")`
2. **Sidebar toggle** - `sidebarToggleButton.addEventListener("click")`
3. **Sheet rename** - `sheetTabsBarElement.addEventListener("dblclick")`
4. **Sheet rename input** - `renameInput.addEventListener("keydown")` + `blur`
5. **Resizer handles** - Multiple `handle.addEventListener("mousedown")`
6. **File tree items** - `item.addEventListener('click')`
7. **Monaco editor click** - `editorDomNode.addEventListener('click')`
8. **Fill handle drag** - `fillHandleElement.addEventListener("mousedown")` + document mousemove/mouseup
9. **Selection drag** - Document mousemove/mouseup for selection
10. **Global mouseup** - `document.addEventListener("mouseup")`
11. **Global keyboard** - `document.addEventListener("keydown", handleGlobalKeyDown, true)`
12. **Pane headers** - `header.addEventListener('click')` (result, messages, dependencies)
13. **Word wrap button** - `wordWrapBtn.addEventListener('click')`
14. **Tidy button** - `broomBtn.addEventListener('click')`
15. **Window resize** - `window.addEventListener('resize')`
16. **Chat form** - (in clappyChat.js)
17. **Resizable panes** - (in resizablePanes.js)
18. **Custom scrollbars** - (in customScrollbars.js)

### 3.2 Complex Interactive Features

1. **Cell Selection System**
   - Click selection
   - Drag selection
   - Keyboard navigation
   - Multi-cell selection
   - Selection highlighting
   - Edge handles

2. **Formula Editor**
   - Monaco editor integration
   - Cell reference chips
   - Auto-completion
   - Syntax highlighting
   - Formula parsing
   - Enter mode pointer navigation
   - Range selection in formulas

3. **Autofill System**
   - Fill handle drag
   - Formula reference adjustment
   - Value series generation
   - Visual feedback

4. **Keyboard Shortcuts**
   - Global keydown handler
   - Grid navigation
   - Editor shortcuts
   - Mode switching (select/edit)

5. **Sheet Management**
   - Tab creation
   - Tab renaming (double-click)
   - Tab switching
   - Sheet deletion

6. **UI Controls**
   - Sidebar collapse
   - Chat panel collapse
   - Pane collapse/expand
   - Status bar updates
   - Button handlers (word wrap, tidy, collapse errors)

7. **Monaco Editor Interactions**
   - Click to insert cell references
   - Cursor position tracking
   - Line/column status updates
   - Content change handlers

8. **Grid Interactions**
   - Handsontable event handling
   - Cell editing
   - Row/column header interactions
   - Scroll synchronization

## Phase 4: Rebuild Strategy by Feature

### 4.1 Core Services (No Interactivity)

**HyperFormulaService.ts**
```typescript
export class HyperFormulaService {
  private hf: HyperFormula;
  
  constructor() {
    this.hf = HyperFormula.buildEmpty({...});
  }
  
  getInstance(): HyperFormula { return this.hf; }
  addSheet(name: string): string { ... }
  renameSheet(id: number, name: string): void { ... }
  // Pure service methods, no event handlers
}
```

**EditorService.ts** (Optional - for later when we add code editor pane)
```typescript
// Will use native VS Code editor instead of Monaco
// This will be implemented later when we add the code editor pane
// For now, we focus on grid and formula editing only
```

**GridService.ts**
```typescript
export class GridService extends BaseService {
  private hot: Handsontable;
  
  constructor(container: HTMLElement, hfService: HyperFormulaService) {
    super();
    this.hot = new Handsontable(container, {...});
  }
  
  getInstance(): Handsontable { return this.hot; }
  // Pure service methods, no event handlers
}
```

**Reusable Base Patterns** (for all future features):

**BaseService.ts** (Abstract pattern):
```typescript
export abstract class BaseService {
  protected initialized = false;
  
  abstract initialize(): Promise<void> | void;
  abstract destroy(): void;
  
  isInitialized(): boolean {
    return this.initialized;
  }
}
```

**BaseController.ts** (Abstract pattern):
```typescript
export abstract class BaseController {
  protected eventListeners: Array<{
    element: EventTarget;
    event: string;
    handler: EventListener;
  }> = [];
  
  abstract initialize(): void;
  
  protected addEventListener(
    element: EventTarget,
    event: string,
    handler: EventListener
  ): void {
    element.addEventListener(event, handler);
    this.eventListeners.push({ element, event, handler });
  }
  
  destroy(): void {
    this.eventListeners.forEach(({ element, event, handler }) => {
      element.removeEventListener(event, handler);
    });
    this.eventListeners = [];
  }
}
```

**Note**: These patterns will be reused when adding:
- FileExplorerService (extends BaseService)
- CodeEditorService (extends BaseService)  
- FileExplorerController (extends BaseController)
- CodeEditorController (extends BaseController)
- ClappyChatController (extends BaseController)

### 4.2 Interactivity Modules (Event Handlers)

**CellSelection.ts**
```typescript
export class CellSelectionController extends BaseController {
  constructor(
    private gridService: GridService,
    private hfService: HyperFormulaService
  ) {
    super();
  }
  
  initialize(): void {
    // Attach all selection-related event handlers
    this.attachClickHandlers();
    this.attachDragHandlers();
    this.attachKeyboardHandlers();
  }
  
  private attachClickHandlers(): void { ... }
  private attachDragHandlers(): void { ... }
  private attachKeyboardHandlers(): void { ... }
}
```

**FormulaEditor.ts**
```typescript
export class FormulaEditorController {
  constructor(
    private gridService: GridService,
    private hfService: HyperFormulaService
  ) {}
  
  initialize(): void {
    // Attach all formula editor event handlers
    // Note: For now, this handles formula editing in the grid
    // Code editor pane with VS Code native editor will be added later
    this.attachEditorClickHandlers();
    this.attachContentChangeHandlers();
    this.attachCursorHandlers();
  }
  
  private attachEditorClickHandlers(): void { ... }
  private attachContentChangeHandlers(): void { ... }
  // ... etc
}
```

**KeyboardShortcuts.ts**
```typescript
export class KeyboardShortcutsController {
  constructor(
    private gridService: GridService,
    private selectionController: CellSelectionController
  ) {}
  
  initialize(): void {
    document.addEventListener('keydown', this.handleGlobalKeyDown.bind(this), true);
  }
  
  private handleGlobalKeyDown(event: KeyboardEvent): void {
    // All keyboard shortcut logic (grid-focused)
  }
}
```

**Autofill.ts**
```typescript
export class AutofillController {
  constructor(
    private gridService: GridService,
    private hfService: HyperFormulaService
  ) {}
  
  initialize(): void {
    const fillHandle = document.getElementById('fillHandle');
    if (fillHandle) {
      fillHandle.addEventListener('mousedown', this.handleFillStart.bind(this));
    }
    document.addEventListener('mousemove', this.handleFillDrag.bind(this));
    document.addEventListener('mouseup', this.handleFillEnd.bind(this));
  }
  
  private handleFillStart(event: MouseEvent): void { ... }
  private handleFillDrag(event: MouseEvent): void { ... }
  private handleFillEnd(event: MouseEvent): void { ... }
}
```

**SheetTabsController.ts**
```typescript
export class SheetTabsController {
  constructor(
    private sheetService: SheetService,
    private gridService: GridService
  ) {}
  
  initialize(): void {
    const tabsBar = document.getElementById('sheetTabsBar');
    if (tabsBar) {
      tabsBar.addEventListener('dblclick', this.handleTabDoubleClick.bind(this));
    }
    // Tab add button, etc.
  }
  
  private handleTabDoubleClick(event: MouseEvent): void {
    // Sheet rename logic
  }
}
```

**ButtonHandlers.ts**
```typescript
export class ButtonHandlersController {
  constructor(private monacoService: MonacoService) {}
  
  initialize(): void {
    this.attachWordWrapButton();
    this.attachTidyButton();
    this.attachCollapseErrorsButton();
  }
  
  private attachWordWrapButton(): void {
    const btn = document.getElementById('wordWrapBtn');
    if (btn) {
      btn.addEventListener('click', () => { ... });
    }
  }
  
  // ... etc
}
```

**PaneController.ts**
```typescript
export class PaneController {
  initialize(): void {
    document.querySelectorAll('.pane-header').forEach(header => {
      header.addEventListener('click', () => {
        const pane = header.closest('.sliding-pane');
        if (pane) {
          pane.classList.toggle('collapsed');
          this.updatePaneHeight(pane);
        }
      });
    });
  }
  
  private updatePaneHeight(pane: HTMLElement): void { ... }
}
```

### 4.3 UI Controllers (Rebuild from Scratch in TS)

**SidebarController.ts** (rebuild from scratch)
```typescript
export class SidebarController {
  initialize(): void {
    const sidebar = document.querySelector('.sidebar');
    const toggle = document.getElementById('sidebarCollapseToggle');
    if (sidebar && toggle) {
      toggle.addEventListener('click', () => {
        const isCollapsed = sidebar.classList.toggle('collapsed');
        // ... styling logic
      });
    }
  }
}
```

**FileTreeController.ts** (rebuild from scratch - simple collapse/expand)
```typescript
export class FileTreeController {
  initialize(): void {
    const projectFolder = document.getElementById('projectFolder');
    const projectFolderContent = document.getElementById('projectFolderContent');
    if (projectFolder && projectFolderContent) {
      projectFolder.addEventListener('click', () => {
        const isCollapsed = projectFolderContent.classList.toggle('collapsed');
        // Update chevron, etc.
      });
    }
  }
}
```

**Rebuild existing UI modules from scratch in TypeScript (when ready):**
- `app/ui/customScrollbars.js` → `app/features/ui/CustomScrollbarsController.ts` (rebuild later)
- `app/ui/resizablePanes.js` → `app/features/ui/ResizablePanesController.ts` (rebuild later)
- File explorer pane - new, will be added on left side
- Code editor pane - new, will use native VS Code editor (not Monaco)
- Clappy chat - will be added back when ready

**Note**: For the initial rebuild, we focus on core services and grid interactivity only. UI panes (file explorer, code editor, chat) will be added later as separate features.

## Phase 5: New Bootstrap.ts (Clean Orchestration)

```typescript
import { HyperFormulaService } from './core/services/HyperFormulaService';
import { GridService } from './core/services/GridService';
import { SheetService } from './core/services/SheetService';

import { CellSelectionController } from './features/interactivity/CellSelection';
import { FormulaEditorController } from './features/interactivity/FormulaEditor';
import { KeyboardShortcutsController } from './features/interactivity/KeyboardShortcuts';
import { AutofillController } from './features/interactivity/Autofill';
import { SheetTabsController } from './features/ui/SheetTabsController';
import { PaneController } from './features/ui/PaneController';
import { ButtonHandlersController } from './features/ui/ButtonHandlers';

export async function initializeApp() {
  // 1. Initialize core services (no interactivity)
  const hfService = new HyperFormulaService();
  hfService.addSheet('Sheet1');
  
  const gridContainer = document.getElementById('handsontableRoot');
  if (!gridContainer) throw new Error('Grid container not found');
  const gridService = new GridService(gridContainer, hfService);
  
  const sheetService = new SheetService(hfService, gridService);
  
  // 2. Initialize interactivity controllers (core grid functionality)
  const selectionController = new CellSelectionController(
    gridService,
    hfService
  );
  selectionController.initialize();
  
  const formulaEditorController = new FormulaEditorController(
    gridService,
    hfService
  );
  formulaEditorController.initialize();
  
  const keyboardController = new KeyboardShortcutsController(
    gridService,
    selectionController
  );
  keyboardController.initialize();
  
  const autofillController = new AutofillController(gridService, hfService);
  autofillController.initialize();
  
  // 3. Initialize UI controllers (minimal - just grid-related)
  const sheetTabsController = new SheetTabsController(sheetService, gridService);
  sheetTabsController.initialize();
  
  const paneController = new PaneController();
  paneController.initialize();
  
  const buttonHandlers = new ButtonHandlersController();
  buttonHandlers.initialize();
  
  // 4. Store services globally (for backward compatibility)
  window.hf = hfService.getInstance();
  
  // Note: File explorer pane, code editor pane (VS Code native), and Clappy chat
  // will be added later as separate features
  
  return {
    hfService,
    gridService,
    sheetService,
  };
}
```

## Phase 6: Migration Steps

### Step 1: Setup TypeScript Infrastructure (Reusable for All Future Code)
- [ ] Install TypeScript dependencies
- [ ] Create `tsconfig.json` (will apply to all future features)
- [ ] Create VS Code workspace config (linting/formatting for all code)
- [ ] Create ESLint config (rules apply to all future code)
- [ ] Create Prettier config (formatting for all future code)
- [ ] Create `types/global.d.ts` (reusable type definitions)
- [ ] Create `app/core/BaseService.ts` (reusable service pattern)
- [ ] Create `app/core/BaseController.ts` (reusable controller pattern)
- [ ] Create `app/core/types/service-types.ts` (reusable service interfaces)

### Step 2: Create Core Services (No Events) - Using Base Patterns
- [ ] `BaseService.ts` (abstract base class - reusable pattern)
- [ ] `BaseController.ts` (abstract base class - reusable pattern)
- [ ] `HyperFormulaService.ts` (extends BaseService)
- [ ] `GridService.ts` (extends BaseService)
- [ ] `SheetService.ts` (extends BaseService)
- [ ] Type definitions (`types/` - reusable for all features)

### Step 3: Strip Bootstrap.js
- [ ] Remove all event listeners
- [ ] Remove all handler functions
- [ ] Keep only core initialization
- [ ] Test that app loads (but nothing works interactively)

### Step 4: Build Interactivity Layer (One at a time)
- [ ] `CellSelectionController.ts` - Test cell selection works
- [ ] `FormulaEditorController.ts` - Test formula editing works
- [ ] `KeyboardShortcutsController.ts` - Test keyboard shortcuts
- [ ] `AutofillController.ts` - Test autofill
- [ ] `SheetTabsController.ts` - Test sheet management
- [ ] `ButtonHandlersController.ts` - Test buttons
- [ ] `PaneController.ts` - Test panes
- [ ] `SidebarController.ts` - Test sidebar

### Step 5: Rebuild UI Controllers from Scratch (Minimal - Grid Focused)
- [ ] `SheetTabsController.ts` - Sheet tab interactions
- [ ] `PaneController.ts` - Bottom panes (result, messages, dependencies)
- [ ] `ButtonHandlersController.ts` - Button handlers

**Deferred (will add later):**
- File explorer pane (left side)
- Code editor pane (right side - VS Code native)
- CustomScrollbarsController
- ResizablePanesController
- ClappyChatController

### Step 6: Create New Bootstrap.ts
- [ ] Write clean orchestration code
- [ ] Wire up all controllers
- [ ] Test full application

### Step 7: Cleanup
- [ ] Remove old `bootstrap.js`
- [ ] Remove old `app/ui/` JavaScript files (we'll rebuild them later when needed)
- [ ] Update `main.ts` to import from `bootstrap.ts`
- [ ] Remove any remaining global state
- [ ] Final testing (grid functionality only)

### Step 8: Future Enhancements (Separate Work - Using Established Patterns)

**When adding file explorer pane:**
- [ ] Create `FileExplorerService.ts` (extends BaseService)
- [ ] Create `FileExplorerController.ts` (extends BaseController)
- [ ] Use existing TypeScript config (already set up)
- [ ] Use existing linting rules (already configured)
- [ ] Follow established service/controller patterns

**When adding code editor pane (VS Code native):**
- [ ] Create `CodeEditorService.ts` (extends BaseService)
- [ ] Create `CodeEditorController.ts` (extends BaseController)
- [ ] Use existing TypeScript config (already set up)
- [ ] Use existing linting rules (already configured)
- [ ] Follow established service/controller patterns

**When adding Clappy chat:**
- [ ] Create `ClappyChatController.ts` (extends BaseController)
- [ ] Use existing TypeScript config (already set up)
- [ ] Use existing linting rules (already configured)
- [ ] Follow established controller pattern

**When adding other features:**
- [ ] All use same TypeScript infrastructure
- [ ] All use same linting/formatting rules
- [ ] All follow same BaseService/BaseController patterns

## Benefits of This Approach

1. **Clean Separation**: Services (no events) vs Controllers (only events)
2. **Testable**: Each controller can be tested independently
3. **Maintainable**: Clear responsibility boundaries
4. **Type-Safe**: Full TypeScript from the start
5. **Incremental**: Build and test one feature at a time
6. **No Legacy Baggage**: Fresh start for interactivity layer
7. **Keep What Works**: Preserve business logic (utils, grid adapters)
8. **Rebuild What's Simple**: Event handlers are easier to rebuild than migrate
9. **Reusable Infrastructure**: TypeScript config, linting rules, and patterns apply to all future code
10. **Extensible Architecture**: Base classes and interfaces make adding new features easy
11. **Consistent Code Style**: ESLint and Prettier ensure all code (current and future) follows same rules

## Estimated Timeline

- **Phase 1 (Strip)**: 2-3 hours
- **Phase 2 (Scaffold)**: 2-3 hours
- **Phase 3 (Core Services)**: 4-6 hours
- **Phase 4 (Interactivity)**: 10-14 hours
  - CellSelection: 3-4 hours
  - FormulaEditor: 3-4 hours (grid-focused, no Monaco for now)
  - KeyboardShortcuts: 2-3 hours
  - Autofill: 1-2 hours
  - UI Controllers (minimal - grid-focused): 1-2 hours
    - SheetTabs: 30 min
    - PaneController: 30 min
    - ButtonHandlers: 30 min
- **Phase 5 (Bootstrap)**: 2-3 hours
- **Phase 6 (Testing)**: 4-6 hours

**Total**: 22-32 hours (reduced scope - no Monaco, no Clappy, no file explorer initially)

**Future Additions** (separate work):
- File explorer pane: 2-3 hours
- Code editor pane (VS Code native): 4-6 hours
- Clappy chat: 1-2 hours
- Custom scrollbars: 1 hour
- Resizable panes: 1 hour

## Risk Mitigation

1. **Keep old bootstrap.js** as backup until new one works
2. **Build incrementally** - one controller at a time
3. **Test after each controller** - don't build everything then test
4. **Use feature flags** - can switch between old/new implementations
5. **Document state** - what works, what doesn't at each step

## Key Decisions

1. **Dependency Injection**: Controllers receive services as constructor params
2. **No Global State**: Services manage state, controllers handle events
3. **Event Delegation**: Use event delegation where possible for performance
4. **Type Safety**: Strict TypeScript, no `any` types
5. **Modular**: Each feature is self-contained

