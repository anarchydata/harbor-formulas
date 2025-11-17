# TypeScript Migration Plan for Harbor Formulas

## Overview
This document outlines the plan to migrate the Harbor Formulas project from vanilla JavaScript to TypeScript, with full VS Code support for linting, formatting, and IntelliSense.

## Current State Analysis

### Project Structure
- **Build Tool**: Vite 7.2.2
- **Language**: Vanilla JavaScript (ES Modules)
- **Entry Point**: `app/main.js` → `app/bootstrap.js` (7490 lines!)
- **Key Dependencies**:
  - `handsontable` ^16.1.1
  - `hyperformula` ^3.1.0
  - `monaco-editor` ^0.53.0
- **Module Organization**:
  - `app/` - Main application code
  - `components/` - Component modules
  - `utils/` - Utility functions
  - `app/grid/` - Grid adapters
  - `app/ui/` - UI components
  - `app/setup/` - Setup/config modules

### Key Challenges
1. **Large bootstrap.js file** (7490 lines) - needs to be broken down
2. **No type definitions** for custom code
3. **Global window properties** (window.hf, window.monacoEditor, etc.)
4. **Dynamic imports** and runtime dependencies
5. **Monaco Editor** integration requires special handling

## Migration Strategy

### Phase 1: TypeScript Setup & Configuration

#### 1.1 Install TypeScript Dependencies
```bash
npm install --save-dev typescript @types/node
npm install --save-dev @typescript-eslint/eslint-plugin @typescript-eslint/parser
npm install --save-dev eslint prettier eslint-config-prettier
npm install --save-dev @types/handsontable @types/monaco-editor
```

#### 1.2 Create TypeScript Configuration
- **`tsconfig.json`** - Main TypeScript compiler configuration
  - Target: ES2020 or ES2022
  - Module: ES2020/ESNext
  - Module Resolution: bundler (for Vite)
  - Strict mode enabled
  - Include: `app/`, `components/`, `utils/`
  - Exclude: `node_modules/`, `dist/`

#### 1.3 Update Vite Configuration
- Rename `vite.config.js` → `vite.config.ts`
- Add TypeScript plugin support
- Ensure proper handling of `.ts` and `.tsx` files
- Maintain Monaco Editor plugin configuration

#### 1.4 VS Code Workspace Configuration
- **`.vscode/settings.json`**:
  - TypeScript as default formatter
  - ESLint integration
  - Format on save
  - Type checking on save
- **`.vscode/extensions.json`**:
  - Recommended extensions (ESLint, Prettier, TypeScript)
- **`.eslintrc.json`** or **`.eslintrc.js`**:
  - TypeScript parser
  - Recommended rules
  - Prettier integration
- **`.prettierrc`**:
  - Code formatting rules

### Phase 2: Type Definitions & Declarations

#### 2.1 Global Type Declarations
Create `types/global.d.ts`:
- `window.hf` - HyperFormula instance
- `window.monacoEditor` - Monaco Editor instance
- `window.rawFormulaStore` - Map<string, string>
- `window.namedRanges` - Set<string>
- `window.hfSheetId` - number
- `window.selectedCell` - HTMLElement | null
- `window.isProgrammaticCursorChange` - boolean

#### 2.2 Third-Party Type Definitions
- Check for `@types/handsontable` availability
- Check for `@types/hyperformula` availability
- Create custom type definitions if needed in `types/`
- Monaco Editor types should be available via `monaco-editor` package

#### 2.3 Module Type Definitions
- Create type definitions for:
  - Grid adapters (`app/grid/`)
  - UI components (`app/ui/`)
  - Utilities (`utils/`)
  - Helpers and helpers functions

### Phase 3: Incremental File Migration

#### 3.1 Migration Order (Bottom-Up Approach)
1. **Utility Files** (`utils/`)
   - `utils/helpers.ts`
   - `utils/diagnostics.ts`
   - `utils/autofill.ts`
   - These are leaf dependencies, easiest to migrate first

2. **Component Files** (`components/`)
   - `components/grid.ts`
   - Add proper types for grid state and functions

3. **Setup/Config Files** (`app/setup/`)
   - `app/setup/monacoEnv.ts`
   - Simple configuration files

4. **UI Components** (`app/ui/`)
   - `app/ui/customScrollbars.ts`
   - `app/ui/resizablePanes.ts`
   - `app/ui/clappyChat.ts`
   - DOM manipulation with proper types

5. **Grid Adapters** (`app/grid/`)
   - `app/grid/handsontableAdapter.ts`
   - `app/grid/hyperformulaAdapter.ts`
   - Complex integrations, need careful typing

6. **Main Entry Point** (`app/main.ts`)
   - Simple entry point, easy migration

7. **Bootstrap File** (`app/bootstrap.ts`) - **BIGGEST CHALLENGE**
   - **Strategy**: Break down into smaller modules BEFORE migration
   - Extract logical sections:
     - HyperFormula initialization
     - Monaco Editor setup
     - Event handlers
     - Grid initialization
     - UI initialization
     - Sheet management
     - Formula handling
   - Create separate modules in `app/core/` or `app/features/`
   - Migrate each extracted module to TypeScript
   - Finally migrate the orchestration code

#### 3.2 Bootstrap.js Refactoring Strategy
The 7490-line bootstrap.js should be broken down into:

**Suggested Module Structure:**
```
app/
  core/
    hyperformula.ts          - HyperFormula initialization & management
    monaco.ts                - Monaco Editor setup & configuration
    grid.ts                  - Grid initialization & management
    sheets.ts                 - Sheet management logic
  features/
    formula-editor.ts        - Formula editor functionality
    cell-selection.ts        - Cell selection logic
    autofill.ts              - Autofill functionality
    diagnostics.ts           - Formula diagnostics
  bootstrap.ts               - Main orchestration (much smaller)
```

### Phase 4: Type Safety Improvements

#### 4.1 Add Strict Type Checking
- Enable `strict: true` in tsconfig.json
- Fix `any` types gradually
- Add proper return types to all functions
- Add parameter types

#### 4.2 DOM Type Safety
- Use proper DOM types (HTMLElement, HTMLInputElement, etc.)
- Type querySelector results
- Add null checks where needed

#### 4.3 API Type Safety
- Type HyperFormula API calls
- Type Handsontable API calls
- Type Monaco Editor API calls

### Phase 5: Build & Development Workflow

#### 5.1 Update Package.json Scripts
```json
{
  "scripts": {
    "dev": "vite",
    "build": "tsc && vite build",
    "preview": "vite preview",
    "type-check": "tsc --noEmit",
    "lint": "eslint . --ext .ts,.tsx",
    "format": "prettier --write \"**/*.{ts,tsx,js,jsx,json,css}\""
  }
}
```

#### 5.2 Update HTML Entry Point
- Change `app/main.js` → `app/main.ts` in `index.html`
- Vite will handle the TypeScript compilation

### Phase 6: Testing & Validation

#### 6.1 Type Checking
- Run `tsc --noEmit` to check for type errors
- Fix all type errors before proceeding

#### 6.2 Runtime Testing
- Test all functionality after migration
- Verify Monaco Editor still works
- Verify Handsontable integration
- Verify HyperFormula calculations

#### 6.3 Linting & Formatting
- Ensure ESLint passes
- Ensure Prettier formatting is consistent
- Set up pre-commit hooks (optional)

## File Structure After Migration

```
harbor-formulas/
├── .vscode/
│   ├── settings.json
│   ├── extensions.json
│   └── launch.json (optional)
├── types/
│   ├── global.d.ts
│   ├── handsontable.d.ts (if needed)
│   └── hyperformula.d.ts (if needed)
├── app/
│   ├── main.ts
│   ├── bootstrap.ts (refactored, much smaller)
│   ├── core/
│   │   ├── hyperformula.ts
│   │   ├── monaco.ts
│   │   ├── grid.ts
│   │   └── sheets.ts
│   ├── features/
│   │   ├── formula-editor.ts
│   │   ├── cell-selection.ts
│   │   ├── autofill.ts
│   │   └── diagnostics.ts
│   ├── grid/
│   │   ├── handsontableAdapter.ts
│   │   └── hyperformulaAdapter.ts
│   ├── ui/
│   │   ├── customScrollbars.ts
│   │   ├── resizablePanes.ts
│   │   └── clappyChat.ts
│   └── setup/
│       └── monacoEnv.ts
├── components/
│   └── grid.ts
├── utils/
│   ├── helpers.ts
│   ├── diagnostics.ts
│   └── autofill.ts
├── tsconfig.json
├── vite.config.ts
├── .eslintrc.json
├── .prettierrc
└── package.json
```

## Configuration Files to Create

### tsconfig.json
```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ESNext",
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "moduleResolution": "bundler",
    "resolveJsonModule": true,
    "allowJs": true,
    "checkJs": false,
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true,
    "skipLibCheck": true,
    "esModuleInterop": true,
    "allowSyntheticDefaultImports": true,
    "forceConsistentCasingInFileNames": true,
    "isolatedModules": true,
    "noEmit": true
  },
  "include": [
    "app/**/*",
    "components/**/*",
    "utils/**/*",
    "types/**/*"
  ],
  "exclude": [
    "node_modules",
    "dist"
  ]
}
```

### .vscode/settings.json
```json
{
  "editor.defaultFormatter": "esbenp.prettier-vscode",
  "editor.formatOnSave": true,
  "editor.codeActionsOnSave": {
    "source.fixAll.eslint": "explicit"
  },
  "typescript.tsdk": "node_modules/typescript/lib",
  "typescript.enablePromptUseWorkspaceTsdk": true,
  "[typescript]": {
    "editor.defaultFormatter": "esbenp.prettier-vscode"
  },
  "[javascript]": {
    "editor.defaultFormatter": "esbenp.prettier-vscode"
  }
}
```

### .vscode/extensions.json
```json
{
  "recommendations": [
    "dbaeumer.vscode-eslint",
    "esbenp.prettier-vscode",
    "ms-vscode.vscode-typescript-next"
  ]
}
```

## Migration Checklist

### Setup Phase
- [ ] Install TypeScript and related dependencies
- [ ] Create `tsconfig.json`
- [ ] Create `.vscode/settings.json`
- [ ] Create `.vscode/extensions.json`
- [ ] Create `.eslintrc.json`
- [ ] Create `.prettierrc`
- [ ] Update `vite.config.ts`
- [ ] Create `types/global.d.ts`

### Type Definitions
- [ ] Define global window types
- [ ] Install/verify third-party type definitions
- [ ] Create custom type definitions for missing types

### File Migration (Bottom-Up)
- [ ] Migrate `utils/helpers.ts`
- [ ] Migrate `utils/diagnostics.ts`
- [ ] Migrate `utils/autofill.ts`
- [ ] Migrate `components/grid.ts`
- [ ] Migrate `app/setup/monacoEnv.ts`
- [ ] Migrate `app/ui/customScrollbars.ts`
- [ ] Migrate `app/ui/resizablePanes.ts`
- [ ] Migrate `app/ui/clappyChat.ts`
- [ ] Migrate `app/grid/handsontableAdapter.ts`
- [ ] Migrate `app/grid/hyperformulaAdapter.ts`
- [ ] **Refactor `app/bootstrap.js` into smaller modules**
- [ ] Migrate extracted bootstrap modules
- [ ] Migrate `app/main.ts`

### Testing & Validation
- [ ] Run `tsc --noEmit` - fix all type errors
- [ ] Run `npm run lint` - fix all linting errors
- [ ] Run `npm run format` - ensure consistent formatting
- [ ] Test application in dev mode
- [ ] Test build process
- [ ] Verify all functionality works

## Estimated Timeline

- **Phase 1 (Setup)**: 1-2 hours
- **Phase 2 (Type Definitions)**: 2-3 hours
- **Phase 3 (File Migration)**: 8-12 hours
  - Utilities: 1-2 hours
  - Components: 1-2 hours
  - UI modules: 2-3 hours
  - Grid adapters: 2-3 hours
  - **Bootstrap refactoring**: 4-6 hours (biggest task)
- **Phase 4 (Type Safety)**: 2-4 hours
- **Phase 5 (Build Workflow)**: 1 hour
- **Phase 6 (Testing)**: 2-3 hours

**Total Estimated Time**: 16-25 hours

## Risks & Considerations

1. **Bootstrap.js Size**: The 7490-line file is the biggest risk. Breaking it down first is critical.
2. **Third-Party Types**: Some packages may not have perfect type definitions. May need custom declarations.
3. **Global State**: Heavy use of `window.*` properties needs careful typing.
4. **Monaco Editor**: Complex integration may require custom type definitions.
5. **Runtime Behavior**: TypeScript won't catch runtime errors, thorough testing is essential.
6. **Build Performance**: TypeScript compilation adds build time, but Vite handles this efficiently.

## Next Steps

1. Review and approve this plan
2. Start with Phase 1 (Setup)
3. Proceed incrementally through each phase
4. Test thoroughly after each major migration step
5. Commit frequently with clear messages

## Notes

- Keep JavaScript files until TypeScript migration is complete and tested
- Use `allowJs: true` in tsconfig to support gradual migration
- Consider using `// @ts-check` comments in remaining JS files for basic type checking
- Document any custom type definitions created

