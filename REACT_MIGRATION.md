# DRT Connect React Migration

This workspace now includes a React/Vite application shell.

## What is migrated

- React app bootstrap
- React Router navigation
- Route-level hosting for:
  - Dashboard
  - Create DRT
  - Load DRT

## Current compatibility strategy

The existing business flows are preserved by hosting the current production pages inside React-managed routes. This keeps the application stable while creating a real React codebase and an incremental migration path.

## Run

```bash
npm install
npm run dev
```

Open:

```text
/react-index.html
```

## Next migration steps

1. Port dashboard UI and modal state from `index.html` + `script.js` into React components.
2. Extract DRT header forms into React components.
3. Replace page-level imperative DOM logic in `create-drt.js` with React state/hooks in slices.
4. Wrap Handsontable in a React adapter and move spreadsheet state into React-managed stores.
