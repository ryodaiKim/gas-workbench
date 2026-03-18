# Visit Management Workspace

npm workspaces monorepo for Google Apps Script projects, developed locally with TypeScript and pushed via clasp.

## Structure

```
projects/
└── reminder-system/   Automated reminder emails & alerting from Google Sheets
```

## Getting Started

```sh
npm install          # install dependencies (hoisted to root)
npm run build        # build all projects
npm run push         # build & push all projects to Apps Script
```

## Working on a Single Project

```sh
cd projects/reminder-system
npm run build        # build this project only
npm run push         # build & push this project only
npm run watch        # watch mode (recompile on save)
npm run deploy       # build, push, and list deployments
```

## Adding a New Project

1. Create `projects/<name>/` with `src/`, `scripts/`, `docs/`
2. Add `package.json` with project-specific scripts (no devDependencies needed — they're hoisted)
3. Add `tsconfig.json` extending `../../tsconfig.json` with project-specific `outDir`/`rootDir`
4. Add `.clasp.json` with the target script ID
5. Add `.claspignore`
6. Run `npm install` from root to register the new workspace

## Shared Configuration

| File                 | Purpose                                         |
| -------------------- | ----------------------------------------------- |
| Root `package.json`  | Workspace config, shared devDependencies        |
| Root `tsconfig.json` | Base TypeScript options (target, module, types) |

Each project extends the root tsconfig and adds its own `outDir`, `rootDir`, and `include`.
