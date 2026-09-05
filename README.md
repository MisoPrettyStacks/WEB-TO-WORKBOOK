# Weave — Webpage to Workbook

Weave turns public webpages or pasted webpage text into clean, editable Excel workbooks. It detects tables, lists, structured data, article cards, and page metadata, then lets you refine the result before downloading an `.xlsx` file.

## Run locally

Requirements: Node.js 20+ and pnpm 9+.

```bash
pnpm install
```

Start the API in one terminal:

```bash
pnpm --filter @workspace/api-server run dev
```

Start the web app in a second terminal:

```bash
PORT=21919 BASE_PATH=/ pnpm --filter @workspace/weave-workbook run dev
```

Then open the Vite URL shown in the terminal.

## GitHub

GitHub stores the code but does not run the application by itself. You can upload this source to a repository and run it in GitHub Codespaces, a Node hosting provider, or Replit. The app uses a small Express API for webpage fetching and a Vite React frontend for the workbook editor.

## Features

- Extract from a public URL or pasted page text
- Fall back gracefully when a website blocks automated access
- Edit sheet names, headers, cells, and rows
- Add another source as a new workbook tab
- Save locally in the browser
- Download a real `.xlsx` workbook