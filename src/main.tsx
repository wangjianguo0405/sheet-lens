import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { shortDatePattern, tellShortDate } from "./lib/excel";
import "./lib/i18n";
import "./index.css";

// Before any workbook: the built-in short-date formats are drawn in the
// reader's own pattern, and the parser is told it once — see `shortDatePattern`.
void tellShortDate(shortDatePattern()).catch(() => {
  // A parser that could not be told keeps its own default, which is the
  // American pattern every version before this call drew.
});

// The webview's own right-click menu — Back, Reload, Print — belongs to a
// browser, not to a document viewer: every menu the app shows is its own, and
// a surface without one should show nothing. Text fields keep theirs, since
// that is where cut, copy and paste live.
document.addEventListener("contextmenu", (event) => {
  const target = event.target instanceof Element ? event.target : null;
  if (target?.closest("input, textarea, [contenteditable='true']")) return;
  event.preventDefault();
});

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
