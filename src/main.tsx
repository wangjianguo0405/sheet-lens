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

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
