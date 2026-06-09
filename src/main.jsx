import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

document.body.style.margin = "0";
document.body.style.background = "#0f172a";
document.body.style.color = "#f8fafc";
document.body.style.fontFamily =
  "Inter, system-ui, -apple-system, BlinkMacSystemFont, sans-serif";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);