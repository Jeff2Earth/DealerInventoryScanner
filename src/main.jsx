import React from "react";
import ReactDOM from "react-dom/client";
import LotLedger from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <LotLedger />
  </React.StrictMode>
);

// Auto-update: whenever the app comes back to the foreground, check whether a
// newer build was deployed. If so, reload (filters/scroll are saved in
// localStorage, so the user lands back where they were).
if (typeof __BUILD_ID__ !== "undefined") {
  let checking = false;
  const checkForUpdate = async () => {
    if (checking || document.visibilityState !== "visible") return;
    checking = true;
    try {
      const res = await fetch(`${import.meta.env.BASE_URL}version.json?t=${Date.now()}`, { cache: "no-store" });
      if (res.ok) {
        const { build } = await res.json();
        if (build && build !== __BUILD_ID__) window.location.reload();
      }
    } catch (e) { /* offline: ignore */ }
    checking = false;
  };
  document.addEventListener("visibilitychange", checkForUpdate);
  window.addEventListener("focus", checkForUpdate);
  checkForUpdate();
}
