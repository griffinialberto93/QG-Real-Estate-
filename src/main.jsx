import React from "react";
import ReactDOM from "react-dom/client";
import AuthGate from "./AuthGate";
import App from "./App";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <AuthGate>
      <App />
    </AuthGate>
  </React.StrictMode>
);
