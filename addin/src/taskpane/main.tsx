import { createRoot } from "react-dom/client";
import { App } from "./App";
import "../styles.css";

Office.onReady(() => {
  createRoot(document.getElementById("root")!).render(<App />);
});
