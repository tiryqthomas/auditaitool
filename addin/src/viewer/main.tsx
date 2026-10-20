import { createRoot } from "react-dom/client";
import { Viewer } from "./Viewer";
import "../styles.css";

const params = new URLSearchParams(location.search);

Office.onReady(() => {
  createRoot(document.getElementById("root")!).render(<Viewer params={params} />);
});
