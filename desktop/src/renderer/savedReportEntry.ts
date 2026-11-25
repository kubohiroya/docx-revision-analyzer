import type { SavedVersionReport } from "../reportData";
import { renderSavedVersionReport } from "./savedReportView";

const data = document.getElementById("report-data");
const root = document.getElementById("report");
if (data?.textContent && root) {
  renderSavedVersionReport(root, JSON.parse(data.textContent) as SavedVersionReport);
}
