import { execSync } from "child_process";

try {
  console.log("Restoring src/data files using git...");
  const output = execSync("git checkout -- src/data/vsic.ts src/data/vsicCsvData.ts", { encoding: "utf8" });
  console.log("Git checkout successful:", output);
} catch (error: any) {
  console.error("Failed to restore files via git:", error.message || error);
  try {
    console.log("Trying git status...");
    const status = execSync("git status", { encoding: "utf8" });
    console.log("Status:\n", status);
  } catch (err) {
    console.error("Git status failed:", err);
  }
}
