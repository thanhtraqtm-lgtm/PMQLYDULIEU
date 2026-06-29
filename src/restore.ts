import { execSync } from "child_process";

try {
  console.log("Restoring src/App.tsx using git...");
  const output = execSync("git checkout -- src/App.tsx", { encoding: "utf8" });
  console.log("Git checkout successful:", output);
} catch (error: any) {
  console.error("Failed to restore files via git:", error.message || error);
}

