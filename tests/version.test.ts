import { execFileSync } from "child_process";
import * as path from "path";

// The repository's version must be the same in every file (see scripts/version.js).
it("the version is the same everywhere", () => {
  const script = path.join(__dirname, "..", "scripts", "version.js");
  expect(() => execFileSync("node", [script, "check"], { stdio: "pipe" })).not.toThrow();
});
