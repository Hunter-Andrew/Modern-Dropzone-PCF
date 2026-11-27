import * as fs from "fs";
import * as path from "path";

// Dataverse's solution import validates the manifest against an XSD whose
// string types (noAposStringType) reject apostrophes, so a ' in any attribute
// value fails the import even though the build succeeds.
it("the manifest has no apostrophes in attribute values", () => {
  const manifest = fs.readFileSync(
    path.join(__dirname, "..", "Dropzone", "ControlManifest.Input.xml"),
    "utf8",
  );
  const withApostrophe = [...manifest.matchAll(/([\w-]+)="([^"]*'[^"]*)"/g)].map(
    (m) => `${m[1]}="${m[2]}"`,
  );
  expect(withApostrophe).toEqual([]);
});
