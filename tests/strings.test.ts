import * as fs from "fs";
import * as path from "path";
import { LocalStrings } from "../Dropzone/consts/LocalStrings";

// Every string the code looks up must exist in each language file, or users
// see the raw key instead of text.

function keys(value: unknown): string[] {
  return typeof value === "string" ? [value] : Object.values(value as object).flatMap(keys);
}

const stringsDir = path.join(__dirname, "..", "Dropzone", "strings");

describe.each(fs.readdirSync(stringsDir).filter((f) => f.endsWith(".resx")))("%s", (file) => {
  const resx = fs.readFileSync(path.join(stringsDir, file), "utf8");
  const defined = new Set([...resx.matchAll(/<data name="([^"]+)"/g)].map((m) => m[1]));

  it("defines every key in LocalStrings", () => {
    expect(keys(LocalStrings).filter((key) => !defined.has(key))).toEqual([]);
  });
});
