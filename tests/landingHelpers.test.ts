import {
  currentFolderUrl,
  filterAndSortSharePointItems,
  formatFileSize,
  getFileExtension,
  middleEllipsis,
} from "../Dropzone/Landing/landingHelpers";
import { SharePointDocument } from "../Dropzone/Interfaces";

describe("formatFileSize", () => {
  it.each([
    [0, "0 Bytes"],
    [1023, "1023 Bytes"],
    [1024, "1.00 KB"],
    [1536, "1.50 KB"],
    [1048576, "1.00 MB"],
    [3 * 1048576 + 524288, "3.50 MB"],
  ])("%d bytes -> %s", (bytes, expected) => {
    expect(formatFileSize(bytes)).toBe(expected);
  });
});

describe("middleEllipsis", () => {
  it("leaves short names alone", () => {
    expect(middleEllipsis("photo.png")).toBe("photo.png");
  });

  it("keeps the start, the last characters and the extension of a long file name", () => {
    expect(middleEllipsis("Quarterly report final version.pdf")).toBe("Quarterly...ion.pdf");
  });

  it("cuts folders and names without an extension at the end", () => {
    expect(middleEllipsis("Invoices & Receipts", 18, true)).toBe("Invoices & Rece...");
    expect(middleEllipsis("a-very-long-name-without-dot")).toBe("a-very-long-nam...");
  });
});

describe("getFileExtension", () => {
  it.each([
    ["Report.PDF", "pdf"],
    ["archive.tar.gz", "gz"],
    ["", "folder"],
    [undefined, "folder"],
    ["trailingdot.", "txt"],
  ])("%s -> %s", (name, expected) => {
    expect(getFileExtension(name)).toBe(expected);
  });
});

const spItem = (fullname: string, filetype: string) =>
  ({ fullname, filetype, sharepointdocumentid: fullname }) as SharePointDocument;

describe("filterAndSortSharePointItems", () => {
  it("puts folders first, then sorts by name", () => {
    const items = [
      spItem("b.docx", "docx"),
      spItem("Zeta", "folder"),
      spItem("A.pdf", "pdf"),
      spItem("alpha", "folder"),
    ];
    expect(filterAndSortSharePointItems(items, "").map((i) => i.fullname)).toEqual([
      "alpha",
      "Zeta",
      "A.pdf",
      "b.docx",
    ]);
  });

  it("filters by name, ignoring case and surrounding spaces", () => {
    const items = [spItem("Contract.pdf", "pdf"), spItem("Invoice.pdf", "pdf")];
    expect(filterAndSortSharePointItems(items, "  contract ").map((i) => i.fullname)).toEqual([
      "Contract.pdf",
    ]);
  });

  it("doesn't change the list it's given", () => {
    const items = [spItem("b", "pdf"), spItem("a", "pdf")];
    filterAndSortSharePointItems(items, "");
    expect(items.map((i) => i.fullname)).toEqual(["b", "a"]);
  });
});

describe("currentFolderUrl", () => {
  const location = "https://contoso.sharepoint.com/sites/crm/account/Contoso_1234";
  const withUrl = (absoluteurl: string) => ({ absoluteurl }) as SharePointDocument;

  it("uses the parent of an item's URL when the folder has items", () => {
    const items = [withUrl(`${location}/Invoices%20%26%20Receipts/march.pdf`)];
    expect(currentFolderUrl("", "Invoices & Receipts", items)).toBe(
      `${location}/Invoices%20%26%20Receipts`,
    );
    // A folder item's parent is the folder being shown too.
    expect(currentFolderUrl(location, "", [withUrl(`${location}/Contracts`)])).toBe(location);
  });

  it("uses the location URL for an empty top folder", () => {
    expect(currentFolderUrl(`${location}/`, "", [])).toBe(location);
  });

  it("adds the subfolder path, encoded, for an empty subfolder", () => {
    expect(currentFolderUrl(location, "Invoices & Receipts/2024", [])).toBe(
      `${location}/Invoices%20%26%20Receipts/2024`,
    );
  });

  it('returns "" when nothing is known', () => {
    expect(currentFolderUrl("", "", [])).toBe("");
    expect(currentFolderUrl("", "", [withUrl("")])).toBe("");
  });
});
