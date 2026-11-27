import {
  addRegardingFilter,
  filterAndSortSharePointItems,
  filterNotes,
  formatFileSize,
  getFileExtension,
  middleEllipsis,
} from "../Dropzone/Landing/landingHelpers";
import { FileData, SharePointDocument } from "../Dropzone/Interfaces";

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

describe("addRegardingFilter", () => {
  const id = "11111111-2222-3333-4444-555555555555";
  const condition = `<condition attribute="objectid" operator="eq" value="${id}"/>`;

  it("adds the condition to the view's first filter", () => {
    const fetch =
      '<fetch><entity name="annotation"><filter type="and"><condition attribute="ownerid" operator="eq-userid"/></filter></entity></fetch>';
    expect(addRegardingFilter(fetch, id)).toBe(
      `<fetch><entity name="annotation"><filter type="and">${condition}<condition attribute="ownerid" operator="eq-userid"/></filter></entity></fetch>`,
    );
  });

  it("adds a filter when the view has none", () => {
    const fetch = '<fetch><entity name="annotation"><attribute name="filename"/></entity></fetch>';
    expect(addRegardingFilter(fetch, id)).toBe(
      `<fetch><entity name="annotation"><attribute name="filename"/><filter type="and">${condition}</filter></entity></fetch>`,
    );
  });

  it("escapes the record id", () => {
    expect(addRegardingFilter("<fetch><entity></entity></fetch>", 'a"b')).toContain(
      'value="a&quot;b"',
    );
  });
});

const note = (filename: string): FileData => ({ filename, filesize: 1, createdon: new Date(0) });

describe("filterNotes", () => {
  it("matches the file name case-insensitively", () => {
    const notes = [note("Budget.xlsx"), note("photo.png"), note("budget-2.pdf")];
    expect(filterNotes(notes, "BUDGET").map((n) => n.filename)).toEqual([
      "Budget.xlsx",
      "budget-2.pdf",
    ]);
  });

  it("returns everything for an empty search", () => {
    expect(filterNotes([note("a"), note("b")], "")).toHaveLength(2);
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
