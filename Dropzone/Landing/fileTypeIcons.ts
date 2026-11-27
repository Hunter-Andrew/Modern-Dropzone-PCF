import * as React from "react";
import { registerIcons } from "@fluentui/react";
import { FileTypeIconMap, getFileTypeIconProps } from "@fluentui/react-file-type-icons";

// Registers the Office file-type icons with the platform-provided Fluent UI.
//
// The package's own initializeFileTypeIcons() calls registerIcons from
// @fluentui/style-utilities, which is not supplied by the platform, so using it
// would bundle a second copy of Fluent's styling code. Registering through
// @fluentui/react keeps that code external. Only the size and format the
// control renders (96px SVG, plus the 1.5x variant used on 150% displays) are
// registered, instead of every size and format.

export { getFileTypeIconProps };

export const FILE_ICON_SIZE = 96;

const BASE_URL = "https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/";

let registered = false;

export function initializeFileTypeIcons(): void {
  if (registered) return;
  registered = true;

  const icons: Record<string, React.ReactElement> = {};
  for (const type of Object.keys(FileTypeIconMap)) {
    icons[`${type}${FILE_ICON_SIZE}_svg`] = React.createElement("img", {
      src: `${BASE_URL}${FILE_ICON_SIZE}/${type}.svg`,
      alt: "",
    });
    icons[`${type}${FILE_ICON_SIZE}_1.5x_svg`] = React.createElement("img", {
      src: `${BASE_URL}${FILE_ICON_SIZE}_1.5x/${type}.svg`,
      height: "100%",
      width: "100%",
      alt: "",
    });
  }

  registerIcons({
    fontFace: {},
    style: { width: FILE_ICON_SIZE, height: FILE_ICON_SIZE, overflow: "hidden" },
    icons,
    mergeImageProps: true,
  });
}
