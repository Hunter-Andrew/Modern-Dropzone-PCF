import * as React from "react";
import {
  CommandButton,
  ContextualMenu,
  IContextualMenuItem,
  IStackStyles,
  IStackTokens,
  SearchBox,
  Stack,
} from "@fluentui/react";
import { LocalStrings } from "../consts/LocalStrings";
import { FileAction, FileData } from "../Interfaces";
import { isImage, isPDF, Translate } from "../utils";

// Search box and the actions for the selected files. On narrow layouts the
// actions collapse into a single "Actions" menu.

const ribbonStackTokens: IStackTokens = { childrenGap: 10 };

const ribbonStyles: IStackStyles = {
  root: {
    alignItems: "center",
    display: "flex",
    width: "100%",
    justifyContent: "space-between",
  },
};

const searchRibbonStyles: IStackStyles = {
  root: {
    marginBottom: "0px",
  },
};

const searchBoxStyles = {
  root: {
    width: 200,
    border: "none",
    boxShadow: "none",
    position: "relative" as const,
    selectors: {
      ":hover": {
        border: "none",
        boxShadow: "none",
      },
      ":focus": {
        border: "none",
        boxShadow: "none",
      },
      "::after": {
        content: "none !important",
      },
    },
  },
  field: {
    border: "none",
    boxShadow: "none",
  },
  iconContainer: {
    color: "rgb(17, 94, 163)",
    selectors: {
      ":hover": {
        color: "rgb(17, 94, 163)",
      },
      ":focus": {
        color: "rgb(17, 94, 163)",
      },
    },
  },
};

export interface RibbonProps {
  t: Translate;
  /** All notes (used to check whether the selection can be previewed). */
  files: FileData[];
  selectedFiles: string[];
  sharePointDocLoc: boolean;
  isCollapsed: boolean;
  formIsDisabled: boolean;
  menuVisible: boolean;
  menuTarget: HTMLElement | null;
  onAction: (action: FileAction) => void;
  onSearch: (value?: string) => void;
  onToggleMenu: (event: React.MouseEvent<HTMLElement>) => void;
  onCloseMenu: () => void;
  onCreateFolder: () => void;
}

export const Ribbon: React.FC<RibbonProps> = ({
  t,
  files,
  selectedFiles,
  sharePointDocLoc,
  isCollapsed,
  formIsDisabled,
  menuVisible,
  menuTarget,
  onAction,
  onSearch,
  onToggleMenu,
  onCloseMenu,
  onCreateFolder,
}) => {
  // Preview (and the wide-layout Rename button) need exactly one selected
  // image or PDF.
  const cannotPreview =
    selectedFiles.length !== 1 ||
    !files.some(
      (file) =>
        selectedFiles.includes(file.noteId!) &&
        (isImage(file.mimetype!) || isPDF(file.mimetype!))
    );

  const commandButtons = (
    <Stack horizontal tokens={ribbonStackTokens}>
      {!sharePointDocLoc && selectedFiles.length < 2 && (
        <>
          <CommandButton
            text="Preview"
            iconProps={{ iconName: "View" }}
            onClick={() => onAction("preview")}
            disabled={cannotPreview}
            className="icon-button"
          />
          <CommandButton
            iconProps={{ iconName: "Edit" }}
            text="Rename"
            onClick={() => onAction("edit")}
            disabled={cannotPreview}
            className="icon-button"
          />
        </>
      )}
      <CommandButton
        iconProps={{ iconName: "Download" }}
        text={t(LocalStrings.Button.Label_Download)}
        onClick={() => onAction("download")}
        disabled={selectedFiles.length === 0}
        className="icon-button"
      />
      <CommandButton
        iconProps={{ iconName: "Delete" }}
        text={t(LocalStrings.Button.Label_Delete)}
        onClick={() => onAction("delete")}
        disabled={selectedFiles.length === 0 || formIsDisabled}
        className="icon-button"
      />
    </Stack>
  );

  const menuItems: IContextualMenuItem[] = [
    {
      key: "download",
      text: t(LocalStrings.Button.Label_Download),
      iconProps: { iconName: "Download" },
      onClick: () => onAction("download"),
      disabled: selectedFiles.length === 0,
    },
    {
      key: "delete",
      text: t(LocalStrings.Button.Label_Delete),
      iconProps: { iconName: "Delete" },
      onClick: () => onAction("delete"),
      disabled: selectedFiles.length === 0 || formIsDisabled,
    },
  ];

  if (!sharePointDocLoc) {
    menuItems.unshift(
      {
        key: "preview",
        text: t(LocalStrings.Button.Label_Preview),
        iconProps: { iconName: "View" },
        onClick: () => onAction("preview"),
        disabled: cannotPreview,
      },
      {
        key: "rename",
        text: t(LocalStrings.Button.Label_Rename),
        iconProps: { iconName: "Edit" },
        onClick: () => onAction("edit"),
        disabled:
          selectedFiles.length === 0 ||
          selectedFiles.length >= 2 ||
          formIsDisabled,
      }
    );
  }

  return (
    <Stack horizontal styles={searchRibbonStyles} className="easeIn wideRibbon">
      <Stack horizontal styles={ribbonStyles} className="easeIn">
        <Stack horizontal tokens={ribbonStackTokens}>
          <SearchBox
            placeholder={t(LocalStrings.Input.Placeholder_Search)}
            onSearch={(newValue) => onSearch(newValue)}
            onChange={(_, newValue) => onSearch(newValue)}
            styles={searchBoxStyles}
          />
        </Stack>
        {selectedFiles.length > 0 && (
          <Stack horizontal tokens={ribbonStackTokens}>
            {menuVisible && (
              <ContextualMenu
                items={menuItems}
                target={menuTarget}
                onDismiss={onCloseMenu}
                isBeakVisible={true}
              />
            )}
            {isCollapsed ? (
              <CommandButton
                text={t(LocalStrings.Button.Label_Actions)}
                onClick={onToggleMenu}
                className="icon-button"
              />
            ) : (
              commandButtons
            )}
          </Stack>
        )}
        {selectedFiles.length < 1 && sharePointDocLoc && !formIsDisabled && (
          <CommandButton
            iconProps={{ iconName: "folder" }}
            text={t(LocalStrings.Button.Label_CreateFolder)}
            onClick={onCreateFolder}
            className="icon-button"
          />
        )}
      </Stack>
    </Stack>
  );
};
