import * as React from "react";
import {
  CommandButton,
  ContextualMenu,
  IContextualMenuItem,
  IStackStyles,
  IStackTokens,
  Link,
  SearchBox,
  Stack,
  Text,
} from "@fluentui/react";
import { LocalStrings } from "../consts/LocalStrings";
import { FileAction } from "../Interfaces";
import { Translate } from "../utils";

// The top of the control: a link to the folder in SharePoint, the search box
// and the actions for the selected files. On narrow layouts the actions
// collapse into a single "Actions" menu.

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
  selectedFiles: string[];
  /** The folder being shown, in SharePoint; no link is shown when empty. */
  sharePointUrl: string;
  isCollapsed: boolean;
  formIsDisabled: boolean;
  /** False until the record has a SharePoint location to create folders in. */
  canCreateFolder: boolean;
  menuVisible: boolean;
  menuTarget: HTMLElement | null;
  onAction: (action: FileAction) => void;
  onSearch: (value?: string) => void;
  onToggleMenu: (event: React.MouseEvent<HTMLElement>) => void;
  onCloseMenu: () => void;
  onCreateFolder: () => void;
}

/** "View files in SharePoint", with "SharePoint" linking to the folder. */
export const SharePointLink: React.FC<{ t: Translate; url: string }> = ({ t, url }) => {
  // The text is "View files in {0}", so translations can place the link.
  const [before, after = ""] = t(LocalStrings.Link.View_In_SharePoint).split("{0}");
  return (
    <Text className="sharepoint-link" styles={{ root: { fontSize: "12pt" } }}>
      {before}
      <Link href={url} target="_blank" rel="noopener noreferrer">
        {t(LocalStrings.Link.SharePoint)}
      </Link>
      {after}
    </Text>
  );
};

export const Ribbon: React.FC<RibbonProps> = ({
  t,
  selectedFiles,
  sharePointUrl,
  isCollapsed,
  formIsDisabled,
  canCreateFolder,
  menuVisible,
  menuTarget,
  onAction,
  onSearch,
  onToggleMenu,
  onCloseMenu,
  onCreateFolder,
}) => {
  const nothingSelected = selectedFiles.length === 0;

  const commandButtons = (
    <Stack horizontal tokens={ribbonStackTokens}>
      <CommandButton
        iconProps={{ iconName: "OpenInNewTab" }}
        text={t(LocalStrings.Button.Label_Open)}
        onClick={() => onAction("open")}
        disabled={nothingSelected}
        className="icon-button"
      />
      <CommandButton
        iconProps={{ iconName: "Delete" }}
        text={t(LocalStrings.Button.Label_Delete)}
        onClick={() => onAction("delete")}
        disabled={nothingSelected || formIsDisabled}
        className="icon-button"
      />
    </Stack>
  );

  const menuItems: IContextualMenuItem[] = [
    {
      key: "open",
      text: t(LocalStrings.Button.Label_Open),
      iconProps: { iconName: "OpenInNewTab" },
      onClick: () => onAction("open"),
      disabled: nothingSelected,
    },
    {
      key: "delete",
      text: t(LocalStrings.Button.Label_Delete),
      iconProps: { iconName: "Delete" },
      onClick: () => onAction("delete"),
      disabled: nothingSelected || formIsDisabled,
    },
  ];

  return (
    <Stack className="wideRibbon">
      {sharePointUrl && <SharePointLink t={t} url={sharePointUrl} />}
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
          {!nothingSelected && (
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
          {nothingSelected && canCreateFolder && (
            <CommandButton
              iconProps={{ iconName: "folder" }}
              text={t(LocalStrings.Button.Label_CreateFolder)}
              onClick={onCreateFolder}
              className="icon-button"
            />
          )}
        </Stack>
      </Stack>
    </Stack>
  );
};
