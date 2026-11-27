import * as React from "react";
import { Icon, IconButton, TooltipHost } from "@fluentui/react";
import { LocalStrings } from "../consts/LocalStrings";
import { FileData, SharePointDocument } from "../Interfaces";
import { Translate } from "../utils";
import FolderIcon from "./FolderIcon";
import { FILE_ICON_SIZE, getFileTypeIconProps } from "./fileTypeIcons";
import { formatFileSize, getFileExtension, middleEllipsis } from "./landingHelpers";

// The tiles shown inside the drop zone: one grid for Notes and one for the
// current SharePoint folder.

const fileIcon = (filename: string) => (
  <Icon
    {...getFileTypeIconProps({
      extension: getFileExtension(filename),
      size: FILE_ICON_SIZE,
      imageFileType: "svg",
    })}
  />
);

/** Stops the click from reaching the drop zone, which would open the file picker. */
const swallow = (event: React.MouseEvent) => {
  event.preventDefault();
  event.stopPropagation();
};

export interface NoteFileGridProps {
  files: FileData[];
  selectedFiles: string[];
  onTileClick: (file: FileData) => void;
  onTileContextMenu: (event: React.MouseEvent<HTMLElement>, file: FileData) => void;
}

export const NoteFileGrid: React.FC<NoteFileGridProps> = ({
  files,
  selectedFiles,
  onTileClick,
  onTileContextMenu,
}) => (
  <div style={{ display: "flex", flexWrap: "wrap", flexGrow: 1 }}>
    {files.map((file) => (
      <div
        key={file.noteId}
        className={`file-box ${selectedFiles.includes(file.noteId || "") ? "selected" : ""}`}
        onClick={(event) => {
          swallow(event);
          onTileClick(file);
        }}
        onContextMenu={(event) => onTileContextMenu(event, file)}
      >
        {fileIcon(file.filename)}
        <TooltipHost content={file.filename}>
          <p className="file-name">{middleEllipsis(file.filename)}</p>
        </TooltipHost>
        <p className="file-size">{formatFileSize(file.filesize)}</p>
      </div>
    ))}
  </div>
);

export interface SharePointFileGridProps {
  t: Translate;
  items: SharePointDocument[];
  selectedFiles: string[];
  formIsDisabled: boolean;
  onTileClick: (item: SharePointDocument) => void;
  onTileContextMenu: (event: React.MouseEvent<HTMLElement>, item: SharePointDocument) => void;
  onRemoveFolderClick: (event: React.MouseEvent<unknown>, folder: SharePointDocument) => void;
}

export const SharePointFileGrid: React.FC<SharePointFileGridProps> = ({
  t,
  items,
  selectedFiles,
  formIsDisabled,
  onTileClick,
  onTileContextMenu,
  onRemoveFolderClick,
}) => (
  <div style={{ display: "flex", alignItems: "center" }}>
    <div style={{ display: "flex", flexWrap: "wrap", flexGrow: 1 }}>
      {items.map((item) => {
        const isFolder = item.filetype === "folder";
        return (
          <div
            key={item.sharepointdocumentid}
            className={`file-box ${selectedFiles.includes(item.sharepointdocumentid) ? "selected" : ""}`}
            onClick={(event) => {
              swallow(event);
              onTileClick(item);
            }}
            onContextMenu={(event) => onTileContextMenu(event, item)}
          >
            <div className="file-image" style={{ position: "relative" }}>
              {isFolder ? (
                <>
                  <FolderIcon />
                  {!formIsDisabled && (
                    <IconButton
                      className="remove-button"
                      iconProps={{ iconName: "Cancel" }}
                      title={t(LocalStrings.Button.Label_Remove)}
                      ariaLabel={t(LocalStrings.Button.Label_Remove)}
                      onClick={(event) => onRemoveFolderClick(event, item)}
                    />
                  )}
                </>
              ) : (
                fileIcon(item.fullname)
              )}
            </div>
            <TooltipHost content={item.fullname}>
              <p className="file-name">{middleEllipsis(item.fullname, 18, true)}</p>
            </TooltipHost>
            <p className="file-size">{isFolder ? " " : formatFileSize(item.filesize)}</p>
          </div>
        );
      })}
    </div>
  </div>
);
