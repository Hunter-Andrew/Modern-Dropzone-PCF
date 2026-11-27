import * as React from "react";
import {
  Callout,
  DefaultButton,
  DirectionalHint,
  Dropdown,
  HighContrastSelector,
  IButtonStyles,
  IContextualMenuItem,
  IContextualMenuProps,
  IDropdownOption,
  IDropdownStyles,
  IconButton,
  ITooltipHostStyles,
  Label,
  Stack,
  Toggle,
  TooltipHost,
} from "@fluentui/react";
import { LocalStrings } from "../consts/LocalStrings";
import { Translate } from "../utils";

// The bar under the drop zone: the Notes/SharePoint toggle, and either the
// SharePoint location picker (with "add location" and settings) or the Notes
// view picker.

const tooltipId = "sharePointDocLocTooltip";

const SharePointDocLocTooltip = (
  <span>
    Upload files directly to a SharePoint Document Location. To learn more go to{" "}
    <a
      href="https://learn.microsoft.com/en-us/power-platform/admin/create-edit-document-location-records"
      target="_blank"
      rel="noopener noreferrer"
    >
      MS Docs
    </a>
    .
  </span>
);

const infoButtonStyles: Partial<ITooltipHostStyles> = {
  root: {
    border: "none",
    minWidth: "auto",
    margin: 0,
    padding: 0,
    display: "inline-block",
  },
};

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 200 },
};

const splitButtonStyles: IButtonStyles = {
  splitButtonMenuButton: {
    backgroundColor: "white",
    width: 28,
    border: "none",
  },
  splitButtonMenuIcon: { fontSize: "7px" },
  splitButtonDivider: {
    backgroundColor: "#c8c8c8",
    width: 1,
    right: 26,
    position: "absolute",
    top: 4,
    bottom: 4,
  },
  splitButtonContainer: {
    selectors: {
      [HighContrastSelector]: { border: "none" },
    },
  },
};

export interface FooterBarProps {
  t: Translate;

  /** The "Enable SharePoint Integration?" setting on the form. */
  sharePointEnabledParameter: boolean;
  sharePointDocLoc: boolean;
  sharePointToggleDisabled: boolean;
  onToggleSharePoint: (event: React.MouseEvent<HTMLElement>, checked?: boolean) => void;
  showInfoTooltip: boolean;
  onInfoClick: () => void;

  isSettingsCalloutVisible: boolean;
  settingsCalloutTarget: HTMLElement | null;
  onSettingsCalloutDismiss: () => void;
  rememberLocation: boolean | undefined;
  onRememberLocationChange: (event: React.MouseEvent<HTMLElement>, checked?: boolean) => void;

  showLocationPicker: boolean;
  locationOptions: IDropdownOption[];
  selectedLocation: string | null;
  locationPickerEnabled: boolean;
  onLocationChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  canAddLocation: boolean;
  /** Wraps the add-location button; the settings callout is anchored to it. */
  addLocationRef: React.RefObject<HTMLDivElement>;
  onAddLocation: () => void;
  onSettingsClick: (
    ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>,
    item?: IContextualMenuItem
  ) => boolean | void;

  showViewPicker: boolean;
  viewOptions: IDropdownOption[];
  selectedViewId?: string;
  viewPickerEnabled: boolean;
  onViewChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
}

export const FooterBar: React.FC<FooterBarProps> = (props) => {
  const { t } = props;

  const infoTooltip = (
    <TooltipHost
      content={props.showInfoTooltip ? SharePointDocLocTooltip : undefined}
      id={tooltipId}
      closeDelay={500}
    >
      <DefaultButton
        aria-label="more info"
        aria-describedby={props.showInfoTooltip ? tooltipId : undefined}
        onClick={props.onInfoClick}
        styles={infoButtonStyles}
        iconProps={{ iconName: "Info" }}
      />
    </TooltipHost>
  );

  const locationMenuProps: IContextualMenuProps = {
    items: [
      {
        key: "settings",
        iconProps: { iconName: "Settings" },
        text: t(LocalStrings.Button.Label_Settings),
        onClick: props.onSettingsClick,
      },
    ],
  };

  return (
    <Stack
      horizontal
      style={{
        width: "100%",
        justifyContent: !props.sharePointEnabledParameter ? "flex-end" : "space-between",
        alignItems: "center",
        flexWrap: "wrap",
        marginTop: "8px",
      }}
    >
      {props.sharePointEnabledParameter && (
        <Toggle
          label={<div>SharePoint Documents {infoTooltip}</div>}
          inlineLabel
          onText={t(LocalStrings.Toggle.Value_On)}
          offText={t(LocalStrings.Toggle.Value_Off)}
          checked={props.sharePointDocLoc}
          onChange={props.onToggleSharePoint}
          disabled={props.sharePointToggleDisabled}
          styles={{ root: { marginBottom: "0px" } }}
        />
      )}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }}>
        {props.isSettingsCalloutVisible && (
          <Callout
            target={props.settingsCalloutTarget}
            onDismiss={props.onSettingsCalloutDismiss}
            directionalHint={DirectionalHint.bottomAutoEdge}
            setInitialFocus
            role="dialog"
          >
            <Stack tokens={{ childrenGap: 10 }} padding={20}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <Label>Remember Location?</Label>
                <Toggle
                  onChange={props.onRememberLocationChange}
                  styles={{ root: { marginBottom: "0px" } }}
                  checked={props.rememberLocation}
                />
              </Stack>
            </Stack>
          </Callout>
        )}
        {props.showLocationPicker ? (
          <>
            <Dropdown
              options={props.locationOptions}
              disabled={!props.locationPickerEnabled}
              styles={dropdownStyles}
              onChange={props.onLocationChange}
              selectedKey={props.selectedLocation}
            />
            {props.canAddLocation && (
              <div ref={props.addLocationRef}>
                <IconButton
                  split
                  iconProps={{ iconName: "Add" }}
                  splitButtonAriaLabel="Location options"
                  aria-roledescription="Split button"
                  styles={splitButtonStyles}
                  menuProps={locationMenuProps}
                  ariaLabel="New item"
                  onClick={props.onAddLocation}
                  disabled={!props.locationPickerEnabled}
                />
              </div>
            )}
          </>
        ) : (
          props.showViewPicker && (
            <Dropdown
              options={props.viewOptions}
              disabled={!props.viewPickerEnabled}
              onChange={props.onViewChange}
              selectedKey={props.selectedViewId}
              styles={dropdownStyles}
            />
          )
        )}
      </Stack>
    </Stack>
  );
};
