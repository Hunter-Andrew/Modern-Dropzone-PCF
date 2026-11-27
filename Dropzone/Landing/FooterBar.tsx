import * as React from "react";
import {
  Callout,
  DirectionalHint,
  Dropdown,
  HighContrastSelector,
  IButtonStyles,
  IContextualMenuItem,
  IContextualMenuProps,
  IDropdownOption,
  IDropdownStyles,
  IconButton,
  Label,
  Stack,
  Toggle,
} from "@fluentui/react";
import { LocalStrings } from "../consts/LocalStrings";
import { Translate } from "../utils";

// The bar under the drop zone: the SharePoint location picker, with a button
// to add a location and a settings menu ("Remember Location").

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

  isSettingsCalloutVisible: boolean;
  settingsCalloutTarget: HTMLElement | null;
  onSettingsCalloutDismiss: () => void;
  rememberLocation: boolean | undefined;
  onRememberLocationChange: (event: React.MouseEvent<HTMLElement>, checked?: boolean) => void;

  locationOptions: IDropdownOption[];
  selectedLocation: string | null;
  onLocationChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  canAddLocation: boolean;
  /** Wraps the add-location button; the settings callout is anchored to it. */
  addLocationRef: React.RefObject<HTMLDivElement>;
  onAddLocation: () => void;
  onSettingsClick: (
    ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>,
    item?: IContextualMenuItem,
  ) => boolean | void;
}

export const FooterBar: React.FC<FooterBarProps> = (props) => {
  const { t } = props;

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
      horizontalAlign="end"
      verticalAlign="center"
      wrap
      tokens={{ childrenGap: 5 }}
      style={{ width: "100%", marginTop: "8px" }}
    >
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
              <Label>{t(LocalStrings.Button.Label_RememberLocation)}</Label>
              <Toggle
                onChange={props.onRememberLocationChange}
                styles={{ root: { marginBottom: "0px" } }}
                checked={props.rememberLocation}
              />
            </Stack>
          </Stack>
        </Callout>
      )}
      <Dropdown
        options={props.locationOptions}
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
          />
        </div>
      )}
    </Stack>
  );
};
