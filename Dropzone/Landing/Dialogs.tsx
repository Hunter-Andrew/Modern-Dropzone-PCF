import * as React from "react";
import {
  ComboBox,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IComboBox,
  IComboBoxOption,
  Label,
  PrimaryButton,
  Stack,
  TextField,
} from "@fluentui/react";
import { LocalStrings } from "../consts/LocalStrings";
import { Translate } from "../utils";

// Dialogs used by Landing. Each one is controlled: Landing
// owns the state and passes values and callbacks in.

type TextChangeHandler = (
  event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
  newValue?: string,
) => void;

export interface CreateFolderDialogProps {
  t: Translate;
  hidden: boolean;
  folderName: string;
  onFolderNameChange: TextChangeHandler;
  onCreate: () => void;
  onCancel: () => void;
}

export const CreateFolderDialog: React.FC<CreateFolderDialogProps> = ({
  t,
  hidden,
  folderName,
  onFolderNameChange,
  onCreate,
  onCancel,
}) => (
  <Dialog
    hidden={hidden}
    onDismiss={onCancel}
    dialogContentProps={{
      type: DialogType.normal,
      title: t(LocalStrings.Dialog.CreateFolder.Title),
      subText: t(LocalStrings.Dialog.CreateFolder.Description),
    }}
    modalProps={{
      isBlocking: false,
      styles: { main: { maxWidth: 450 } },
    }}
  >
    <TextField
      value={folderName}
      onChange={onFolderNameChange}
      placeholder={t(LocalStrings.Dialog.CreateFolder.Input_Placeholder)}
    />
    <DialogFooter>
      <PrimaryButton
        onClick={onCreate}
        disabled={!folderName.trim()}
        text={t(LocalStrings.Dialog.CreateFolder.Button_Create)}
      />
      <DefaultButton onClick={onCancel} text={t(LocalStrings.Dialog.CreateFolder.Button_Cancel)} />
    </DialogFooter>
  </Dialog>
);

export interface DeleteFolderDialogProps {
  t: Translate;
  hidden: boolean;
  folderName?: string;
  onConfirm: () => void;
  onCancel: () => void;
}

export const DeleteFolderDialog: React.FC<DeleteFolderDialogProps> = ({
  t,
  hidden,
  folderName,
  onConfirm,
  onCancel,
}) => (
  <Dialog
    hidden={hidden}
    onDismiss={onCancel}
    dialogContentProps={{
      type: DialogType.normal,
      title: t(LocalStrings.Button.Label_Remove_Folder),
      subText: `${t(LocalStrings.Dialog.DeleteFolder.Confirmation)} "${folderName}"?`,
    }}
  >
    <DialogFooter>
      <PrimaryButton onClick={onConfirm} text={t(LocalStrings.Dialog.DeleteFolder.Button_Yes)} />
      <DefaultButton onClick={onCancel} text={t(LocalStrings.Dialog.DeleteFolder.Button_No)} />
    </DialogFooter>
  </Dialog>
);

export interface AddLocationDialogProps {
  t: Translate;
  hidden: boolean;
  displayName: string;
  folderName: string;
  selectedParentSite: string;
  parentSiteOptions: IComboBoxOption[];
  canSave: boolean;
  /** Receives changes for the inputs named "createLocationDisplayName" and "createLocationFolderName". */
  onInputChange: TextChangeHandler;
  onParentSiteChange: (event: React.FormEvent<IComboBox>, option?: IComboBoxOption) => void;
  onSave: () => void;
  onCancel: () => void;
}

export const AddLocationDialog: React.FC<AddLocationDialogProps> = ({
  t,
  hidden,
  displayName,
  folderName,
  selectedParentSite,
  parentSiteOptions,
  canSave,
  onInputChange,
  onParentSiteChange,
  onSave,
  onCancel,
}) => (
  <Dialog
    hidden={hidden}
    onDismiss={onCancel}
    dialogContentProps={{
      type: DialogType.normal,
      title: t(LocalStrings.Dialog.AddLocation.Title),
    }}
  >
    <Stack tokens={{ childrenGap: 20 }} padding={20}>
      <Label>{t(LocalStrings.Dialog.AddLocation.Description)}</Label>
      <TextField
        label={t(LocalStrings.Dialog.AddLocation.Input_Label_DisplayName)}
        required
        name="createLocationDisplayName"
        value={displayName}
        onChange={onInputChange}
      />
      <ComboBox
        label={t(LocalStrings.Dialog.AddLocation.Input_Label_ParentSite)}
        required
        selectedKey={selectedParentSite}
        options={parentSiteOptions}
        onChange={onParentSiteChange}
      />
      <TextField
        label={t(LocalStrings.Dialog.AddLocation.Input_Label_FolderName)}
        required
        name="createLocationFolderName"
        value={folderName}
        onChange={onInputChange}
      />
    </Stack>
    <DialogFooter>
      <PrimaryButton
        onClick={onSave}
        text={t(LocalStrings.Dialog.AddLocation.Button_Save)}
        disabled={!canSave}
      />
      <DefaultButton onClick={onCancel} text={t(LocalStrings.Dialog.AddLocation.Button_Cancel)} />
    </DialogFooter>
  </Dialog>
);
