import React, { Component } from "react";
import Dropzone from "react-dropzone";
import { IInputs } from "../generated/ManifestTypes";
import "../css/Dropzone.css";
import {
  createSharePointDocument,
  deleteSharePointDocument,
  createSharePointFolder,
  createSharePointLocation,
} from "../DataverseActions";
import { DocumentLocation, FileAction, SharePointDocument } from "../Interfaces";
import {
  errorMessage,
  formatString,
  getLocalString,
  isValidSharePointFolderName,
  getControlValue,
  getRecordId,
  LocalizationKey,
} from "../utils";
import {
  IComboBox,
  IComboBoxOption,
  IContextualMenuItem,
  IDropdownOption,
  IconButton,
  Link,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import toast, { Toaster } from "react-hot-toast";
import { initializeFileTypeIcons } from "./fileTypeIcons";
import { Ribbon } from "./Ribbon";
import { SharePointFileGrid } from "./FileGrid";
import { FooterBar } from "./FooterBar";
import { AddLocationDialog, CreateFolderDialog, DeleteFolderDialog } from "./Dialogs";
import { currentFolderUrl, filterAndSortSharePointItems, formatFileSize } from "./landingHelpers";
import {
  delay,
  loadLocationUrl,
  loadSharePointFolder,
  loadSharePointLocations,
  readRememberLocation,
  saveUserSettings,
  shouldFilterByLocation,
  waitForSharePointLocation,
} from "./landingData";
import { isLikelyTooLarge } from "./uploadErrors";
import { LocalStrings } from "../consts/LocalStrings";

// Landing is the container for the control: it owns all state and handles user
// actions. Reading data lives in landingData; the UI pieces live in Ribbon,
// FileGrid, FooterBar and Dialogs.

export interface LandingProps {
  context?: ComponentFramework.Context<IInputs>;
  isDisabled: boolean;
}

interface LandingState {
  selectedFiles: string[];
  searchText: string;
  isLoading: boolean;
  documentLocations: DocumentLocation[];
  selectedDocumentLocation: string | null;
  selectedDocumentLocationName: string;
  /** The selected location's folder in SharePoint ("" until known). */
  locationUrl: string;
  sharePointData: SharePointDocument[];
  currentFolderPath: string;
  folderStack: string[];
  showModal: boolean;
  newFolderName: string;
  isCollapsed: boolean;
  menuVisible: boolean;
  target: HTMLElement | null;
  userPreference: boolean | undefined;
  isCalloutVisible: boolean;
  isFolderDeletionDialogVisible: boolean;
  isCreateLocationDialogVisible: boolean;
  createLocationDisplayName: string;
  selectedCreateLocation: string;
  createLocationFolderName: string;
  isSaveButtonEnabled: boolean;
  selectedFolderForDelete: SharePointDocument | null;
}

export class Landing extends Component<LandingProps, LandingState> {
  private targetRef: React.RefObject<HTMLDivElement>;
  constructor(props: LandingProps) {
    super(props);
    initializeFileTypeIcons();
    this.state = {
      selectedFiles: [],
      searchText: "",
      isLoading: true,
      documentLocations: [],
      selectedDocumentLocation: null,
      selectedDocumentLocationName: "",
      locationUrl: "",
      sharePointData: [],
      currentFolderPath: "",
      folderStack: [],
      showModal: false,
      newFolderName: "",
      isCollapsed: window.innerWidth < 767,
      menuVisible: false,
      target: null,
      userPreference: false,
      isCalloutVisible: false,
      isFolderDeletionDialogVisible: false,
      isCreateLocationDialogVisible: false,
      createLocationDisplayName: "",
      selectedCreateLocation: "",
      createLocationFolderName: "",
      isSaveButtonEnabled: false,
      selectedFolderForDelete: null,
    };
    this.removeFile = this.removeFile.bind(this);
    this.openFile = this.openFile.bind(this);
    this.getSharePointData = this.getSharePointData.bind(this);
    this.handleFolderClick = this.handleFolderClick.bind(this);
    this.handleBackClick = this.handleBackClick.bind(this);
    this.handleDropdownChange = this.handleDropdownChange.bind(this);
    this.handleResize = this.handleResize.bind(this);
    this.toggleMenu = this.toggleMenu.bind(this);
    this.closeMenu = this.closeMenu.bind(this);
    this.onGearIconClick = this.onGearIconClick.bind(this);
    this.onCalloutDismiss = this.onCalloutDismiss.bind(this);
    this.toggleRememberLocation = this.toggleRememberLocation.bind(this);
    this.handleRemoveFolderClick = this.handleRemoveFolderClick.bind(this);
    this.targetRef = React.createRef();
  }

  handleRemoveFolderClick = (event: React.MouseEvent<unknown>, file: SharePointDocument): void => {
    event.preventDefault();
    event.stopPropagation();

    this.setState({ selectedFolderForDelete: file }, () => {
      this.toggleRemoveFolderDialog();
    });
  };

  toggleRemoveFolderDialog = (): void => {
    this.setState((prevState) => ({
      isFolderDeletionDialogVisible: !prevState.isFolderDeletionDialogVisible,
    }));
  };

  private handlecreateLocation = async () => {
    const { createLocationDisplayName, createLocationFolderName, selectedCreateLocation } =
      this.state;
    const { context } = this.props;

    const createLocationPromise = createSharePointLocation(
      context!,
      createLocationDisplayName,
      createLocationFolderName,
      selectedCreateLocation,
    );

    toast.promise(createLocationPromise, {
      loading: "Creating location...",
      success: (res) => {
        this.setState({ isCreateLocationDialogVisible: false });
        const createdSPLocation: IDropdownOption = {
          key: res,
          text: createLocationDisplayName,
        };
        this.handleDropdownChange({} as React.FormEvent<HTMLDivElement>, createdSPLocation);

        this.setState((prevState) => ({
          documentLocations: [
            ...prevState.documentLocations,
            {
              name: createLocationDisplayName,
              sharepointdocumentlocationid: res,
            },
          ],
          selectedDocumentLocation: res,
        }));

        return `Location "${createLocationDisplayName}" created successfully!`;
      },
      error: "Error creating location",
    });
  };

  private handleInputChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string,
  ) => {
    const target = event.currentTarget;
    const name = target.name as keyof LandingState;
    const value = newValue ?? target.value;
    this.setState(
      (prevState) =>
        ({
          ...prevState,
          [name]: value,
        }) as unknown as LandingState,
      this.validateForm,
    );
  };
  private validateForm = (): void => {
    const { createLocationDisplayName, selectedCreateLocation, createLocationFolderName } =
      this.state;
    const isSaveButtonEnabled =
      createLocationDisplayName.trim() !== "" &&
      selectedCreateLocation.trim() !== "" &&
      createLocationFolderName.trim() !== "";
    this.setState({ isSaveButtonEnabled });
  };
  private handleCreateLocationDropdownChange = (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption,
  ): void => {
    if (option) {
      this.setState({ selectedCreateLocation: option.key as string }, this.validateForm);
    }
  };

  private openCreateLocationDialog = (): void => {
    this.setState({ isCreateLocationDialogVisible: true });
  };

  private closeCreateLocationDialog = (): void => {
    this.setState({ isCreateLocationDialogVisible: false });
  };
  private targetElement: HTMLElement | null = null;
  private onGearIconClick = (
    ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>,
    item?: IContextualMenuItem,
  ): boolean | void => {
    if (ev && ev.currentTarget instanceof HTMLElement) {
      this.targetElement = ev.currentTarget;
      this.toggleCallout();
    }
  };

  private toggleCallout = (): void => {
    this.setState((prevState) => ({
      isCalloutVisible: !prevState.isCalloutVisible,
      target: this.targetElement,
    }));
  };
  private onCalloutDismiss = (): void => {
    this.setState({ isCalloutVisible: false });
  };
  private async handleFolderClick(folderPath: string): Promise<void> {
    this.setState(
      (prevState) => ({
        folderStack: [...prevState.folderStack, prevState.currentFolderPath],
        isLoading: true,
      }),
      async () => {
        try {
          const data = await loadSharePointFolder(
            this.props.context!,
            folderPath,
            this.state.selectedDocumentLocation,
            this.state.selectedDocumentLocationName,
          );
          this.setState({
            sharePointData: data,
            currentFolderPath: folderPath,
          });
        } catch (error) {
          console.error("Error fetching folder data:", error);
        } finally {
          this.setState({ isLoading: false });
        }
      },
    );
  }

  private handleBackClick = async () => {
    this.setState(
      (prevState) => {
        const folderStack = [...prevState.folderStack];
        const previousFolderPath = folderStack.pop();
        return {
          folderStack,
          currentFolderPath: previousFolderPath || "",
          isLoading: true,
        };
      },
      async () => {
        try {
          const { currentFolderPath } = this.state;
          const data = await loadSharePointFolder(
            this.props.context!,
            currentFolderPath,
            this.state.selectedDocumentLocation,
            this.state.selectedDocumentLocationName,
          );
          this.setState({ sharePointData: data });
        } catch (error) {
          console.error("Error fetching SharePoint folder data:", error);
        } finally {
          this.setState({ isLoading: false });
        }
      },
    );
  };

  private async loadSharePointLocations(): Promise<void> {
    const response = await loadSharePointLocations(this.props.context!);

    if (response.length === 0) {
      this.setState({
        documentLocations: [],
        selectedDocumentLocation: null,
        selectedDocumentLocationName: "",
        sharePointData: [],
        currentFolderPath: "",
        folderStack: [],
      });
      return;
    }

    this.setState({
      documentLocations: response,
      selectedDocumentLocation: response[0].sharepointdocumentlocationid,
      selectedDocumentLocationName: response[0].name,
    });
  }

  /** See shouldFilterByLocation. */
  private filterByLocation(): boolean {
    return shouldFilterByLocation(
      this.state.documentLocations,
      this.state.selectedDocumentLocationName,
    );
  }

  private async getSharePointData(popFolderStack = false): Promise<void> {
    if (!this.state.selectedDocumentLocation) {
      this.setState({ sharePointData: [], isLoading: false });
      return;
    }
    this.setState(
      (prev) => {
        const folderStack = [...prev.folderStack];
        const currentFolderPath = popFolderStack
          ? folderStack.pop() || ""
          : prev.currentFolderPath || "";
        return { folderStack, currentFolderPath };
      },
      async () => {
        const { currentFolderPath } = this.state;
        const data = await loadSharePointFolder(
          this.props.context!,
          currentFolderPath,
          this.state.selectedDocumentLocation,
          this.state.selectedDocumentLocationName,
          this.filterByLocation(),
        );
        this.setState({ sharePointData: data });
      },
    );
  }

  async componentDidMount() {
    const userPreference = await readRememberLocation(this.props.context!);
    this.setState({ userPreference }, () => {
      this.loadExistingFiles().then(() => {
        this.setState({ isLoading: false });
      });
    });

    window.addEventListener("resize", this.handleResize);
  }

  componentDidUpdate(prevProps: LandingProps, prevState: LandingState) {
    const prevId = getRecordId(prevProps.context);
    const currId = getRecordId(this.props.context);

    if (!prevId && currId) {
      // The record was just saved for the first time. Dynamics creates its
      // SharePoint folder asynchronously (e.g. when a Documents grid on the
      // form loads), so look for the new document location for a short while.
      this.waitForSharePointLocation();
    }

    if (prevId !== currId) {
      this.setState(
        {
          sharePointData: [],
          currentFolderPath: "",
          folderStack: [],
          selectedFiles: [],
          isLoading: true,
        },
        () => {
          this.loadExistingFiles().then(() => this.setState({ isLoading: false }));
        },
      );
    }

    if (prevState.selectedDocumentLocation !== this.state.selectedDocumentLocation) {
      this.loadLocationUrl();
    }
  }

  componentWillUnmount() {
    this.isUnmounted = true;
    window.removeEventListener("resize", this.handleResize);
  }

  private isUnmounted = false;

  /** True when files can be added or changed: the record is saved and the control is enabled. */
  private isEditable(): boolean {
    return !!getRecordId(this.props.context) && !this.props.isDisabled;
  }

  private async waitForSharePointLocation(): Promise<void> {
    const found = await waitForSharePointLocation(this.props.context!, () => this.isUnmounted);
    if (found && !this.isUnmounted) {
      await this.loadExistingFiles();
    }
  }

  /** Looks up the selected location's SharePoint folder URL for the "View files in SharePoint" link. */
  private async loadLocationUrl(): Promise<void> {
    const locationId = this.state.selectedDocumentLocation;
    if (!locationId) {
      this.setState({ locationUrl: "" });
      return;
    }
    const url = await loadLocationUrl(this.props.context!, locationId);
    // Ignore the answer if another location was selected meanwhile.
    if (!this.isUnmounted && this.state.selectedDocumentLocation === locationId) {
      this.setState({ locationUrl: url });
    }
  }

  /** The SharePoint URL of the folder being shown, or "" if unknown. */
  private sharePointUrl(): string {
    return currentFolderUrl(
      this.state.locationUrl,
      this.state.currentFolderPath,
      this.state.sharePointData,
    );
  }

  handleResize() {
    const element = document.querySelector(".ribbon-dropzone-wrapper");
    if (element) {
      const wrapper = element.getBoundingClientRect();
      this.setState({ isCollapsed: wrapper.width < 703 });
    }
  }

  toggleMenu(event: React.MouseEvent<HTMLElement>) {
    this.setState({
      menuVisible: !this.state.menuVisible,
      target: event.currentTarget,
    });
  }

  closeMenu() {
    this.setState({ menuVisible: false });
  }

  handleDrop = async (acceptedFiles: File[]) => {
    const { context, isDisabled } = this.props;
    if (!this.isEditable()) {
      return;
    }
    const allowUploads = !isDisabled && getControlValue(context!, "allowSharePointDrops") === true;
    if (!allowUploads) {
      return;
    }

    acceptedFiles.forEach((file) => {
      const reader = new FileReader();
      reader.onload = () => {
        const binaryStr = reader.result as string;
        // An empty location id uploads to the record's default location,
        // creating it if the record doesn't have one yet.
        const locationId = this.filterByLocation() ? this.state.selectedDocumentLocation! : "";
        const uploadPromise = createSharePointDocument(
          context!,
          file.name,
          binaryStr,
          this.state.currentFolderPath,
          locationId,
        );

        toast.promise(
          uploadPromise,
          {
            loading: this.t(LocalStrings.Toast.Message_Uploading_SharePoint),
            success: () => {
              this.reloadAfterUpload();
              return `${file.name} ${this.t(LocalStrings.Toast.Message_Upload_Success_SharePoint)}`;
            },
            error: (err) => {
              this.setState({ isLoading: false });
              if (isLikelyTooLarge(err, file.size)) {
                return this.tooLargeMessage(file);
              }
              return `${this.t(LocalStrings.Toast.Message_Upload_Error_SharePoint)}: ${errorMessage(err)}`;
            },
          },
          { error: { duration: 10000 } },
        );
      };

      reader.onerror = () => {
        // The browser can't read very large files into memory.
        if (isLikelyTooLarge(reader.error, file.size)) {
          toast.error(
            formatString(
              this.t(LocalStrings.Toast.Message_Read_Error_Large),
              file.name,
              formatFileSize(file.size),
            ),
            { duration: 10000 },
          );
          return;
        }
        toast.error(`${this.t(LocalStrings.Toast.Message_Read_Error_SharePoint)} ${file.name}`);
      };

      reader.readAsDataURL(file);
    });
  };

  /** Refreshes the list after an upload; a first upload may have created the record's location. */
  private reloadAfterUpload() {
    const reload = this.state.selectedDocumentLocation
      ? this.getSharePointData(false)
      : this.loadExistingFiles();
    reload
      .then(() => this.setState({ isLoading: false }))
      .catch((err) => {
        console.error("Failed to reload files:", err);
        this.setState({ isLoading: false });
        toast.error(this.t(LocalStrings.Toast.Message_Refresh_Error_SharePoint));
      });
  }

  /** "big.zip (80 MB) is too large to upload here. Upload it in SharePoint instead." */
  private tooLargeMessage(file: File): React.ReactElement {
    const url = this.sharePointUrl();
    const [before, after = ""] = this.t(LocalStrings.Toast.Message_Upload_Too_Large_Hint).split(
      "{0}",
    );
    const sharePoint = this.t(LocalStrings.Link.SharePoint);
    return (
      <span>
        {formatString(
          this.t(LocalStrings.Toast.Message_Upload_Too_Large),
          file.name,
          formatFileSize(file.size),
        )}{" "}
        {before}
        {url ? (
          <Link href={url} target="_blank" rel="noopener noreferrer">
            {sharePoint}
          </Link>
        ) : (
          sharePoint
        )}
        {after}
      </span>
    );
  }

  loadExistingFiles = async () => {
    if (!this.props.context) return;
    this.setState({ isLoading: true });

    if (!getRecordId(this.props.context)) {
      this.setState({ isLoading: false, sharePointData: [] });
      return;
    }
    this.setState({ isLoading: false, sharePointData: [] });
    try {
      await this.loadSharePointLocations();
      await this.getSharePointData();
    } catch (err) {
      console.error("Error loading files:", err);
    } finally {
      await delay(3000);
      this.setState({ isLoading: false });
    }
  };

  openFile(file: SharePointDocument) {
    if (!file.absoluteurl) {
      toast.error(this.t(LocalStrings.Toast.Message_Open_Missing_Url));
      return;
    }
    try {
      const link = document.createElement("a");
      link.href = file.absoluteurl;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      toast.success(`Opening ${file.fullname} in a new tab.`);
    } catch (error) {
      console.error("Error opening the file:", error);
      toast.error(`${this.t(LocalStrings.Toast.Message_Open_Error)} ${errorMessage(error)}`);
    }
  }

  removeFile = async (fileId?: string) => {
    if (!fileId) {
      console.error("No file ID provided for deletion");
      toast.error("No file ID provided");
      return;
    }

    this.setState({ isLoading: true });

    const file = this.state.sharePointData.find(
      (f: SharePointDocument) => f.sharepointdocumentid === fileId,
    );
    if (!file) {
      this.setState({ isLoading: false });
      toast.error(this.t(LocalStrings.Toast.Message_Delete_Error_File_Not_Found_SharePoint));
      return;
    }
    toast.promise(
      deleteSharePointDocument(
        this.props.context!,
        file.sharepointdocumentid,
        file.documentid,
        file.filetype,
        this.state.selectedDocumentLocation!,
        this.filterByLocation(),
      ),
      {
        loading: this.t(LocalStrings.Toast.Message_Delete_Loading),
        success: () => {
          this.setState((prevState) => ({
            sharePointData: prevState.sharePointData.filter(
              (f) => f.sharepointdocumentid !== fileId,
            ),
            isLoading: false,
          }));
          return this.t(LocalStrings.Toast.Message_Delete_Success);
        },
        error: (err) => {
          this.setState({ isLoading: false });
          console.error("Error deleting SharePoint document: ", err);
          return `${this.t(LocalStrings.Toast.Message_Delete_Failure)} ${errorMessage(err)}`;
        },
      },
    );
  };

  createSharePointFolder = () => {
    const { context } = this.props;
    const { newFolderName, currentFolderPath } = this.state;
    if (!newFolderName) {
      toast.error(this.t(LocalStrings.Toast.Message_Create_Folder_Empty_Error));
      return;
    }
    if (!isValidSharePointFolderName(newFolderName)) {
      toast.error(this.t(LocalStrings.Toast.Message_Create_Folder_Invalid_Name_Error));
      return;
    }
    const folderCreationPromise = createSharePointFolder(
      context!,
      newFolderName,
      currentFolderPath,
      this.state.selectedDocumentLocation!,
      this.filterByLocation(),
    );

    toast.promise(folderCreationPromise, {
      loading: this.t(LocalStrings.Toast.Message_Create_Folder_Loading),
      success: () => {
        this.toggleModal();
        this.loadExistingFiles();
        this.setState({ newFolderName: "" });
        return `${newFolderName} ${this.t(LocalStrings.Toast.Message_Create_Folder_Success)}`;
      },
      error: (err) => {
        return `${this.t(LocalStrings.Toast.Message_Create_Folder_Error)} ${errorMessage(err)}`;
      },
    });
  };

  performActionOnSelectedFiles = (action: FileAction) => {
    this.state.selectedFiles.forEach((fileId) => {
      const file = this.state.sharePointData.find(
        (f: SharePointDocument) => f.sharepointdocumentid === fileId,
      );
      if (!file) {
        toast.error("File not found.");
        return;
      }
      if (action === "open") {
        this.openFile(file);
      } else {
        this.removeFile(file.sharepointdocumentid);
      }
    });

    this.setState({ selectedFiles: [] });
  };

  handleSearch = (newValue?: string) => {
    this.setState({ searchText: newValue || "" });
  };

  toggleRememberLocation = async (_: React.MouseEvent<HTMLElement>, checked?: boolean) => {
    const remember = !!checked;
    const { selectedDocumentLocation, selectedDocumentLocationName } = this.state;
    this.setState({ userPreference: remember }, async () => {
      await saveUserSettings(
        this.props.context!,
        remember,
        selectedDocumentLocation,
        selectedDocumentLocationName,
      );
    });
  };

  async handleDropdownChange(
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
  ): Promise<void> {
    const userPreference = false;
    if (option) {
      const selectedDocumentLocation = option.key as string;
      const selectedDocumentLocationName = option.text as string;

      this.setState(
        {
          userPreference,
          selectedDocumentLocation,
          selectedDocumentLocationName,
        },
        async () => {
          await saveUserSettings(
            this.props.context!,
            userPreference,
            selectedDocumentLocation,
            selectedDocumentLocationName,
          );
          this.setState({ currentFolderPath: "" }, () => {
            this.loadExistingFiles();
          });
        },
      );
    }
  }

  toggleFileSelection = (
    fileId: string,
    file?: SharePointDocument,
    forceSelect: boolean = false,
  ) => {
    if (file && file.filetype === "folder") {
      this.handleFolderClick(file.relativelocation);
      return;
    }

    const isSelected = this.state.selectedFiles.includes(fileId);

    if (forceSelect && !isSelected) {
      this.setState((prevState) => ({
        selectedFiles: [...prevState.selectedFiles, fileId],
      }));
    } else if (!forceSelect) {
      this.setState((prevState) => ({
        selectedFiles: isSelected
          ? prevState.selectedFiles.filter((id) => id !== fileId)
          : [...prevState.selectedFiles, fileId],
      }));
    }
  };

  toggleModal = () => {
    this.setState((prevState) => ({ showModal: !prevState.showModal }));
  };

  handleFolderNameChange = (
    _: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string,
  ) => {
    this.setState({ newFolderName: newValue || "" });
  };

  private t = (key: LocalizationKey): string => getLocalString(this.props.context!, key);

  private handleTileContextMenu = (
    event: React.MouseEvent<HTMLElement>,
    item: SharePointDocument,
  ) => {
    event.preventDefault();
    if (item.filetype !== "folder" && event.button === 2 && item.sharepointdocumentid) {
      const isSelected = this.state.selectedFiles.includes(item.sharepointdocumentid);
      this.toggleFileSelection(item.sharepointdocumentid, item, isSelected);
      this.toggleMenu(event);
    }
  };

  private confirmRemoveFolder = () => {
    this.removeFile(this.state.selectedFolderForDelete?.sharepointdocumentid);
    this.toggleRemoveFolderDialog();
  };

  render() {
    const { isLoading, selectedFiles, documentLocations, currentFolderPath } = this.state;
    const { context, isDisabled } = this.props;
    const t = this.t;

    if (!getRecordId(context)) {
      return (
        <div className="record-not-created-message">
          {t(LocalStrings.Input.Placeholder_Record_Not_Created)}
        </div>
      );
    }

    const formIsDisabled = !this.isEditable();
    const items = filterAndSortSharePointItems(this.state.sharePointData, this.state.searchText);
    const isEmpty = items.length === 0;
    // Hidden until the record has a location: adding one needs an existing
    // location's parent site.
    const showLocationPicker =
      getControlValue(context!, "displaySharePointDocumentLocationsControl") === true &&
      documentLocations.length > 0;

    const locationOptions: IDropdownOption[] = documentLocations.map((location) => ({
      key: location.sharepointdocumentlocationid,
      text: location.name,
    }));
    const parentSiteOptions: IComboBoxOption[] = documentLocations
      .filter((location) => location.parentsiteid)
      .map((location) => ({ key: location.parentsiteid!, text: location.name }));

    return (
      <>
        <Toaster position="top-right" reverseOrder={false} />

        <CreateFolderDialog
          t={t}
          hidden={!this.state.showModal}
          folderName={this.state.newFolderName}
          onFolderNameChange={this.handleFolderNameChange}
          onCreate={this.createSharePointFolder}
          onCancel={this.toggleModal}
        />
        <DeleteFolderDialog
          t={t}
          hidden={!this.state.isFolderDeletionDialogVisible}
          folderName={this.state.selectedFolderForDelete?.fullname}
          onConfirm={this.confirmRemoveFolder}
          onCancel={this.toggleRemoveFolderDialog}
        />
        <AddLocationDialog
          t={t}
          hidden={!this.state.isCreateLocationDialogVisible}
          displayName={this.state.createLocationDisplayName}
          folderName={this.state.createLocationFolderName}
          selectedParentSite={this.state.selectedCreateLocation}
          parentSiteOptions={parentSiteOptions}
          canSave={this.state.isSaveButtonEnabled}
          onInputChange={this.handleInputChange}
          onParentSiteChange={this.handleCreateLocationDropdownChange}
          onSave={this.handlecreateLocation}
          onCancel={this.closeCreateLocationDialog}
        />

        <div className="ribbon-dropzone-wrapper">
          <Ribbon
            t={t}
            selectedFiles={selectedFiles}
            sharePointUrl={this.sharePointUrl()}
            isCollapsed={this.state.isCollapsed}
            formIsDisabled={formIsDisabled}
            canCreateFolder={!formIsDisabled && !!this.state.selectedDocumentLocation}
            menuVisible={this.state.menuVisible}
            menuTarget={this.state.target}
            onAction={this.performActionOnSelectedFiles}
            onSearch={(value) => this.handleSearch(value)}
            onToggleMenu={this.toggleMenu}
            onCloseMenu={this.closeMenu}
            onCreateFolder={this.toggleModal}
          />

          <Dropzone onDrop={this.handleDrop} disabled={formIsDisabled || isDisabled}>
            {({ getRootProps, getInputProps }) => (
              <div className="dropzone-wrapper">
                <div
                  {...getRootProps()}
                  className={`dropzone ${isEmpty ? "empty" : ""}`}
                  style={{ cursor: formIsDisabled || isDisabled ? "default" : "pointer" }}
                >
                  {isLoading ? (
                    <div className="spinner-box">
                      <Spinner size={SpinnerSize.medium} />
                    </div>
                  ) : (
                    <>
                      <input {...getInputProps()} />
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          width: "100%",
                        }}
                      >
                        <div>
                          {currentFolderPath && (
                            <IconButton
                              iconProps={{ iconName: "Back" }}
                              title={t(LocalStrings.Button.Label_Back)}
                              ariaLabel={t(LocalStrings.Button.Label_Back)}
                              onClick={(event) => {
                                event.preventDefault();
                                event.stopPropagation();
                                this.handleBackClick();
                              }}
                            />
                          )}
                        </div>
                        <div style={{ flexGrow: 1, textAlign: "center" }}>
                          {isEmpty ? (
                            <p>{t(LocalStrings.Input.Placeholder_Dropzone)}</p>
                          ) : (
                            <SharePointFileGrid
                              t={t}
                              items={items}
                              selectedFiles={selectedFiles}
                              formIsDisabled={formIsDisabled}
                              onTileClick={(item) =>
                                this.toggleFileSelection(item.sharepointdocumentid, item)
                              }
                              onTileContextMenu={this.handleTileContextMenu}
                              onRemoveFolderClick={this.handleRemoveFolderClick}
                            />
                          )}
                        </div>
                      </div>
                    </>
                  )}
                </div>
              </div>
            )}
          </Dropzone>

          {showLocationPicker && (
            <FooterBar
              t={t}
              isSettingsCalloutVisible={this.state.isCalloutVisible}
              settingsCalloutTarget={this.targetRef.current}
              onSettingsCalloutDismiss={this.onCalloutDismiss}
              rememberLocation={this.state.userPreference}
              onRememberLocationChange={this.toggleRememberLocation}
              locationOptions={locationOptions}
              selectedLocation={this.state.selectedDocumentLocation}
              onLocationChange={this.handleDropdownChange}
              canAddLocation={!formIsDisabled}
              addLocationRef={this.targetRef}
              onAddLocation={this.openCreateLocationDialog}
              onSettingsClick={this.onGearIconClick}
            />
          )}
        </div>
      </>
    );
  }
}

export default Landing;
