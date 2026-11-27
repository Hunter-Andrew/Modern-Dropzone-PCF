export const LocalStrings = {
  App: {
    Name: "App__Name",
    Description: "App__Description",
  },
  Input: {
    Placeholder_Search: "Input__Placeholder_Search",
    Placeholder_Dropzone: "Input__Placeholder_Dropzone",
    Placeholder_Record_Not_Created: "Input__Placeholder_Record_Not_Created",
  },
  Button: {
    Label_Open: "Button__Label_Open",
    Label_Delete: "Button__Label_Delete",
    Label_Actions: "Button__Label_Actions",
    Label_CreateFolder: "Button__Label_CreateFolder",
    Label_Settings: "Button__Label_Settings",
    Label_RememberLocation: "Button__Label_RememberLocation",
    Label_Back: "Button__Label_Back",
    Label_Remove: "Button__Label_Remove",
    Label_Remove_Folder: "Button__Label_Remove_Folder",
  },
  Dialog: {
    DeleteFolder: {
      Confirmation: "Dialog__Button_Delete_Confirmation",
      Button_Yes: "Dialog__Button_Yes",
      Button_No: "Dialog__Button_No",
    },
    CreateFolder: {
      Title: "Dialog__Title_CreateFolder",
      Description: "Dialog__Description_CreateFolder",
      Input_Placeholder: "Dialog__Input_Placeholder_CreateFolder",
      Button_Create: "Dialog__Button_Create_CreateFolder",
      Button_Cancel: "Dialog__Button_Cancel_CreateFolder",
    },
    AddLocation: {
      Title: "Dialog__Title_AddLocation",
      Description: "Dialog__Description_AddLocation",
      Input_Label_DisplayName: "Dialog__Input_Label_DisplayName_AddLocation",
      Input_Label_ParentSite: "Dialog__Input_Label_ParentSite_AddLocation",
      Input_Label_FolderName: "Dialog__Input_Label_FolderName_AddLocation",
      Button_Save: "Dialog__Button_Save_AddLocation",
      Button_Cancel: "Dialog__Button_Cancel_AddLocation",
    },
  },
  Link: {
    /** "View files in {0}", where {0} is the SharePoint link. */
    View_In_SharePoint: "Link__View_In_SharePoint",
    SharePoint: "Link__SharePoint",
  },
  Toast: {
    Message_Upload_Success_SharePoint: "Toast__Message_Upload_Success_SharePoint",
    Message_Upload_Error_SharePoint: "Toast__Message_Upload_Error_SharePoint",
    Message_Refresh_Error_SharePoint: "Toast__Message_Refresh_Error_SharePoint",
    Message_Uploading_SharePoint: "Toast__Message_Uploading_SharePoint",
    Message_Read_Error_SharePoint: "Toast__Message_Read_Error_SharePoint",
    /** "{0} ({1}) is too large to upload here." {0}: file name, {1}: size. */
    Message_Upload_Too_Large: "Toast__Message_Upload_Too_Large",
    /** "Upload it in {0} instead.", where {0} is the SharePoint link. */
    Message_Upload_Too_Large_Hint: "Toast__Message_Upload_Too_Large_Hint",
    /** "{0} ({1}) could not be read. ..." {0}: file name, {1}: size. */
    Message_Read_Error_Large: "Toast__Message_Read_Error_Large",

    Message_Open_Error: "Toast__Message_Download_Open_Error_SharePoint",
    Message_Open_Missing_Url: "Toast__Message_Download_Error_SharePoint",

    Message_Delete_Error_File_Not_Found_SharePoint:
      "Toast__Message_Delete_Error_File_Not_Found_SharePoint",
    Message_Delete_Loading: "Toast__Message_Delete_Loading",
    Message_Delete_Success: "Toast__Message_Delete_Success",
    Message_Delete_Failure: "Toast__Message_Delete_Failure",

    Message_Create_Folder_Loading: "Toast__Message_Create_Folder_Loading",
    Message_Create_Folder_Success: "Toast__Message_Create_Folder_Success",
    Message_Create_Folder_Error: "Toast__Message_Create_Folder_Error",
    Message_Create_Folder_Empty_Error: "Toast__Message_Create_Folder_Empty_Error",
    Message_Create_Folder_Invalid_Name_Error: "Toast__Message_Create_Folder_Invalid_Name_Error",
  },
} as const;
