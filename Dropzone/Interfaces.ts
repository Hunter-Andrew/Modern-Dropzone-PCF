export interface SharePointDocument {
  readurl: string;
  modified: string;
  modifiedFormatted: string;
  sharepointmodifiedby: string;
  editurl: string;
  sharepointdocumentid: string;
  documentid: number;
  documentidFormatted: string;
  ischeckedout: boolean;
  ischeckedoutFormatted: string;
  sharepointcreatedon: string;
  sharepointcreatedonFormatted: string;
  locationname: string;
  iconclassname: string;
  absoluteurl: string;
  fullname: string;
  locationid: string;
  title: string;
  filetype: string;
  relativelocation: string;
  servicetypeFormatted: string;
  servicetype: number;
  author: string;
  filesize: number;
  filename?: string;
}

export interface EntityMetadata {
  schemaName: string;
  logicalCollectionName: string;
  clientUrl: string;
  entityId: string;
}

export type ActivityType =
  | "email"
  | "phonecall"
  | "appointment"
  | "task"
  | "fax"
  | "letter"
  | "serviceappointment"
  | "campaignresponse"
  | "campaignactivity"
  | "bulkoperation"
  | "socialactivity"
  | "recurringappointmentmaster"
  | "appointmentrecurrence";

/** What the ribbon and context menu can do with the selected files. */
export type FileAction = "open" | "delete";

export interface DocumentLocation {
  name: string;
  sharepointdocumentlocationid: string;
  parentsiteid?: string;
}
