import { IInputs } from "./generated/ManifestTypes";
import { DocumentLocation, SharePointDocument } from "./Interfaces";
import { escapeXml, getEntityMetadata, isActivityType } from "./utils";

/**
 * A failed Dataverse request, carrying the HTTP status and Dataverse's own
 * error message (from the response body) when there is one.
 */
export class DataverseError extends Error {
  constructor(
    message: string,
    public readonly status: number,
  ) {
    super(message);
    this.name = "DataverseError";
  }
}

/**
 * Builds a DataverseError from a failed response. Dataverse puts the reason in
 * the body ({"error": {"message": ...}}); statusText is often empty over HTTP/2.
 */
export async function toDataverseError(
  response: Response,
  action: string,
): Promise<DataverseError> {
  let detail = "";
  try {
    const text = await response.text();
    try {
      detail = JSON.parse(text)?.error?.message ?? text;
    } catch {
      detail = text;
    }
  } catch {
    // No readable body.
  }
  detail = (detail || response.statusText || "").trim();
  const message = `${action} (HTTP ${response.status})${detail ? `: ${detail}` : ""}`;
  return new DataverseError(message, response.status);
}

export async function getSharePointLocations(
  context: ComponentFramework.Context<IInputs>,
): Promise<DocumentLocation[]> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return [];
  }
  const fetchXml = `
  <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="sharepointdocumentlocation">
    <attribute name="statecode" />
    <attribute name="name" />
    <attribute name="sharepointdocumentlocationid" />
    <attribute name="servicetype" />
    <attribute name="parentsiteorlocation" />
    <order attribute="modifiedon" descending="true" />
    <filter type="and">
      <condition attribute="statuscode" operator="eq" value="1" />
      <condition attribute="statecode" operator="eq" value="0" />
    </filter>
    <filter type="and">
      <filter type="and">
        <condition attribute="regardingobjectid" operator="eq" value="${escapeXml(metadata.entityId)}" />
      </filter>
      <filter type="and">
        <condition attribute="servicetype" operator="in">
          <value>0</value>
        </condition>
      </filter>
    </filter>
  </entity>
</fetch>
  `;
  const encodedFetchXml = encodeURIComponent(fetchXml);
  const url = `${metadata.clientUrl}/api/data/v9.0/sharepointdocumentlocations?fetchXml=${encodedFetchXml}`;

  const headers = {
    Accept: "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
  };

  const response = await fetch(url, {
    method: "GET",
    headers: headers,
  });

  if (!response.ok) {
    throw await toDataverseError(response, "Error fetching SharePoint document locations");
  }
  let data = await response.json();
  return data.value.map((item: any) => ({
    name: item.name,
    sharepointdocumentlocationid: item.sharepointdocumentlocationid,
    parentsiteid: item._parentsiteorlocation_value,
  }));
}
/**
 * The absolute URL of a document location's folder in SharePoint, e.g.
 * https://contoso.sharepoint.com/sites/crm/account/Contoso_1234.
 */
export async function getLocationUrl(
  context: ComponentFramework.Context<IInputs>,
  locationId: string,
): Promise<string> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return "";
  }
  const id = locationId.replace(/[{}]/g, "");
  const url = `${metadata.clientUrl}/api/data/v9.0/sharepointdocumentlocations(${id})/Microsoft.Dynamics.CRM.RetrieveAbsoluteAndSiteCollectionUrl()`;
  const response = await fetch(url, {
    headers: {
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
  });
  if (!response.ok) {
    throw await toDataverseError(response, "Error getting the SharePoint folder URL");
  }
  const data = await response.json();
  return data.AbsoluteUrl ?? "";
}

export async function getSharePointFolderData(
  context: ComponentFramework.Context<IInputs>,
  folderPath?: string,
  selectedDocumentLocation?: string | null,
  selectedDocumentLocationName?: string,
  defaultSite?: boolean,
): Promise<SharePointDocument[]> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return [];
  }
  const isActivity = isActivityType(metadata.schemaName);
  let folderFilter = "";
  let filters = "";
  if (folderPath) {
    folderFilter = `
      <filter type="and">
        <condition attribute="relativelocation" operator="eq" value="${escapeXml(folderPath)}"/>
      </filter>`;
  }
  if (selectedDocumentLocation) {
    if (defaultSite) {
      filters = `
      <filter type="and">
         <filter type="and">
            <condition attribute="locationid" operator="eq" value="${escapeXml(selectedDocumentLocation)}" />
         </filter>
         <filter type="and">
            <condition attribute="locationname" operator="eq" value="${escapeXml(selectedDocumentLocationName)}" />
         </filter>
         <filter type="and">
            <condition attribute="servicetype" operator="eq" value="0" />
         </filter>
         ${folderFilter}
      </filter>`;
    } else {
      filters = `${folderFilter}`;
    }
  }
  const fetchXml = `
    <fetch distinct="false" mapping="logical" returntotalrecordcount="true" no-lock="false">
      <entity name="sharepointdocument">
        <attribute name="documentid"/>
        <attribute name="fullname"/>
        <attribute name="relativelocation"/>
        <attribute name="sharepointcreatedon"/>
        <attribute name="ischeckedout"/>
        <attribute name="filetype"/>
        <attribute name="modified"/>
        <attribute name="sharepointmodifiedby"/>
        <attribute name="servicetype"/>
        <attribute name="absoluteurl"/>
        <attribute name="title"/>
        <attribute name="author"/>
        <attribute name="sharepointdocumentid"/>
        <attribute name="readurl"/>
        <attribute name="editurl"/>
        <attribute name="locationid"/>
        <attribute name="iconclassname"/>
        <attribute name="locationname"/>
        <attribute name="filesize"/>
        <filter>
          <condition attribute="isrecursivefetch" operator="eq" value="0"/>
        </filter>
${filters}
        <order attribute="relativelocation" descending="false"/>
        <link-entity name="${metadata.schemaName}" from="${isActivity ? "activity" : metadata.schemaName}id" to="regardingobjectid" alias="bb">
          <filter type="and">
            <condition attribute="${isActivity ? "activity" : metadata.schemaName}id" operator="eq" uitype="${isActivity ? "activity" : metadata.schemaName}" value="${escapeXml(metadata.entityId)}"/>
          </filter>
        </link-entity>
      </entity>
    </fetch>
  `;

  const encodedFetchXml = encodeURIComponent(fetchXml);
  const url = `${metadata.clientUrl}/api/data/v9.0/sharepointdocuments?fetchXml=${encodedFetchXml}`;

  const headers = {
    Accept: "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
  };

  const response = await fetch(url, {
    method: "GET",
    headers: headers,
  });

  if (!response.ok) {
    throw await toDataverseError(response, "Error fetching SharePoint data");
  }

  const data = await response.json();

  return data.value.map((item: any) => ({
    readurl: item.readurl,
    modified: item.modified,
    modifiedFormatted: item["modified@OData.Community.Display.V1.FormattedValue"],
    sharepointmodifiedby: item.sharepointmodifiedby,
    editurl: item.editurl,
    sharepointdocumentid: item.sharepointdocumentid,
    documentid: item.documentid,
    documentidFormatted: item["documentid@OData.Community.Display.V1.FormattedValue"],
    ischeckedout: item.ischeckedout,
    ischeckedoutFormatted: item["ischeckedout@OData.Community.Display.V1.FormattedValue"],
    sharepointcreatedon: item.sharepointcreatedon,
    sharepointcreatedonFormatted:
      item["sharepointcreatedon@OData.Community.Display.V1.FormattedValue"],
    locationname: item.locationname,
    iconclassname: item.iconclassname,
    absoluteurl: item.absoluteurl,
    fullname: item.fullname,
    locationid: item.locationid,
    title: item.title,
    filetype: item.filetype,
    relativelocation: item.relativelocation,
    servicetypeFormatted: item["servicetype@OData.Community.Display.V1.FormattedValue"],
    servicetype: item.servicetype,
    author: item.author,
    filesize: item.filesize,
  }));
}

async function getSiteUrl(
  context: ComponentFramework.Context<IInputs>,
  parentLocationId: string,
): Promise<string> {
  const metadata = await getEntityMetadata(context);
  const url = `${metadata?.clientUrl}/api/data/v9.0/FetchSiteUrl`;
  const body = JSON.stringify({
    DocumentId: parentLocationId,
    ParentEntityReference: {
      "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
      [`${metadata?.schemaName}id`]: `${metadata?.entityId}`,
    },
  });

  const fetchOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: body,
  };
  try {
    const response = await fetch(url, fetchOptions);
    if (!response.ok) {
      throw await toDataverseError(response, "Error getting SharePoint site");
    }
    const responseData = await response.json();
    if (!responseData.SiteAndLocationUrl) {
      throw new Error("SiteAndLocationUrl not found in response");
    }
    return responseData.SiteAndLocationUrl;
  } catch (error) {
    console.error("Error getting SharePoint site: ", error);
    throw error;
  }
}

export async function createSharePointLocation(
  context: ComponentFramework.Context<IInputs>,
  locationName: string,
  relativePath: string,
  parentLocationId: string,
): Promise<string> {
  const metadata = await getEntityMetadata(context);
  const url = `${metadata?.clientUrl}/api/data/v9.0/AddOrEditLocation`;
  const absUrl = await getSiteUrl(context, parentLocationId);
  const body = JSON.stringify({
    AbsUrl: absUrl + "/" + locationName,
    DocumentId: "",
    IsAddOrEditMode: true,
    IsCreateFolder: true,
    LocationName: locationName,
    ParentEntityReference: {
      "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
      [`${metadata?.schemaName}id`]: metadata?.entityId,
    },
    ParentId: parentLocationId,
    ParentType: "sharepointdocumentlocation",
    RelativePath: locationName,
  });

  const fetchOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "OData-Version": "4.0",
      "OData-MaxVersion": "4.0",
      Accept: "application/json",
    },
    body: body,
  };

  try {
    const response = await fetch(url, fetchOptions);
    if (!response.ok) {
      throw await toDataverseError(response, "Failed to create a location");
    }
    const responseData = await response.json();
    if (!responseData.LocationId) {
      throw new Error("LocationId not found in response");
    }
    return responseData.LocationId;
  } catch (error) {
    console.error("Error creating a location:", error);
    throw error;
  }
}

export async function createSharePointDocument(
  context: ComponentFramework.Context<IInputs>,
  fileName: string,
  dataURL: string,
  folderPath: string,
  locationId: string,
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return;
  }
  const isActivity = isActivityType(metadata.schemaName);
  const url = `${metadata?.clientUrl}/api/data/v9.0/UploadDocument`;
  const base64 = dataURL.split(",")[1];
  const body = JSON.stringify({
    Content: base64,
    Entity: {
      "@odata.type": "Microsoft.Dynamics.CRM.sharepointdocument",
      locationid: locationId,
      title: `${fileName}`,
    },
    OverwriteExisting: true,
    ParentEntityReference: {
      "@odata.type": `Microsoft.Dynamics.CRM.${metadata.schemaName}`,
      [`${isActivity ? "activity" : metadata.schemaName}id`]: `${metadata?.entityId}`,
    },
    FolderPath: `${folderPath}`,
  });

  const fetchOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: body,
  };

  try {
    const response = await fetch(url, fetchOptions);
    if (!response.ok) {
      throw await toDataverseError(response, "Failed to upload document");
    }
  } catch (error) {
    console.error("Error uploading document:", error);
    throw error;
  }
}

export async function deleteSharePointDocument(
  context: ComponentFramework.Context<IInputs>,
  sharepointDocumentId: string,
  documentId: number,
  fileType: string,
  rawlocationId: string,
  isDefaultLocation: boolean,
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return;
  }
  let locationId = rawlocationId;
  if (!isDefaultLocation) {
    locationId = "00000000-0000-0000-0000-000000000000";
  }
  const isActivity = isActivityType(metadata.schemaName);
  const url = `${metadata?.clientUrl}/api/data/v9.0/DeleteDocument`;
  const body = JSON.stringify({
    Entities: [
      {
        "@odata.type": "Microsoft.Dynamics.CRM.sharepointdocument",
        sharepointdocumentid: `{${sharepointDocumentId.toUpperCase()}}`,
        documentid: documentId,
        filetype: fileType,
        locationid: locationId,
      },
    ],
    ParentEntityReference: {
      "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
      [`${isActivity ? "activity" : metadata.schemaName}id`]: `${metadata?.entityId}`,
    },
  });

  const fetchOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: body,
  };

  try {
    const response = await fetch(url, fetchOptions);
    if (!response.ok) {
      throw await toDataverseError(response, "Failed to delete document");
    }
  } catch (error) {
    console.error("Error deleting document:", error);
    throw error;
  }
}

export async function createSharePointFolder(
  context: ComponentFramework.Context<IInputs>,
  folderName: string,
  folderPath: string,
  rawlocationId: string,
  isDefaultLocation: boolean,
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return;
  }

  let locationId = rawlocationId;
  if (!isDefaultLocation) {
    locationId = "00000000-0000-0000-0000-000000000000";
  }

  const isActivity = isActivityType(metadata.schemaName);
  const url = `${metadata?.clientUrl}/api/data/v9.0/NewDocument`;
  const body = JSON.stringify({
    FileName: folderName,
    LocationId: locationId,
    IsFolder: true,
    FolderPath: folderPath,
    ParentEntityReference: {
      "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
      [`${isActivity ? "activity" : metadata.schemaName}id`]: `${metadata?.entityId}`,
    },
  });

  const fetchOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: body,
  };

  try {
    const response = await fetch(url, fetchOptions);
    if (!response.ok) {
      throw await toDataverseError(response, "Failed to create a folder");
    }
  } catch (error) {
    console.error("Error creating a folder:", error);
    throw error;
  }
}
