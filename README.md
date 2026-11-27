![Dropzone-main](https://github.com/user-attachments/assets/6c182dac-beed-4bfa-a700-3e5aeb618011)
![Dropzone-key-features](https://github.com/GorgonUK/Modern-Dropzone-PCF/assets/59618079/8ad3d457-328e-4563-adb5-01f1958b1c4a)


## Installation

The component supports SharePoint Document Locations and comes pre-configured to work with your existing Document Locations. For setup instructions, please refer to [Microsoft Learn's guide on enabling SharePoint document management for specific entities](https://learn.microsoft.com/en-us/power-platform/admin/enable-sharepoint-document-management-specific-entities).

### Steps to Install:

1. **Set up SharePoint document management**

   The control shows and uploads the SharePoint documents of the record it's on, so SharePoint document management must be enabled for the table (see the guide above). Nothing else needs configuring: the control finds the record's document locations and its SharePoint folder by itself.

2. **Install the Solution**

   Install the Dropzone PCF solution package into your Dynamics 365 environment via the Solution Management area.

3. **Add Component to a Form**

   Open the desired form in edit mode and navigate to **Get more components**:
   
   ![Get more components](https://github.com/GorgonUK/DropzonePCF/assets/59618079/d737906e-29f2-4217-bb04-09a748ff3209)

4. **Search and Add Dropzone**

   In the components search, type "Dropzone" in the search box, select the component, and then press "Add":
   
   ![Select Dropzone](https://github.com/GorgonUK/DropzonePCF/assets/59618079/5e1c1298-bd50-4d0e-8b10-4110ebc5dc44)

5. **Position the Component on the Form**

   The component will appear under "More components". You can now drag and drop it anywhere on the form:
   
   ![More components](https://github.com/GorgonUK/DropzonePCF/assets/59618079/fdea28a7-925f-46db-b698-8edcf26d9348)

   Drag the component to your desired location on the form. Note that the "table" and "view" selected during this step do not affect the component's functionality.

   ![Dropzone Component on Form](https://github.com/GorgonUK/Modern-Dropzone-PCF/assets/59618079/e9461aa3-fb3f-4170-a71c-7a96ef1ca421)

6. **Configure the component (optional)**

   | Setting | Default | What it does |
   | --- | --- | --- |
   | Allow uploads? | Yes | Lets users upload files by dropping them on the control or browsing for them. |
   | Show document location picker? | Yes | Shows the picker for switching between the record's SharePoint document locations, with options to add a location and to remember the selected one. |

   To make the control read-only, disable it on the form; that also hides Delete and Create Folder.

### Upload size

Uploads go through Dataverse (the same `UploadDocument` call the out-of-the-box Documents grid uses), which sends the whole file in a single request. Microsoft doesn't document a limit for it; community reports put it at about 50 MB, and there's no known setting to raise it. When an upload fails because the file is too large, the control says so and links to the record's folder in SharePoint, where files of any size up to SharePoint's own limit can be uploaded directly.

For any issues or suggestions, use the "Issues" tab at the top of the page.

## Q: Why can’t my users see files in the drop zone?  
**A:** The files live in SharePoint. Users need access to the record, read access to the Document Location table in Dataverse, and access to the folder in SharePoint.

---

## Q: Why does the drop zone ask me to create the record first?  
**A:** A record gets its SharePoint folder once it exists. After the first save, the control checks for the folder for about 20 seconds and shows it when it appears.

---

## Q: Why does the drop zone show unrelated documents, and the subgrid display a throttling error?  
**A:** The drop zone is not currently designed to handle document libraries with over 5000 items. This limitation may cause unexpected behavior, such as displaying unrelated documents. A future improvement will address this issue.

---

## Q: How can I turn off uploads?  
**A:** Edit the component on the form and set **Allow uploads?** to `No`. Users can still browse, open and delete files; disable the control on the form to make it fully read-only.

---

## Building from source

Requires Node.js 18+ and the [Power Platform CLI](https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction) (`pac`). Building the solution package also requires the .NET SDK.

```bash
npm ci                # install dependencies
npm run build         # build the control into out/controls (also runs ESLint)
npm run start:watch   # open the PCF test harness with live reload
```

To produce importable solution zips (unmanaged and managed):

```bash
cd Solution/ModernDropzone
dotnet build -c Release -p:SolutionPackageType=Both
# output: bin/Release/ModernDropzone.zip (unmanaged) and ModernDropzone_managed.zip
```

### Tests

```bash
npm test
```

Unit tests live in `tests/` and run with Jest against a mocked PCF context, so they need no Dynamics environment. They cover the data loading (`Dropzone/Landing/landingData.ts`), the display helpers and the shared utilities.

### Formatting

Code is formatted with Prettier. Run `npm run format` before committing; CI fails on unformatted code (`npm run format:check`).

### CI

The [Build solution](.github/workflows/build-solution.yml) workflow runs on every push, to any branch. It builds both zips and attaches them to the run as the `ModernDropzone-solution` artifact (Actions tab → the run → Artifacts).

To publish a release, bump the version (see below), then push a tag starting with `v`, for example `v2.9.6`. The workflow creates a GitHub release with the zips attached.

### Versioning

The version is written in several files: `package.json`, the control manifest, the solution's `Solution.xml` and `Dropzone/version.ts`. Dataverse only picks up a new build of the control when the manifest version goes up. Set it everywhere at once with:

```bash
npm run version:set -- 2.9.6
```

CI fails if the files disagree (`npm run version:check`), or if a release tag doesn't match the version.
