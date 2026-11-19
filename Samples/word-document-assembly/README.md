---
title: Assemble a Statement of Work in Word
page_type: sample
urlFragment: word-document-assembly
products:
  - office-word
  - office
  - m365
languages:
  - typescript
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "11/17/2026 12:00:00 AM"
description: "Assemble a Statement of Work with reusable content, customer controls, project costs, and a configurable team."
---

# Assemble a Statement of Work in Word

Build an illustrative Statement of Work (SOW) using a Word task-pane add-in.

- Insert an OOXML template, boilerplate paragraphs, and a company logo.
- Find customer names, wrap them in content controls, and update them together.
- Insert a bundled project-costs document into a content control.
- Find currency amounts with wildcards and add a confidentiality footer.
- Choose four consultants and hours per person, then insert a customized OOXML team table.

The template, terms, historical dates, and costs are illustrative, not legal or financial advice. Consultants A-F are synthetic demo labels. All demo content is bundled locally; no database, authentication service, or live pricing source is used.

## Applies to

This sample uses a **unified manifest for Microsoft 365**, production schema **1.30**, and **WordApi 1.3** (including language-independent paragraph styles).

- Word on Windows with a Microsoft 365 subscription, Version 2501 (Build 18407.20002) or later.
- Word on Mac, Version 16.103 (25101427) or later.
- Word on the web.

Older perpetual Office versions and mobile clients aren't supported by this manifest. See [unified manifest client and platform support](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support).

Builds and static checks do not establish Office-host compatibility. Manually test the complete workflow on each intended platform before relying on it.

## Prerequisites

- An active LTS release of [Node.js](https://nodejs.org/) version 22 or later, and npm.
- A supported Word client and Microsoft 365 account that permits sideloading add-ins.
- Internet access for the Office.js CDN and initial npm dependency installation.
- An available local HTTPS port 3000.

No Microsoft Graph permissions or app registration are required. The manifest requests only `Document.ReadWrite.User` to modify the open document.

## Run the sample

1. Clone or download this repository. Open a terminal in **Samples/word-document-assembly**.
1. Install the dependencies.

   ```console
   npm install
   ```

   If npm reports blocked lifecycle scripts, review them under your environment's package-execution policy before sideloading. The Office debugging tools may need native credential-store modules such as `keytar`; a successful build does not verify those modules or account sign-in.

1. On Windows, close Word and run the following command.

   ```console
   npm start
   ```

   This starts the local HTTPS webpack server and sideloads the unified manifest into Word. Accept the development certificate prompt if prompted. On first use, you may need to select the add-in from **Home > Add-ins** to finish activation.

   For Mac or Word on the web, run `npm run dev-server`, leave it running, and follow [Sideload Office Add-ins that use the unified manifest](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest) for your platform. If a ZIP package is needed, package **manifest.json**, **assets/color.png**, and **assets/outline.png**, preserving those paths. The package does not replace the HTTPS server. Sideloading from Windows also makes the add-in available to Word on the web for that account.

1. Open a **blank or disposable document**. On the **Home** tab, in **Document Assembly**, select **Open Assembly**.
1. Follow these steps in the task pane.

   | Action | Expected result |
   | --- | --- |
   | Check the replacement confirmation, then **1. Insert Template** | The document body is replaced with the SOW and a small inline logo. Existing headers and footers aren't reset. |
   | **2. Add Paragraphs** | Timeline, costs, team, and terms headings are appended, including a `ProjectCosts` content control and a page break. |
   | **3. Change Logo** | The first inline image is replaced with the bundled sample logo. |
   | **4. Add Customer Content Controls** | Each whole-word `Contoso` match becomes a highlighted, tagged customer control. |
   | Enter a customer name, then **5. Change Customer** | All controls tagged `customer` change to the entered name. The default is `Fabrikam`. |
   | **6. Insert Project Costs File** | The bundled document's cost table replaces the placeholder inside `ProjectCosts`. |
   | **7. Search with Wildcards** | USD amounts such as `$442,896.00` are highlighted, and the final match is selected. |
   | **8. Insert Footer** | The first section's primary footer is replaced with `Confidential`. Existing footer content in that section is overwritten. |

1. In **Team builder**, choose consultants and enter hours **per person**. Each role uses its selected consultant's hourly rate; a consultant may fill multiple roles. With the default choices and 16 hours, the estimate is **$8,320.00**, and the average hourly rate is **$130.00**.
1. Place the Word cursor after **Project Team**, then select **Insert Team at Selection**. An eight-row, three-column table is inserted, including the header, four assignments, average rate, hours, and total. Any current document selection is replaced.
1. To start over, confirm replacement and insert the template again. Adding boilerplate or customer controls twice reports an error rather than duplicating those controls. Costs, customer names, and the footer can be updated repeatedly. The task pane reports missing prerequisites, load failures, and Office API errors.
1. Stop debugging and remove the sideloaded add-in when finished.

   ```console
   npm stop
   ```

   If you used a manual installation or another sideloading tool, also follow that tool's uninstall instructions and stop the HTTPS server with **Ctrl+C**. Closing Word alone does not reliably end a debugging session.

## Key parts of the sample

### Start with reusable document content

Inserting the template uses [Body.insertOoxml](https://learn.microsoft.com/javascript/api/word/word.body#word-word-body-insertooxml-member(1)) to preserve the SOW's rich formatting. The [bundled OOXML](./assets/statement-of-work.xml) is a Flat OPC package rather than a ZIP-encoded `.docx` file. An explicit confirmation protects against accidentally replacing the active document body.

The next step uses [Body.insertParagraph](https://learn.microsoft.com/javascript/api/word/word.body), language-independent paragraph styles, and a page break to append reusable sections. [InlinePicture.insertInlinePictureFromBase64](https://learn.microsoft.com/javascript/api/word/word.inlinepicture) then replaces the inline logo. This replacement API requires WordApi 1.2.

These actions use `Word.run`, load collection items before reading them, and synchronize queued changes with `context.sync`. See [the document assembly implementation](./src/assembly/document-assembly.ts) and [Word's application-specific API model](https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model).

### Turn customer text into reusable fields

Searching for `Contoso` returns ranges that are wrapped using [Range.insertContentControl](https://learn.microsoft.com/javascript/api/word/word.range). Each control receives the tag `customer`. Subsequent updates use [ContentControlCollection.getByTag](https://learn.microsoft.com/javascript/api/word/word.contentcontrolcollection) instead of searching for the customer's current text. Changing the name therefore continues to work after the first replacement.

The costs placeholder is addressed independently by its title, `ProjectCosts`. [ContentControl.insertFileFromBase64](https://learn.microsoft.com/javascript/api/word/word.contentcontrol) inserts the [bundled costs document](./assets/project-costs.docx) without replacing the rest of the SOW. The file's package bytes must be Base64 encoded; the Flat OPC template must not be.

### Review amounts and identify the document

Wildcard [Body.search](https://learn.microsoft.com/javascript/api/word/word.body) finds USD amounts using `$[0-9,]@.[0-9]{2}`. This pattern deliberately targets demo amounts with two decimal places, not every international currency format. Matching ranges are formatted and the last match is selected. See [Use search options in your Word add-in](https://learn.microsoft.com/office/dev/add-ins/word/search-option-guidance).

Finally, [Section.getFooter](https://learn.microsoft.com/javascript/api/word/word.section) accesses the first section's primary footer. Replacing its content makes repeated use predictable; it does not modify other sections or first-page/even-page footers.

### Configure and insert a project team

The [team builder](./src/assembly/team.ts) calculates the average hourly rate and the combined cost for four assignments. It validates finite, positive hours up to 10,000 and escapes text before embedding it in an OOXML table. The Common API [Document.setSelectedDataAsync](https://learn.microsoft.com/javascript/api/office/office.document) inserts that table at the Word selection with `Office.CoercionType.Ooxml`; its callback is wrapped in a promise so failures reach the same visible error handler.

This demonstrates both the Word-specific batch API and the Common API. For guidance on choosing OOXML for rich content, see [Use Office Open XML in Word add-ins](https://learn.microsoft.com/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml).

## Development and validation

The TypeScript/npm/webpack project was scaffolded with Microsoft 365 Agents Toolkit CLI **1.1.17**, using `office-addin-wxpo-taskpane` with the Word host. Running the sample does not require the Toolkit CLI. Native HTML controls and browser APIs provide the task-pane UI and asset loading. There is one hosted task-pane runtime and no function-command runtime or backend.

From the sample directory:

```console
npm run type-check
npm run lint
npm run build
npm run validate
```

`npm run build` writes the local development package's web files to **dist**. The manifest deliberately retains `https://localhost:3000` URLs; this is not a production deployment. Serve it with `npm run dev-server`. `npm run validate` validates the manifest; local file paths and Office behavior must also be checked. For repeatable dependency installation after the lockfile is present, use `npm ci`.

The type-check command supplies compiler options and source files directly, without a separate TypeScript configuration file. Manually test all numbered actions, repeated actions, error messages, logo replacement, cost-table rendering, wildcard matches, and cursor placement on Windows, Mac, and the web using the versions listed above.

## Solution

| Solution | Author(s) |
| --- | --- |
| Assemble a Statement of Work in Word | Microsoft |

## Version history

| Version | Date | Comments |
| --- | --- | --- |
| 1.0 | November 17, 2026 | Initial release in this repository. |

## Additional resources

- [Word add-ins overview](https://learn.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
- [Word JavaScript API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/word/word-api-requirement-sets)
- [Unified manifest for Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview)
- [Microsoft 365 app manifest schema](https://learn.microsoft.com/microsoft-365/extensibility/schema)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2015, 2026 Microsoft Corporation. All rights reserved.

This sample is licensed under the [MIT license](./LICENSE.txt).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-document-assembly" />
