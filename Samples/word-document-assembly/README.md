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
  createdDate: "11/20/2026 12:00:00 AM"
description: "Assemble a Statement of Work with reusable content, customer controls, project costs, and a configurable team."
---

# Assemble a Statement of Work in Word

Build an illustrative Statement of Work (SOW) using a Word task-pane add-in.

- Insert an OOXML template, boilerplate paragraphs, and a company logo.
- Find customer names, wrap them in content controls, and update them together.
- Insert a bundled project-costs document into a content control.
- Find currency amounts with wildcards and add a confidentiality footer.
- Choose four consultants and hours per person, then insert a customized OOXML team table.

The template, terms, services agreement date, and costs are illustrative, not legal or financial advice. The SOW effective date is set to the current local date whenever the template is inserted; the services agreement date is unchanged. Consultants A-F are demo labels. All demo content is bundled locally; no database, authentication service, or live pricing source is used.

## Applies to

This sample uses a **unified manifest for Microsoft 365** and **WordApi 1.3** (including language-independent paragraph styles).

- Word on Windows, Version 2501 (Build 18407.20002) or later.
- Word on Mac, Version 16.103 (25101427) or later.
- Word on the web.

For guidance on clients and platforms that support the unified manifest, see [Office Add-ins with the unified app manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support).

## Prerequisites

- An active LTS release of [Node.js](https://nodejs.org/) and npm.
- A supported Word client and Microsoft 365 account that permits sideloading add-ins.

## Run the sample

1. Clone or download this repository. Open a terminal in **Samples/word-document-assembly**.
1. Install the dependencies.

   ```console
   npm install
   ```

1. On Windows, close Word and run the following command.

   ```console
   npm start
   ```

   This starts the local HTTPS webpack server and sideloads the unified manifest into Word. Accept the development certificate prompt if prompted.

   Sideloading from Windows also makes the add-in available to Word on the web for that account.

1. From the **Home** tab, select **Open Assembly**.
1. Follow these steps in the task pane.

   | Action | Expected result |
   | --- | --- |
   | Check the replacement confirmation, then **1. Insert Template** | The document body is replaced with the SOW and an inline logo. Existing headers and footers aren't reset. |
   | **2. Add Paragraphs** | Timeline, costs, team, and terms headings are appended, including a `ProjectCosts` content control and a page break. |
   | **3. Change Logo** | The inline image is replaced with the bundled sample logo. |
   | **4. Add Customer Content Controls** | Each whole-word `Contoso` match becomes a highlighted, tagged customer control. |
   | Enter a customer name, then **5. Change Customer** | All controls tagged `customer` change to the entered name. The default is `Fabrikam`. |
   | **6. Insert Project Costs File** | The bundled document's cost table replaces the placeholder inside `ProjectCosts`. |
   | **7. Search with Wildcards** | USD amounts such as `$442,896.00` are highlighted, and the final match is selected. |
   | **8. Insert Footer** | The first section's primary footer is replaced with `Confidential`. Existing footer content in that section is overwritten. |

1. In **Team builder**, choose consultants and enter hours **per person**. Each role uses its selected consultant's hourly rate; a consultant may fill multiple roles. With the default choices and 16 hours, the estimate is **$8,320.00**, and the average hourly rate is **$130.00**.
1. Place the Word cursor after **Project Team**, then select **Insert Team at Selection**. An eight-row, three-column table is inserted, including the header, four assignments, average rate, hours, and total. Any current document selection is replaced.
1. To start over, confirm replacement and insert the template again. Adding boilerplate or customer controls twice reports an error rather than duplicating those controls. Costs, customer names, and the footer can be updated repeatedly. The task pane reports missing prerequisites, load failures, and Office API errors.
1. When finished, stop the localhost server and remove the sideloaded add-in.

   ```console
   npm stop
   ```

## Key parts of the sample

### Start with reusable document content

The sample starts by inserting the [bundled OOXML template](./assets/statement-of-work.xml) with [Body.insertOoxml](https://learn.microsoft.com/javascript/api/word/word.body#word-word-body-insertooxml-member(1)), preserving the SOW's rich formatting. The template is a Flat OPC package rather than a ZIP-encoded `.docx` file. Because insertion replaces the active document body, the task pane requires confirmation before proceeding.

With the template in place, [Body.insertParagraph](https://learn.microsoft.com/javascript/api/word/word.body#word-word-body-insertparagraph-member(1)) appends reusable sections with language-independent paragraph styles and a page break. A separate action replaces the inline logo with [InlinePicture.insertInlinePictureFromBase64](https://learn.microsoft.com/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertinlinepicturefrombase64-member(1)), which requires WordApi 1.2 or later.

The [document assembly implementation](./src/assembly/document-assembly.ts) uses `Word.run` to group these operations and `context.sync` to synchronize queued changes. When an action needs to read collection items, it loads them first. For more information, see [Word's application-specific API model](https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model).

### Turn customer text into reusable fields

To make the customer name reusable, the sample searches for `Contoso` and wraps each matching range with [Range.insertContentControl](https://learn.microsoft.com/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1)). Each control receives the tag `customer`. Later updates find the controls with [ContentControlCollection.getByTag](https://learn.microsoft.com/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytag-member(1)), so users can change the name repeatedly without searching for its current text.

The project costs use the same content-control approach, but the sample finds the placeholder by its title, `ProjectCosts`. It then uses [ContentControl.insertFileFromBase64](https://learn.microsoft.com/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-insertfilefrombase64-member(1)) to insert the [bundled costs document](./assets/project-costs.docx) without replacing the rest of the SOW. Unlike the Flat OPC template, this `.docx` file must be passed as Base64-encoded package bytes.

### Highlight amounts and add a footer

After inserting costs, users can highlight USD amounts with [Body.search](https://learn.microsoft.com/javascript/api/word/word.body#word-word-body-search-member(1)) and the wildcard pattern `$[0-9,]@.[0-9]{2}`. The sample formats each matching range and selects the last match. This pattern targets the demo's amounts with two decimal places; it isn't intended to handle every international currency format. For details, see [Use search options in your Word add-in](https://learn.microsoft.com/office/dev/add-ins/word/search-option-guidance).

The footer action uses [Section.getFooter](https://learn.microsoft.com/javascript/api/word/word.section#word-word-section-getfooter-member(1)) to replace the first section's primary footer with `Confidential`. Repeating the action replaces the text rather than appending another copy. Other sections and first-page or even-page footers remain unchanged.

### Configure and insert a project team

The [team builder](./src/assembly/team.ts) lets users configure four assignments, then calculates their average hourly rate and combined cost. Before generating the OOXML table, it checks that hours are finite, greater than zero, and no more than 10,000, and escapes text for use in XML.

To insert the table, the sample uses the Common API's [Document.setSelectedDataAsync](https://learn.microsoft.com/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) with `Office.CoercionType.Ooxml`. The table replaces the current selection. Wrapping the API's callback in a promise routes failures to the same task-pane error handler used by the other actions.

Together, the template and team table demonstrate how both the Word-specific batch API and the Common API can insert rich content. For more information, see [Use Office Open XML in Word add-ins](https://learn.microsoft.com/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml).

## Solution

| Solution | Author(s) |
| --- | --- |
| Assemble a Statement of Work in Word | Microsoft |

## Version history

| Version | Date | Comments |
| --- | --- | --- |
| 1.0 | November 20, 2026 | Initial release into this repository. |

## Additional resources

- [Word add-ins overview](https://learn.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
- [Word JavaScript API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/word/word-api-requirement-sets)
- [Unified manifest for Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview)
- [Microsoft 365 app manifest schema](https://learn.microsoft.com/microsoft-365/extensibility/schema)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2015, 2026 Microsoft Corporation. All rights reserved.

This sample is licensed under the [MIT license](./LICENSE.txt).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-document-assembly" />