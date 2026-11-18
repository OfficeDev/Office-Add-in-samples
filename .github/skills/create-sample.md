---
name: create-sample
description: Create or modernize Office sample solutions from a scenario and specified Office.js APIs. Select scenario-dependent workflows for hosted task panes, function commands, Excel custom functions, or packaged Copilot skills for Excel. Use for sample authoring, legacy migration, starter selection, packaging, and OfficeDev sample README authoring.
---

# Create or modernize an Office Add-ins sample

## Scope

Create a new sample from a user-described scenario and requested Office.js features or APIs, or preserve an existing sample's educational purpose while replacing obsolete tooling and UI dependencies. When a scenario is needed, suggest practical user workflows that highlight the requested features or APIs. Deliver a runnable solution using the scenario's required build, packaging, and installation workflow, plus a README modeled on the OfficeDev samples template.

A task pane is not required for every sample. The scenario may need a task pane, ribbon function commands, both, or another supported runtime. Keep only the UI and runtimes that serve the demonstrated functionality.

Prefer the simplest tooling that satisfies the scenario, not a universal minimal-file layout. Existing samples may legitimately require a framework, backend, authentication, or infrastructure. Preserve those dependencies when they are part of the demonstrated solution.

This developer-facing authoring skill is distinct from a Copilot-in-Excel skill that it may help create.

Do not scaffold or overwrite a populated workspace without approval. Reuse a suitable existing project. Ask before changing the demonstrated features, data source, manifest format, or supported platforms.

## Repository contribution context

This skill is intended for [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples). Before changing a sample:

- Read applicable repository and agent instructions, `CONTRIBUTING.md`, the current README template, and the pull request template if present. Follow destination-repository guidance rather than assuming a generated starter project's conventions apply.
- Limit implementation and cleanup to the selected sample. Never delete or overwrite the repository root, shared files, or neighboring samples as part of cleaning up a generated starter project.
- Before scaffolding, dependency installation, or validation, record whether `dist` and `node_modules` exist within the selected sample. Remove directories created by this run after validation and before completion without requiring a separate cleanup request. Preserve pre-existing directories unless their removal is explicitly approved.
- Place new samples under `Samples/` using the contribution guidelines' lowercase naming convention, such as `<product>-<scenario>`. Follow existing folder conventions for modernization and obtain agreement before renaming a sample.
- Before scaffolding or editing, search existing samples and related open or closed PRs by scenario, Office host, requested APIs or features, and proposed folder name. Inspect relevant matches rather than relying on names alone. Prefer extending or modernizing a suitable sample when the educational purpose overlaps.
- Report the relevant matches and ask the user whether creation or modernization should proceed. When a match exists, offer to extend/modernize that sample, create a distinct sample with an agreed educational difference, or stop. When no match is found, report the search scope and ask whether to proceed with creation or the requested modernization. Use the question tool when available and wait for the user's decision before making changes. Reuse an explicit decision already given after these findings; if the search cannot be completed, disclose the gap rather than claiming no sample exists.
- Check whether existing repository-wide files need an entry added or updated for the sample, such as sample lists, documentation links, metadata used to display or search samples, or dependency-update configuration (for example, Dependabot). Make only the updates required by the repository's existing conventions; do not create a new tracking or listing system.
- Preserve relevant licenses, source attribution, and copyright notices. Use verified author information, never inferred identities or fabricated publication history.
- Keep one coherent sample change per contribution. Do not create a branch, commit, push, or submit a PR unless requested; when preparing a contribution, follow the repository's PR requirements and report any outstanding checklist items.

### Data, permissions, and configuration

- Use synthetic or appropriately licensed public demo data. Do not package personal data, credentials, tokens, tenant-specific identifiers, or account-specific configuration.
- Request only permissions required by the demonstrated APIs. Document authentication and consent requirements when applicable.
- Keep configuration reproducible and separate from secrets. Document required user-supplied values and where to set them; never invent working credentials or hide missing configuration behind fallback success.

### Check cross-repository references

Before modernizing, moving, renaming, or removing an existing sample, search for incoming links and references in:

- [OfficeDev/office-js-docs-pr](https://github.com/OfficeDev/office-js-docs-pr)
- [OfficeDev/office-js-docs-reference](https://github.com/OfficeDev/office-js-docs-reference)
- [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets)

Include other relevant Office Add-ins repositories when identified. Search the sample's canonical URL, repository-relative folder path, sample name, previous names/locations when known, and deep links to source files or README sections. Use authenticated GitHub code search when available; search URL/path variants rather than relying on one exact URL.

Inspect matching content to determine whether the proposed changes would invalidate links, setup instructions, screenshots, code excerpts, or descriptions of sample behavior. A working link does not establish that the referring content remains accurate. Report affected repository/file links and any coordinated updates needed before making changes that affect those references. When references need changes, explicitly advise the user to update each affected file and describe the required correction; do not label these updates optional. For references that remain accurate, state why no update is needed. Do not modify other repositories without approval; approval governs making those edits, not whether to recommend required updates.

Record the repositories, branches, and search terms actually checked. Disclose inaccessible repositories, incomplete searches, and indexing limitations; no matches does not prove that no references exist. For a genuinely new sample with no existing identity or predecessor, report the incoming-reference check as not applicable rather than claiming a search found no references.

## 1. Establish the scenario

For a new sample, use the user's scenario and requested Office.js features or APIs as the specification. Identify the Office host, user actions, input data, expected document changes, and desired UI or runtime. Ask only for missing decisions that materially affect implementation.

When the user provides Office.js features or APIs without a scenario, or asks for scenario ideas, propose two or three distinct, realistic user workflows that make those capabilities central rather than incidental. For each proposal, describe the target user and problem, Office host, input data, key user actions, expected document changes or visible result, and how the requested features or APIs enable the workflow. Favor focused, reproducible demonstrations using synthetic or appropriately licensed public data; identify any additional APIs, authentication, services, or preview dependencies rather than assuming them. Recommend the simplest scenario that clearly teaches the requested capabilities and explain why. If the request is ideation only, stop after presenting the proposals and tradeoffs; otherwise, ask the user to select or refine a scenario before scaffolding or implementation. Do not replace a supplied scenario without agreement.

Verify the requested APIs against current official documentation: signatures, API requirement sets, platform support, and any preview-only status. Map each requested feature to an implementation step and a measurable expected result. Do not silently substitute a different API, add unrelated features, or infer authentication or deployment requirements. Explain unavailable or incompatible APIs and agree on an alternative.

Verify API feasibility before recommending a proposed scenario, and apply the existing-sample search below to the selected scenario before implementation.

Complete the existing-sample search and proceed confirmation described above before implementation. Reuse suitable patterns; do not copy unrelated functionality.

For modernization, read the original sample's README, manifest, runtime code, and any data files. In either workflow, identify:

- Office host and minimum API requirement sets.
- User workflow and the APIs the sample is intended to teach.
- Data source, schema, available dates, and whether data is illustrative or live.
- UI libraries that can be replaced with standard browser controls.
- Required runtime features, including any actual ribbon function commands.

### Select the solution workflow

Classify the solution before choosing a starter, build tools, directory layout, or installation commands:

| Scenario | Workflow |
| --- | --- |
| Hosted task pane, function commands, or both | ATK Office add-in starter; npm/webpack by default; retain scenario-required framework/backend; local HTTPS and sideloading |
| Excel custom functions | Appropriate ATK custom-functions template; preserve required shared or JavaScript-only runtime, metadata, and registration; use its supported build and sideloading workflow |
| Packaged Copilot skill for Excel | Verified skill sample/schema; bundled instructions, resources, and Office.js scripts; ZIP packaging and installation in Copilot in Excel |
| Other solution type | Inspect current official documentation and a matching sample; clarify the workflow before proceeding |

Ask the user when the scenario does not establish which solution type is intended. Do not add webpack, a server, npm startup scripts, or a task pane to a solution that does not need them. Do not apply hosted-add-in cleanup rules to packaged skills or other solution types.

### Single-host and multi-host samples

Support one or multiple Office applications according to the scenario. Distinguish Office applications (Excel, Word, PowerPoint, Outlook) from client platforms (Windows, Mac, web). Do not infer multi-host support from a starter that enables several hosts by default, or reduce an existing multi-host sample to a single host without agreement.

When requirements are incomplete, ask the user to confirm:

- Which Office applications should the sample support? Offer a multi-select when available.
- Which client platforms and minimum Office versions should it target?
- Should each application demonstrate the same workflow, or have different functionality?
- Which Office.js features or APIs and runtimes apply to each application?

Reuse answers already provided or established by the original sample. Ask only about missing or conflicting information before making decisions that affect implementation.

For multi-host samples, distinguish application-specific APIs (such as Excel APIs) from Common APIs that can be used across multiple Office applications. Common APIs are not necessarily supported in every application or on every client platform. Use the [Microsoft Learn Office Add-ins API reference documentation](https://learn.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) as the source of truth for each API's supported applications, requirement sets, client platforms, and limitations. Verify each requested API against every intended application and client platform. Share supported Common API calls, host-independent UI, and helpers where appropriate, and dispatch host-specific behavior using `Office.onReady` and the host identity. Configure manifest scopes, runtime actions, and requirement checks for the intended applications rather than applying one application's requirements to all of them.

Inspect the destination project before modifying it. For new samples, scaffold only when a suitable starter is not already present. Separate requested sample functionality from unused starter demonstrations.

## 2. Select and create the starter

Apply the following CLI scaffolding guidance to hosted add-ins and custom-functions projects. For packaged Copilot skills, use the dedicated workflow below instead; do not assume an ATK Office add-in template is appropriate.

Prefer the **Microsoft 365 Agents Toolkit CLI (ATK)**. It creates the current Toolkit starter without a manual editor wizard. Use the Toolkit only for scaffolding when the delivered sample needs just npm and webpack; remove unneeded provisioning configuration later.

1. Check `atk --version`, `atk list templates`, and `atk new --help` against the current Toolkit documentation. Use a compatible active-LTS Node.js version.
2. If the CLI is missing or an update is required, install the supported current release of `@microsoft/m365agentstoolkit-cli`. Use a prerelease only when required by the Office add-in capability.
3. Confirm the Office host, TypeScript, manifest requirements, project name, and empty destination. Office add-in Toolkit templates are TypeScript-based.
4. Supply the host and project selections through CLI arguments and use `-i false` for non-interactive creation. Respect any prerequisite or consent requirements; do not bypass them.
5. Record the CLI version used for reproducibility and preserve the generated dependency lockfile.

Select an Office Add-in template from the installed CLI's `atk list templates` table based on the requested scenario, host, and runtime. Do not always select the task-pane template. Examples available in ATK 1.1.17 include:

| Template ID | Scenario |
| --- | --- |
| `office-addin-outlook-taskpane` | Outlook task-pane add-in |
| `office-addin-wxpo-taskpane` | Office task-pane starter, also adaptable for function commands |
| `office-addin-excel-cfshortcut` | Excel custom functions with a shared runtime |
| `office-addin-excel-customfunctions` | Excel custom functions with a JavaScript-only runtime |
| `office-addin-sso-naa` | Nested App Auth single sign-on |

Use the current table as the authority; template availability can change. Check template-specific language support and required options before scaffolding.

Run from the intended parent directory, replacing the placeholders with the selected template ID and sample name:

```powershell
$env:ATK_CLI_SKILL = "true"
atk new -c "<office-addin-template-id>" -l typescript -n "<sample-name>" -i false
```

Add only options applicable to that template. ATK 1.1.17 documents `--office-addin-hosts` in `atk new --help`: it selects task-pane Office hosts and defaults to Word, PowerPoint, Outlook, and Excel. For `office-addin-wxpo-taskpane`, append `--office-addin-hosts "<office-hosts>"` when limiting hosts, using a supported host or comma-separated list. This is not a universal option for every Office add-in template. For example, the Nested App Auth template has a separate `--office-addin-naa-host` option.

Always verify these options against the installed CLI's help; the Microsoft Learn introduction may not list every template-specific parameter.

For multi-host scenarios, choose a template that supports the agreed application combination. Pass the explicit host list where the template supports it; do not rely on its all-host default. If no supported template covers the combination or required runtimes, explain the limitation and ask whether to adapt the manifest/runtime configuration or create separate sample projects.

Wait for scaffolding to complete before using the generated folder. Confirm the generated manifest, host, package scripts, and dependencies rather than assuming a particular template layout.

The ATK task-pane template also includes a function-command example. It can be used as a starter for a command-only sample: retain and adapt the command code, its runtime, and ribbon action when applicable, and remove the unused task pane and its open-page action. A function command executes code without requiring a task pane to open. For other runtime requirements, select the appropriate supported template rather than forcing the scenario into a task pane.

Current scaffolding provides the template's selected package versions, not a guarantee that every dependency is the newest published version. Review compatibility and validate the generated dependency set. Do not blindly upgrade all packages to `latest`. Use the package manager for any justified dependency updates and keep `package.json` and `package-lock.json` consistent.

Verify prerequisites from a clean checkout of the sample. Avoid undocumented reliance on globally installed tools, private package registries, or local machine state. Document required CLI tools and installation steps, use public dependencies where available, and retain the appropriate lockfile. If a private registry is essential to the scenario, document the access requirement explicitly without including credentials.

### Manifest decision

- Use the unified JSON manifest by default when its current client, platform, and feature support meets the agreed requirements.
- Use the add-in-only XML manifest when explicitly requested or when a verified requirement cannot be met by the unified manifest and the user approves that choice. Verify specific client, platform, and feature limitations using [Client and platform support](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support). If a required capability is unsupported, explain the gap and ask whether to use the add-in-only manifest, maintain both manifest versions following existing samples in the repository, or revise the requirements. Do not silently drop required support or add a second manifest.
- Do not assume that support for an API implies support for a particular manifest.
- Do not change an existing project's manifest format solely to match a starter default.

### Unified manifest schema version

For every unified Microsoft 365 app manifest, including packaged skill manifests:

- Consult the current [Microsoft 365 app manifest schema reference](https://learn.microsoft.com/microsoft-365/extensibility/schema) during each run. Use the **latest generally available production version** by default; do not hardcode a version in this skill or retain a starter's older schema without reviewing it.
- When the sample uses preview APIs or features, use `manifestVersion: "devPreview"` and the matching `vDevPreview` schema URL. Explain the preview dependency and required client versions or update channels in the README and completion summary.
- Keep `$schema` and `manifestVersion` aligned. For production, use the schema URL published for the selected generally available version. The manifest's separate `version` field is the sample/app version, not its schema version.
- Validate the complete manifest against the chosen schema and check that the target hosts support its features. Do not assume that changing these two fields alone makes a manifest compatible.
- If the requested workflow cannot satisfy these rules or target-host support, surface the conflict and ask for a decision rather than silently using an older production schema or dropping features.

For hosted add-ins, the same scenario code can often serve either format, but runtime and feature support must be checked; npm scripts and webpack must reference the selected manifest consistently. Packaged skills use their documented manifest schema and package structure, not this XML-versus-JSON add-in decision.

If the Toolkit template does not support the required manifest or project configuration, explain the limitation and obtain agreement before changing the approach. A verified official starter repository is the fallback; inspect its license, manifest, dependency requirements, and scripts. Do not hand-create an entire build toolchain or fall back to an editor wizard without explaining the choice.

## 3. Implement the sample in the starter

Follow the selected project's structure and include only the components needed by the scenario. For task panes, keep UI markup, styles, and application logic clearly organized. For function commands, retain the required runtime files, action registration, and handlers, and complete command events as required by the Office host. Do not impose a task-pane layout on samples that do not use one.

Use TypeScript where selected, and retain the necessary Office typings. Use descriptive module names instead of generic host names for scenario logic.

### Modernize dependencies and prefer native capabilities

- Inspect package manifests, lockfiles, CDN references, and actual library usage. Check official maintenance and deprecation notices to identify outdated, deprecated, unsupported, or unused dependencies; age alone does not establish deprecation.
- Prefer native HTML controls, browser APIs, and straightforward JavaScript/TypeScript over external libraries when they satisfy the scenario. For example, use native date inputs for simple date selection, DOM APIs for basic UI interactions, and `fetch` for supported HTTP requests.
- Verify replacements in the selected Office runtime and supported client webviews, including accessibility, localization, validation, and required behavior. Do not assume that a browser API is available in a JavaScript-only runtime.
- Preserve libraries that are part of the sample's educational purpose or provide necessary capabilities that native implementations cannot reasonably match. Do not replace a maintained library with complex custom code merely to reduce dependencies. Where a native replacement is unsuitable, choose a maintained, compatible alternative and explain the choice.
- Keep replacements scoped to the scenario and obtain agreement for behavioral changes. Remove superseded imports, scripts, styles, assets, and configuration only after checking their remaining uses. Use the package manager to update dependencies and the lockfile together; do not blindly upgrade everything to `latest`.
- Validate the affected workflows and document significant replacements and intentional differences in the README and completion summary.

## 4. Apply scenario-specific build and packaging

### Hosted task panes and function commands

Apply this npm/webpack layout and cleanup guidance only to hosted task-pane/function-command solutions that fit it. Preserve required framework, server, and infrastructure files when they serve the scenario. For custom functions, preserve the chosen template's required runtime, metadata, registration, and build files rather than forcing this layout.

For a minimal hosted sample, prefer:

```text
assets/
src/
  taskpane/  (only if needed)
  commands/  (only if needed)
manifest.json OR manifest.xml
package.json
package-lock.json
webpack.config.js
README.md
```

The minimal setup still requires installing dependencies before running the sample. Document `npm install` in the sample directory as a required setup step, followed by `npm start` to start local HTTPS hosting and sideload the add-in. Use `npm stop` to stop debugging. Keep the lockfile aligned with the dependency manifest; `npm install` generates it for a new project. For reproducible validation of an existing project with a matching lockfile, use `npm ci` instead of `npm install`.

Remove Toolkit provisioning files, infrastructure, editor launch files, and separate Babel/TypeScript/lint configuration only when approved and unnecessary for the chosen workflow. Resolve each cleanup target within the selected sample and inspect its purpose before deleting it. Do not delete repository-level configuration or make blanket deletions across unrelated projects.

If Babel's TypeScript preset is supplied inline in webpack, a duplicate Babel configuration can be removed. Babel strips types; preserve a separate type-check command. Removing `tsconfig.json` requires explicit compiler options and entry points in the npm type-check script. Confirm that lint still works without a project-specific configuration.

Treat `dist` and `node_modules` as generated directories. After validation, remove those created by this run, resolving and inspecting each exact path within the selected sample before deletion. Preserve pre-existing directories unless removal is explicitly approved. Leave package manifests and the lockfile intact.

### Packaged Copilot skills for Excel

Use the current [Copilot skill documentation](https://learn.microsoft.com/office/dev/add-ins/excel/excel-copilot-skill) and [Excel Copilot skill sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-copilot-skill) as the authorities for manifest schema, execution model, package structure, and supported versions.

- Confirm that the scenario targets **Copilot in Excel**, not standalone Microsoft 365 Copilot. Check current preview status, Office versions, update channels, account, and access requirements.
- Preserve the documented package hierarchy: manifest and icons, with skill instructions, resources, and scripts under the required skill folders. Do not flatten it to the hosted-add-in root layout.
- Author the packaged `SKILL.md` with clear invocation intent, input/output expectations, and constraints. Provide only the resources and Office.js scripts required by the scenario, following the documented execution environment.
- Bundle all required files in the ZIP. Do not introduce local HTTPS hosting or webpack when the skill is entirely contained in the app package.
- Validate manifest references, skill names, script/resource paths, and ZIP structure. Inspect the archive, not just the source directory.
- Document the applicable installation command, such as `atk install --file-path "<package.zip>" --scope Personal`, after verifying installed CLI support. Installation changes the user's account; perform it only when requested.
- Test through the skill's documented invocation in Copilot in Excel, with a suitable workbook and measurable expected results. Include install, update/repackage, and uninstall instructions in the README.
- Keep deployment packages and required resources unless their deletion is explicitly requested. Do not treat them as disposable hosted build output.

## 5. Author the README using the samples template

Read the current [OfficeDev README template](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Templates/readme-template.md) and [contribution guidelines](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/CONTRIBUTING.md). Model the sample README on their structure rather than retaining the starter's Toolkit instructions:

Every sample README must begin with YAML metadata enclosed by `---` delimiters, before the Markdown title. Refer to existing sample README files in the repository for metadata structure and conventions, adapting the values to the actual sample.

Include `title`, `page_type: sample`, `urlFragment`, `products`, `languages`, `description`, and applicable `extensions` metadata such as `contentType: samples`, `technologies`, and `createdDate`, following current repository discovery conventions. Match the metadata title to the descriptive heading and use the agreed sample identifier for `urlFragment`. List only supported products and languages actually used. Preserve a verified existing creation date; for a new sample, use its actual creation date in the repository's expected format. Do not invent historical dates or copy identities and values from the example. Ask for any required metadata that cannot be established.

1. **Descriptive title** and a short introduction with demonstrated tasks.
2. **Applies to**: actual Office hosts/platforms, distinguishing API compatibility, manifest support, and tested platforms.
3. **Prerequisites**: Node.js, npm, Office, and verified minimum versions.
4. **Configure the sample**: include only when configuration is necessary.
5. **Run the sample**: numbered setup and usage steps appropriate to the selected solution. Hosted add-ins use `npm install`, `npm start`, their actual ribbon labels or add-in commands, and `npm run stop`; packaged skills use package installation, workbook setup, Copilot invocation, and uninstall instructions.
6. **Key parts of the sample**: focus on what developers can learn about the Office.js APIs and Office platform features demonstrated.
   - Name the actual APIs or features used and explain their role in enabling the scenario and their observable result. Where possible, link API names to their official Microsoft Learn API references, preferring verified member links or the containing type's reference page. Also search Microsoft Learn for relevant conceptual articles that explain the demonstrated capabilities, workflows, and recommended patterns; include useful, verified references alongside the API links. Link platform features to their official documentation and include relevant source links alongside the documentation, not instead of it.
   - Write a connected explanation of the scenario, not an API inventory. Start with a brief overview of the user workflow, then order capability-focused subsections to follow that workflow where appropriate. Lead each explanation with the user's action or intended result, introduce the APIs that enable it, and connect it naturally to the next step. Use short paragraphs and meaningful transitions; integrate reference links without letting consecutive API names dominate the prose. Explain shared runtime patterns in context rather than repeating them under every capability.
   - Organize by capability rather than files, UI controls, frameworks, or build tooling. Avoid a file-by-file implementation walkthrough or details about npm, webpack, CSS, and generic data loading unless those are themselves the sample's educational purpose. Put setup, data configuration, and tooling details in the appropriate run/configuration sections.
7. **Solution**: solution and author table, preserving original attribution and using only verified author information.
8. **Version history**: accurate version/date/comments table; do not invent releases.

For new samples, document the supplied scenario and the Office.js features or APIs demonstrated. For modernization and migration efforts, do not describe the project as a migration. Additionally, do not link to the original sample or explain intentional differences.

Use relative links for sample source files and assets in the README; use official URLs for external documentation.

Retain useful existing images. If capture is not possible, report it as an item to confirm instead of creating a broken image reference or implying that a screenshot exists.

Include required feedback links, copyright/attribution, and accurate version history. For a contribution to OfficeDev/Office-Add-in-samples, append the following tracking image as the README's final entry, as required by [CONTRIBUTING.md](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/CONTRIBUTING.md#update-tracking-image):

```html
<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/<sample-folder-name>" />
```

Replace `<sample-folder-name>` with the actual sample folder identifier, consistent with README `urlFragment` and `assets/sample.json` `name`. This transparent image tracks anonymous sample views. Preserve or correct an existing tracking image rather than adding a duplicate, and leave no content after it. Treat required repository elements as contribution requirements, not optional additions. If an unpublished sample has no agreed folder name, ask for it rather than leaving a placeholder or inventing an identity.

## 6. Create or update sample discovery metadata

Create or update the sample's `assets/sample.json` discovery metadata. Refer to existing `assets/sample.json` files in the repository for structure and conventions, adapting the values to the actual sample.

- Use a top-level array containing the sample entry. Include `name`, `reponame`, `source`, `title`, `shortDescription`, `url`, `longDescription`, `creationDateTime`, `updateDateTime`, `products`, `metadata`, `thumbnails`, `authors`, and `references`, following current repository conventions.
- Use the agreed sample folder identifier for `name` and its canonical repository location for `url`. Keep the identity, descriptions, supported products, and language metadata consistent with the README, respecting each format's vocabulary.
- Use verified creation and update dates in `YYYY-MM-DD` format. Preserve original attribution and include only verified author details; never copy authors or dates from other samples without evidence.
- Include useful, verified conceptual and API documentation in `references`. Reference only actual thumbnail assets; an empty `thumbnails` array is appropriate when none exist.
- Treat this as repository discovery metadata, not imported demo data or a runtime resource. Preserve it during asset cleanup; do not add hosting or package references solely because it resides under `assets`.

## 7. Validate and finish

- Run the smallest applicable validation for the selected workflow: type-check/lint/build/manifest checks for hosted add-ins, runtime/metadata checks for custom functions, or manifest/resource/archive checks for packaged skills. Do not run nonexistent npm scripts.
- Confirm that unified manifests use the latest production schema, or `devPreview` for preview APIs/features, with matching `$schema` and `manifestVersion`.
- Verify the exact feature behavior, relevant input boundaries, and output shape.
- For new samples, verify each requested Office.js feature/API is used as agreed and produces its specified result.
- For multi-host samples, verify each supported Office application's manifest configuration, runtime routing, API requirements, and intended workflow separately.
- Verify required assets, manifest references, data files, runtime resources, and the delivered build output or package.
- Verify installation and documented startup or package installation from a clean checkout where feasible, not only from an already configured developer machine.
- Check README discovery metadata and parse `assets/sample.json` as JSON. Verify their identities, descriptions, dates, products, languages, and references are consistent and use the respective repository formats. Confirm the tracking image uses the agreed sample folder identifier and is the README's final entry. Check relative links, required images/feedback elements, source attribution, and any necessary sample-catalog or dependency-automation registration.
- Reassess the cross-repository reference findings against the completed changes, including link targets and the accuracy of referring content. Explicitly advise the user to make any required follow-up updates, identifying each affected repository/file and the correction needed. Treat updates not performed because approval was not granted as outstanding required follow-up, not optional work; do not imply that another repository was updated unless it was explicitly authorized and completed.
- Test in the intended Office host where available; distinguish host testing from mocks or build checks.
- Reconcile any directly dependent scripts after removing files.
- After validation, remove `dist` and `node_modules` directories created by this run within the selected sample. Resolve and inspect each exact cleanup path; preserve pre-existing directories unless removal is explicitly approved, and preserve required deployment packages and resources. Verify the run-created directories are absent before completion; report any cleanup failures or explicitly requested retention.
- End every run with a concise completion summary containing:
  - **What was done:** the scenario implemented or modernized, selected solution workflow, Office.js features or APIs used, runtime choices, and significant project/package/documentation changes.
  - **Validation:** checks performed and their results, distinguishing builds or mocks from actual Office-host testing; generated-directory cleanup results and any retained pre-existing directories.
  - **Link/reference findings:** repositories, branches, and search terms checked; links to matching files; impact on links and referring content; and coordinated updates completed or still required. For each outstanding required update, advise the user to update the linked file and describe what must change. Do not describe required updates as optional merely because cross-repository edits need approval. Explain when matching references need no change. Explicitly state when no matches were found within the searched scope, the check was not applicable, or access/search limitations prevented completion.
  - **Items to confirm:** outstanding decisions, assumptions, prerequisites, or unverified behavior. State explicitly when there are no outstanding items.
  - **Testing reminder:** test the sample in every supported Office application and host/platform listed in the README, using versions that satisfy both manifest and API requirements. Include the relevant desktop and web clients where supported, and Copilot-in-Excel invocation for packaged skills; do not imply that testing one host validates the others.

Do not publish, deploy infrastructure, or launch a new Office session unless requested.

## 8. Evaluate changes to this authoring skill

Before contributing changes to the skill itself, evaluate its instructions against representative scenarios. Use bounded dry runs for routing and file-retention decisions; run actual scaffolding, installation, or Office testing only when authorized. Do not create five sample projects merely to validate a documentation edit.

| Scenario | Expected decisions to verify |
| --- | --- |
| New task-pane sample | Choose a suitable ATK template; keep required UI, hosted runtime, and sample metadata |
| Command-only sample | Keep command registration/runtime and ribbon action; remove only unused task-pane content |
| Multi-host sample | Clarify applications/platforms; select explicit hosts; verify each application's APIs and workflow |
| Excel custom-functions sample | Preserve runtime, function metadata, registration, and required build files |
| Packaged Copilot skill for Excel | Preserve package hierarchy; use ZIP/install workflow; do not introduce unnecessary webpack or hosting |

For each scenario, check that the skill asks for missing decisions, scopes changes to one sample, preserves required files, produces the correct README/run instructions, and ends with confirmation items and a supported-host testing reminder. Report which evaluations were performed and which remain unverified; do not claim runtime support based on a dry run.

Also dry-run the reference-impact instructions for a genuinely new sample, an existing sample with unchanged URLs but changed behavior, a moved or renamed sample with deep links, an unchanged and still-accurate reference, and an inaccessible reference repository. Verify that findings appear in the completion summary, required updates include explicit advice to the user with file links and corrections rather than being called optional, unaffected references are explained, search gaps are disclosed, and cross-repository edits still require approval.

## References

- [Microsoft 365 Agents Toolkit](https://learn.microsoft.com/microsoftteams/platform/toolkit/agents-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit source and CLI](https://github.com/OfficeDev/microsoft-365-agents-toolkit)
- [Manifest overview](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
- [Microsoft 365 app manifest schema reference](https://learn.microsoft.com/microsoft-365/extensibility/schema)
- [Unified manifest platform support](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support)
- [OfficeDev samples README template](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Templates/readme-template.md)
- [OfficeDev samples contribution guidelines](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/CONTRIBUTING.md)
- [Copilot skills for Excel](https://learn.microsoft.com/office/dev/add-ins/excel/excel-copilot-skill)
- [Excel Copilot skill sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-copilot-skill)
