# Contribute a new code sample

If you'd like to contribute a new code sample to this repository, use the following guidelines and instructions. Sharing your code is a great way to help other developers in the community!

## Sample naming and structure guidelines

When you are submitting a new sample, use the following guidelines.

**Check for existing samples**
If you find a similar sample that already exists in the repo, we would prefer that you extend the existing one, rather than submit a new similar sample. For more information, see [Contributing guidelines](CONTRIBUTING.md).

**Create a README.md file**
Create a README.md file for your code sample, and base it on the [provided template](readme-template.md). The README must be named README.md with capital letters.

**Update tracking image**
The README template contains specific tracking image as a final entry in the page with img tag by default to https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/readme-template. This is transparent image, which is used to track popularity of individual samples in GitHub.

Update the image src element according to correct repository name and folder information. If your sample is for example in samples folder and named as react-todo, src element should be updated as https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/react-todo.

**Name the solution**
When you are submitting a new sample solution, please name the sample solution folder accordingly.
* The solution folder should be under the **/Samples** folder of this repo.
* The folder name should start by identifying JS library used - like "react-", "angular-", "knockout-"
* If you are not using any specific JS library, please use "js-" as the prefix for your sample, or use "ts-" if your sample uses TypeScript.
* If your solution is demonstrating multiple technologies, please use functional terms as the name for the solution folder.
* Do not use period/dot in the folder name of the provided sample.

For example, if your solution shows how to create a task pane in any Office app using JavaScript, you could name the folder **/Samples/office-js-task-pane**.

## Add your code sample to a pull request

Use the following steps to submit a pull request for your new code sample.

1. Fork this repository [OfficeDev/PnP-OfficeAddins](https://github.com/OfficeDev/PnP-OfficeAddins) to your GitHub account.
2. Create a new branch off the `main` branch for your fork for the contribution.
3. In the new branch, create the folder for your new sample using the previous naming guidelines.
4. Add your code sample to the folder.
5. Commit the new code using descriptive commit message. Commit messages are used to track changes on the repositories for monthly communications.
6. Push the changes up to your fork.
7. Create a pull request in your own fork and target the `main` branch on OfficeDev org.
8. Fill up the provided PR template with the requested details.

> **Note:** If you haven't signed a contributor license agreement (CLA), then you will automatically be asked to sign a CLA as part of submitting the PR.

When you submit your changes, via a pull request, our team will be notified and will review your pull request. You will receive notifications about your pull request from GitHub; you may also be notified by someone from our team if we need more information. We reserve the right to edit your submission for legal, style, clarity, or other issues.

If you need help keeping your fork in sync with the original repo, see [GitHub Help: Syncing a Fork](https://help.github.com/articles/syncing-a-fork/).

## Code hygiene

Follow these guidelines to be sure your code is ready for the world.

* Code added from Stack Overflow, or any other source, is clearly attributed.
* Any code that has associated documentation displayed in an IDE (such as IntelliSense, or JavaDocs) has code comments.
* Classes, methods, parameters, and return values have clear descriptions.
* Exceptions and errors are documented.
* Remarks exist for anything special or notable about the code.
* Sections of code that have complex algorithms have appropriate comments describing what they do.
* Follow the code style that is appropriate for the platform and language that your sample uses.
* Test your code.
* Test the UI thoroughly to be sure nothing was broken during the process of moving code into the pull request.

Thank you for your contribution!
