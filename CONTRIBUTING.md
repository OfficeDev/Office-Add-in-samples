# Contribute

This project welcomes contributions and suggestions. Most contributions require you to
agree to a Contributor License Agreement (CLA) declaring that you have the right to,
and actually do, grant us the rights to use your contribution. For details, visit
https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need
to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the
instructions provided by the bot. You will only need to do this once across all repositories using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)
or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Typos, issues, bugs and contributions

When you submit changes to this repository, please follow these recommendations.

* Always fork the repository to your own account for applying modifications.
* Don't combine multiple changes in one pull request. Please submit separate PRs for each fix, update, or new sample.
* If you are submitting a typo or documentation fix, you can combine modifications in a single PR where suitable.

## Sample guidelines

When you are submitting a new sample, use the following guidelines.

### Check for existing samples
If you find a similar sample that already exists in the repository, we would prefer that you extend the existing one, rather than submit a new similar sample.

### Create a README.md file
Create a README.md file for your code sample, and base it on the [provided template](/Templates/readme-template.md). The README must be named README.md with capital letters.

### Update tracking image
The README template contains specific tracking image as a final entry in the page with img tag by default to https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/name-of-sample. This is transparent image, which is used to track anonymous view counts of individual samples in GitHub.

Update the image `src` element according to correct repository name and folder information. For example, if your sample is **/samples** folder and named as "react-todo", `src` element should be updated as https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/react-todo.

### Name the solution
When you are submitting a new sample solution, please name the sample solution folder accordingly.

* The solution folder should be under the **/Samples** folder of this repository.
* Folder name should be in the format *[product]-[scenario]-[platform]*. For example, a sample that shows how to create data-bound Excel tables using React, should be named "excel-data-bound-tables-react". The names are all lowercase.
* If your solution is demonstrating multiple technologies, please use functional terms as the name for the solution folder.
* Do not use a period in the folder name of the provided sample.

### Add your sample to the Office Add-ins Development Kit
To have your sample be included in Office Add-ins VS Code extension, follow the instructions in [How to add samples to sample gallery of Office Add-ins Development Kit](./ADD_SAMPLE_INTO_DEV_KIT.md).

## Add your code sample to a pull request

Use the following steps to submit a pull request for your new code sample.

1. Fork this repository [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) to your GitHub account.
2. Create a new branch off the `main` branch for your fork for the contribution.
3. In the new branch, create the folder for your new sample using the previous naming guidelines.
4. Add your code sample to the folder.
5. Commit the new code using descriptive commit message. Commit messages are used to track changes on the repositories for monthly communications.
6. Push the changes up to your fork.
7. Create a pull request in your own fork and target the `main` branch on OfficeDev org.
8. Fill up the provided PR template with the requested details.

> **Note:** If you haven't signed a contributor license agreement (CLA), then you will automatically be asked to sign a CLA as part of submitting the PR.

When you submit your changes, via a pull request, our team will be notified and will review your pull request. You will receive notifications about your pull request from GitHub; you may also be notified by someone from our team if we need more information. We reserve the right to edit your submission for legal, style, clarity, or other issues.

If you need help keeping your fork in sync with the original repository, see [GitHub Help: Syncing a Fork](https://help.github.com/articles/syncing-a-fork/).

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

## Additional guidelines

Before you submit your pull request, consider the following guidelines.

* Search [GitHub](https://github.com/OfficeDev/Office-Add-in-samples/pulls) for an open or closed Pull Request
  that relates to your submission. You don't want to duplicate effort.

* Make sure you have a link in your local cloned fork to the [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) repository.

  ```shell
  # check if you have a remote pointing to the Microsoft repository
  git remote -v

  # if you see a pair of remotes (fetch & pull) that point to https://github.com/OfficeDev/Office-Add-in-samples, you're ok... otherwise you need to add one

  # add a new remote named "upstream" and point to the Microsoft repository
  git remote add upstream https://github.com/OfficeDev/Office-Add-in-samples.git
  ```

* Make your changes in a new git branch.

  ```shell
  git checkout -b react-taxonomypicker main
  ```

* Ensure your fork is updated and not behind the upstream **Office-Add-in-samples** repository. Refer to these resources for more information on syncing your repository:
  * [GitHub Help: Syncing a Fork](https://help.github.com/articles/syncing-a-fork/)
  * [Keep Your Forked Git Repo Updated with Changes from the Original Upstream Repo](http://www.andrewconnell.com/blog/keep-your-forked-git-repo-updated-with-changes-from-the-original-upstream-repo)
  * For a quick cheat sheet:

    ```shell
    # assuming you are in the folder of your locally cloned fork....
    git checkout main

    # assuming you have a remote named `upstream` pointing to the official Office-Add-in-samples repository
    git fetch upstream

    # update your local main to be a mirror of what's in the main repository
    git pull --rebase upstream main

    # switch to your branch where you are working, say "react-taxonomypicker"
    git checkout react-taxonomypicker

    # update your branch to update it's fork point to the current tip of main & put your changes on top of it
    git rebase main
    ```

* Push your branch to GitHub.

  ```shell
  git push origin react-taxonomypicker
  ```

## Merging your existing Github projects with this repository

If the sample you wish to contribute is stored in your own Github repository, you can use the following steps to merge it with the this repository.

1. Fork the **Office-Add-in-samples** repository from GitHub.

1. Create a local git repository.

    ```shell
    md Office-Add-in-samples
    cd Office-Add-in-samples
    git init
    ```

1. Pull your forked copy of **Office-Add-in-samples** into your local repository.

    ```shell
    git remote add origin https://github.com/yourgitaccount/Office-Add-in-samples.git
    git pull origin main
    ```

1. Pull your other project from Github into the samples folder of your local copy of the Office-Add-in-samples repository.

    ```shell
    git subtree add --prefix=samples/projectname https://github.com/yourgitaccount/projectname.git main
    ```

1. Push the changes up to your forked repository

    ```shell
    git push origin main
    ```

Thank you for your contribution!
