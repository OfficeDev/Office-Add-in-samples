# Contribution guidance

If you'd like to contribute to this repository, please read the following guidelines. Contributors are more than welcome to share your learnings with others from centralized location.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Question or problem?

Please do not open GitHub issues for general support questions as the GitHub list should be used for feature requests and bug reports. This way we can more easily track actual issues or bugs from the code and keep the general discussion separate from the actual code.

If you have questions about how to use office.js or the Office developer platform, please post your question on [stackoverflow](https://stackoverflow.com). Tag your question with office-js or outlook-web-addins

## Typos, issues, bugs and contributions

When you submit changes to this PnP repository, please follow these recommendations.

* Always fork repository to your own account for applying modifications
* Do not combine multiple changes to one pull request, please submit for example any samples and documentation updates using separate PRs
* If you are submitting multiple samples, please create specific PR for each of them
* If you are submitting typo or documentation fix, you can combine modifications to single PR where suitable

## Sample naming and structure guidelines

When you are submitting a new sample, use the following guidelines

* You will need to have a README file for your contribution, which is based on the [provided template](../readme-template.md). Please copy this template and update accordingly. The README has to be named as README.md with capital letters.
  * You will need to have a picture of the web part in practice in the README file ("pics or it didn't happen"). Preview image must be located in /assets/ folder in the root for your solution.
* The README template contains specific tracking image as a final entry in the page with img tag by default to https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/readme-template. This is transparent image, which is used to track popularity of individual samples in GitHub.
  * Update the image src element according to correct repository name and folder information. If your sample is for example in samples folder and named as react-todo, src element should be updated as https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/react-todo.
* If you find a similar kind of sample from the existing samples, we would prefer you extend the existing one, rather than submit a new similar sample.
  * When you update existing samples, please update the README as well with new information on provided changes and with your author details.
* When you are submitting a new sample solution, please name the sample solution folder accordingly.
  * Folder should start by identifying JS library used - like "react-", "angular-", "knockout-"
  * If you are not using any specific JS library, please use "js-" as the prefix for your sample, or use "ts-" if your sample uses TypeScript.
  * If your solution is demonstrating multiple technologies, please use functional terms as the name for the solution folder.
* Do not use period/dot in the folder name of the provided sample.

## The Contribution License Agreement

For most contributions, you'll be asked to sign a Contribution License Agreement (CLA). This will happen when you submit a pull request. Microsoft will send a link to the CLA to sign via email. Once you sign the CLA, your pull request can proceed. Read the CLA carefully, because you may need to have your employer sign it.

## Submitting pull requests

Here's a high level process for submitting new samples or updates to existing ones.

1. Sign the Contributor License Agreement.
1. Fork this repository [OfficeDev/PnP-OfficeAddins](https://github.com/OfficeDev/PnP-OfficeAddins) to your GitHub account.
1. Create a new branch off the `master` branch for your fork for the contribution.
1. Include your changes to your branch.
1. Commit your changes using descriptive commit message * These are used to track changes on the repositories for monthly communications
1. Create a pull request in your own fork and target `dev` branch.
1. Fill up the provided PR template with the requested details.

Before you submit your pull request consider the following guidelines:

* Search [GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/pulls) for an open or closed Pull Request
  that relates to your submission. You don't want to duplicate effort.
* Make sure you have a link in your local cloned fork to the [OfficeDev/PnP-OfficeAddins](https://github.com/OfficeDev/PnP-OfficeAddins):

  ```shell
  # check if you have a remote pointing to the Microsoft repo:
  git remote -v

  # if you see a pair of remotes (fetch & pull) that point to https://github.com/OfficeDev/PnP-OfficeAddins, you're ok... otherwise you need to add one

  # add a new remote named "upstream" and point to the Microsoft repo
  git remote add upstream https://github.com/OfficeDev/PnP-OfficeAddins.git
  ```

* Make your changes in a new git branch:

  ```shell
  git checkout -b react-taxonomypicker master
  ```

* Ensure your fork is updated and not behind the upstream **pnp-officeaddins** repo. Refer to these resources for more information on syncing your repo:
  * [GitHub Help: Syncing a Fork](https://help.github.com/articles/syncing-a-fork/)
  * [Keep Your Forked Git Repo Updated with Changes from the Original Upstream Repo](http://www.andrewconnell.com/blog/keep-your-forked-git-repo-updated-with-changes-from-the-original-upstream-repo)
  * For a quick cheat sheet:

    ```shell
    # assuming you are in the folder of your locally cloned fork....
    git checkout master

    # assuming you have a remote named `upstream` pointing official **pnp-officeaddins** repo
    git fetch upstream

    # update your local master to be a mirror of what's in the main repo
    git pull --rebase upstream master

    # switch to your branch where you are working, say "react-taxonomypicker"
    git checkout react-taxonomypicker

    # update your branch to update it's fork point to the current tip of master & put your changes on top of it
    git rebase master
    ```

* Push your branch to GitHub:

  ```shell
  git push origin react-taxonomypicker
  ```

## Merging your existing Github projects with this repository

If the sample you wish to contribute is stored in your own Github repository, you can use the following steps to merge it with the this repository:

* Fork the `pnp-officeaddins` repository from GitHub
* Create a local git repository

    ```shell
    md pnp-officeaddins
    cd pnp-officeaddins
    git init
    ```

* Pull your forked copy of pnp-officeaddins into your local repository

    ```shell
    git remote add origin https://github.com/yourgitaccount/pnp-officeaddins.git
    git pull origin dev
    ```

* Pull your other project from github into the samples folder of your local copy of pnp-officeaddins

    ```shell
    git subtree add --prefix=samples/projectname https://github.com/yourgitaccount/projectname.git master
    ```

* Push the changes up to your forked repository

    ```shell
    git push origin dev
    ```

Thank you for your contribution!

