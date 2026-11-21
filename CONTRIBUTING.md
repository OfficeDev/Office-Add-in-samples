# Contribute

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit <https://cla.microsoft.com>.

When you submit a pull request (PR), a CLA-bot will automatically determine whether you need to provide a CLA and update the PR with a label or comment. Follow the instructions provided by the bot. You'll only need to do this once across all repositories using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Typos, issues, bugs, and contributions

When you submit changes to this repository, please follow these recommendations.

* Always fork the repository to your own account before making changes.
* Don't combine unrelated changes in one PR. Submit a separate PR for each fix, update, or new sample.
* You can combine related typo or documentation fixes in a single PR.

## Sample guidelines

When you submit a new sample, follow these guidelines.

> [!TIP]
> To simplify creating or updating a sample with an AI coding assistant, see [Use the authoring skill (optional)](#use-the-authoring-skill-optional).

### Check for existing samples

If a similar sample already exists in the repository, we'd prefer that you extend it rather than submit a new one.

### Create a README.md file

Create a `README.md` file for your code sample based on the [provided template](/Templates/readme-template.md). Use uppercase letters in the file name: `README.md`.

### Update the tracking image

The README template ends with an `<img>` tag for a transparent tracking image. This image tracks anonymous view counts for individual samples on GitHub. Its default `src` is `https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/name-of-sample`.

Update the image's `src` attribute to match your sample's folder name. For example, if your sample is in the `/Samples/react-todo` folder, set `src` to `https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/react-todo`.

### Name the solution

When you submit a new sample solution, follow these folder naming guidelines.

* Place the solution folder under the `/Samples` folder of this repository.
* Use the format `[product]-[scenario]` for the folder name. For example, name a sample that shows how to create data-bound Excel tables `excel-data-bound-tables`. Use lowercase letters.
* If your solution demonstrates multiple technologies, use a folder name that describes its functionality.
* Don't use periods in the sample's folder name.

### Code hygiene

Follow these guidelines to prepare your code for others to use.

* Clearly attribute code from Stack Overflow or any other source.
* Include code comments that support IDE documentation features, such as IntelliSense or Javadoc.
* Clearly describe classes, methods, parameters, and return values.
* Document exceptions and errors.
* Add remarks about anything special or notable about the code.
* Add comments explaining complex algorithms.
* Follow the code style that's appropriate for the platform and language that your sample uses.
* Test your code.
* Test the UI thoroughly to ensure the changes in your PR don't break existing functionality.

## Use the authoring skill (optional)

Optionally, use the [create-sample skill](.github/skills/create-sample.md) with an AI coding assistant to simplify creating or updating a sample. The skill helps apply this repository's guidelines, but isn't required to contribute.

Open your local clone in your editor and include the skill file in the assistant's context, or explicitly ask the assistant to read and follow it.

Describe the scenario, Office applications, Office.js APIs or features, and target platforms. To update an existing sample, provide the sample's folder and explain what should change and what behavior should be preserved. The skill guides the assistant through checking for related samples, confirming the scope with you, implementing the solution, validating it, and updating its README and repository listings.

**Example prompt for a new sample**

```
Read and follow `.github/skills/create-sample.md` to create an Excel sample that lets users select a range and format it as a table using the Office.js Excel APIs. Target Excel on Windows, Mac, and the web. Check for related samples and confirm the scope with me before implementation.
```

**Example prompt for updating a sample**

```
Read and follow `.github/skills/create-sample.md` to modernize the sample in `Samples/<sample-folder>`. Replace obsolete tooling and UI dependencies while preserving its scenario, demonstrated Office.js APIs, and supported platforms. Report affected incoming references and confirm the scope with me before making changes.
```

Review the generated changes and test the sample in the supported Office clients before submitting a PR. Using the skill doesn't replace this repository's contribution requirements.

## Prepare your changes for a PR

Before you submit your PR, consider the following guidelines.

* Search [GitHub](https://github.com/OfficeDev/Office-Add-in-samples/pulls) for an open or closed PR that relates to your submission. You don't want to duplicate effort.

* Make sure your local clone has an upstream remote pointing to the [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) repository.

  ```shell
  # Check if you have a remote pointing to the Microsoft repository.
  git remote -v

  # If no remote points to the Microsoft repository, add one named "upstream".
  git remote add upstream https://github.com/OfficeDev/Office-Add-in-samples.git
  ```

* Make your changes in a new Git branch.

  ```shell
  git checkout -b [product]-[scenario] main
  ```

* Keep your fork up to date with the upstream **Office-Add-in-samples** repository. See these resources for more information on syncing your repository:

  * [GitHub Help: Syncing a Fork](https://help.github.com/articles/syncing-a-fork/)
  * [Keep Your Forked Git Repo Updated with Changes from the Original Upstream Repo](http://www.andrewconnell.com/blog/keep-your-forked-git-repo-updated-with-changes-from-the-original-upstream-repo)
  * For a quick cheat sheet:

    ```shell
    # Run these commands from your local clone.
    git checkout main

    # Fetch changes from the upstream remote.
    git fetch upstream

    # Update your local main branch with upstream changes.
    git pull --rebase upstream main

    # Switch to your working branch.
    git checkout [product]-[scenario]

    # Rebase your changes onto the latest main branch.
    git rebase main
    ```

* Push your branch to GitHub.

  ```shell
  git push origin [product]-[scenario]
  ```

## Add your code sample to a pull request

Use the following steps to submit a PR for your new code sample.

1. Fork the [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) repository to your GitHub account.
1. Create a contribution branch from the `main` branch in your fork.
1. In the new branch, create a folder for your sample using the naming guidelines in this article.
1. Add your code sample to the folder.
1. Commit the new code using a descriptive commit message.
1. Push the changes to your fork.
1. Create a PR from your contribution branch to the `main` branch of **OfficeDev/Office-Add-in-samples**.
1. Fill out the provided PR template with the requested details.

> [!NOTE]
> If you haven't signed a contributor license agreement (CLA), then you'll automatically be asked to sign a CLA as part of submitting the PR.

When you submit a PR, our team will be notified and will review it. You'll receive notifications from GitHub, and our team may contact you if we need more information. We reserve the right to edit your submission for legal, style, clarity, or other issues.

If you need help keeping your fork in sync with the original repository, see [GitHub Help: Syncing a Fork](https://help.github.com/articles/syncing-a-fork/).

## Merge your existing GitHub projects with this repository

If the sample you want to contribute is stored in your own GitHub repository, use the following steps to merge it into the Office Add-in samples repository.

1. Fork the **Office-Add-in-samples** repository from GitHub.

1. Create a local Git repository.

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

1. Import your other project from GitHub into the `Samples` folder of your local **Office-Add-in-samples** repository.

   ```shell
   git subtree add --prefix=Samples/projectname https://github.com/yourgitaccount/projectname.git main
   ```

1. Push the changes to your fork.

   ```shell
   git push origin main
   ```

Thank you for your contribution!
