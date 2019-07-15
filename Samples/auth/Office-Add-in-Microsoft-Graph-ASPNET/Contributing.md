# Contribute to this code sample 

Thank you for your interest in our sample

* [Ways to contribute](#ways-to-contribute)
* [To contribute using Git](#to-contribute-using-git)
* [Contribute code](#contribute-code)
* [FAQ](#faq)
* [More resources](#more-resources)

## Ways to contribute

Here are some ways you can contribute to this sample:

* Add better comments to the sample code.
* Fix issues opened in GitHub against this sample.
* Add a new feature to the sample.

We want your contributions. Help the developer community by improving this sample. 
Contributions can include Bug fixes, new features, and better code documentation. 
Submit code comment contributions where you want a better explanation of what's going on. 
See a good example of [code commenting](https://github.com/OfficeDev/O365-Android-Microsoft-Graph-Connect/blob/master/app/src/main/java/com/microsoft/office365/connectmicrosoftgraph/AuthenticationManager.java).

Another great way to improve the sample in this repository is to take on some of the open issues filed against the repository. You may have a solution to an bug in the sample code that hasn't been addressed. Fix the issue and then create a pull request following our [Contribute code](#contribute-code) guidance. 

If you want to add a new feature to the sample, be sure you have the agreement of the repository owner before writing the code. Start by opening an issue in the repository. Use the new issue to propose the feature. The repository owner will respond and will usually ask you for more information. When the owner agrees to take the new feature, code it and submit a pull request.

## To contribute using Git
For most contributions, you'll be asked to sign a Contribution License Agreement (CLA). For those contributions that need it, The Office 365 organization on GitHub will send a link to the CLA that we want you to sign via email. 
By signing the CLA, you acknowledge the rights of the GitHub community to use any code that you submit. The intellectual property represented by the code contribution is licensed for use by Microsoft open source projects.

If Office 365 emails a CLA to you, you need to sign it before you can contribute large submissions to a project. You only need to complete and submit it once. 
Read the CLA carefully. You may need to have your employer sign it.

Signing the CLA does not grant you rights to commit to the main repository, but it does mean that the Office Developer and Office Developer Content Publishing teams will be able to review and approve your contributions. You will be credited for your submissions.

Pull requests are typically reviewed within 10 business days.

## Use GitHub, Git, and this repository

**Note:** Most of the information in this section can be found in [GitHub Help] articles.  If you're familiar with Git and GitHub, skip to the **Contribute code** section for the specifics of the code contributions for this repository.

### To set up your fork of the repository

1.	Set up a GitHub account so you can contribute to this project. If you haven't done this, go to [GitHub](https://github.com/join) and do it now.
2.	Install Git on your computer. Follow the steps in the [Setting up Git Tutorial] [Set Up Git].
3.	Create your own fork of this repository. To do this, at the top of the page, choose the **Fork** button.
4.	Copy your fork to your computer. To do this, open Git Bash. At the command prompt enter:

		git clone https://github.com/<your user name>/<repo name>.git

	Next, create a reference to the root repository by entering these commands:

		cd <repo name>
		git remote add upstream https://github.com/OfficeDev/<repo name>.git
		git fetch upstream

Congratulations! You've now set up your repository. You won't need to repeat these steps again.

## Contribute code

To make the contribution process as seamless as possible, follow these steps.

### To contribute code

1. Create a new branch.
2. Add new code or modify existing code.
3. Submit a pull request to the main repository.
4. Await notification of acceptance and merge.
5. Delete the branch.


### To create a new branch

1.	Open Git Bash.
2.	At the Git Bash command prompt, type `git pull upstream master:<new branch name>`. This creates a new branch locally that is copied from the latest OfficeDev master branch.
3.	At the Git Bash command prompt, type `git push origin <new branch name>`. This alerts GitHub to the new branch. You should now see the new branch in your fork of the repository on GitHub.
4.	At the Git Bash command prompt, type `git checkout <new branch name>` to switch to your new branch.

### Add new code or modify existing code

Navigate to the repository on your computer. On a Windows  PC, the repository files are in `C:\Users\<yourusername>\<repo name>`.

Use the IDE of your choice to modify and build the sample. Once you have completed your change, commented your code, and test, check the code
into the remote branch on GitHub.

#### Code contribution checklist
Be sure to satisfy all of the requirements in the following list before submitting a pull request:


-  Follow the code style found in the cloned repository code. Our Android code follows the style conventions found in the [Code Style for Contributors](https://source.android.com/source/code-style.html) guide. 
- Code must be tested.
- Test the sample UI thoroughly to be sure nothing has been broken by your change.
- Keep the size of your code change reasonable. If the repository owner cannot review your code change in 4 hours or less, your pull request may not be reviewed and approved quickly.
- Avoid unnecessary changes to cloned or forked code. The reviewer will use a tool to find the differences between your code and the original code. Whitespace changes are called out along with your code. Be sure your changes will help improve the content.

### Push your code to the remote GitHub branch
The files in `C:\Users\<yourusername>\<repo name>` are a working copy of the new branch that you created in your local repository. Changing anything in this folder doesn't affect the local repository until you commit a change. To commit a change to the local repository, type the following commands in GitBash:

	git add .
	git commit -v -a -m "<Describe the changes made in this commit>"

The `add` command adds your changes to a staging area in preparation for committing them to the repository. The period after the `add` command specifies that you want to stage all of the files that you added or modified, checking subfolders recursively. (If you don't want to commit all of the changes, you can add specific files. You can also undo a commit. For help, type `git add -help` or `git status`.)

The `commit` command applies the staged changes to the repository. The switch `-m` means you are providing the commit comment in the command line. The -v  and -a switches can be omitted. The -v switch is for verbose output from the command, and -a does what you already did with the add command.

You can commit multiple times while you are doing your work, or you can commit once when you're done.

### Submit a pull request to the master repository

When you're finished with your work and are ready to have it merged into the master repository, follow these steps.

#### To submit a pull request to the master repository

1.	In the Git Bash command prompt, type `git push origin <new branch name>`. In your local repository, `origin` refers to your GitHub repository that you cloned the local repository from. This command pushes the current state of your new branch, including all commits made in the previous steps, to your GitHub fork.
2.	On the GitHub site, navigate in your fork to the new branch.
3.	Choose the **Pull Request** button at the top of the page.
4.	Verify the Base branch is `OfficeDev/<repo name>@master` and the Head branch is `<your username>/<repo name>@<branch name>`.
5.	Choose the **Update Commit Range** button.
6.	Add a title to your pull request, and describe all the changes you're making.
7.	Submit the pull request.

One of the site administrators will process your pull request. Your pull request will surface on the `OfficeDev/<repo name>` site under Issues. When the pull request is accepted, the issue will be resolved.

### Repository owner code review
The owner of the repository will review your pull request to be sure that all requirements are met. If the reviewer
finds any issues, she will communicate with you and ask you to address them and then submit a new pull request. If your pull 
request is accepted, then the repository owner will tell you that your pull request is to be merged.

### Create a new branch after merge

After a branch is successfully merged (that is, your pull request is accepted), don't continue working in that local branch. This can lead to merge conflicts if you submit another pull request. To do another update, create a new local branch from the successfully merged upstream branch, and then delete your initial local branch.

For example, if your local branch X was successfully merged into the OfficeDev/O365-Android-Microsoft-Graph-Connect master branch and you want to make additional updates to the code that was merged. Create a new local branch, X2, from the OfficeDev/O365-Android-Microsoft-Graph-Connect branch. To do this, open GitBash and execute the following commands:

	cd <repo name>
	git pull upstream master:X2
	git push origin X2

You now have local copies (in a new local branch) of the work that you submitted in branch X. The X2 branch also contains all the work other developers have merged, so if your work depends on others' work (for example, a base class), it is available in the new branch. You can verify that your previous work (and others' work) is in the branch by checking out the new branch...

	git checkout X2

...and verifying the code. (The `checkout` command updates the files in `C:\Users\<yourusername>\O365-Android-Microsoft-Graph-Connect` to the current state of the X2 branch.) Once you check out the new branch, you can make updates to the code and commit them as usual. However, to avoid working in the merged branch (X) by mistake, it's best to delete it (see the following **Delete a branch** section).

### Delete a branch

Once your changes are successfully merged into the main repository, delete the branch you used because you no longer need it.  Any additional work should be done in a new branch.  

#### To delete a branch

1.	In the Git Bash command prompt, type `git checkout master`. This ensures that you aren't in the branch to be deleted (which isn't allowed).
2.	Next, at the command prompt, type `git branch -d <branch name>`. This deletes the branch on your computer only if it has been successfully merged to the upstream repository. (You can override this behavior with the `-D` flag, but first be sure you want to do this.)
3.	Finally, type `git push origin :<branch name>` at the command prompt (a space before the colon and no space after it).  This will delete the branch on your github fork.  

Congratulations, you have successfully contributed to the sample app!


## FAQ

### How do I get a GitHub account?

Fill out the form at [Join GitHub](https://github.com/join) to open a free GitHub account. 

### Where do I get a Contributor's License Agreement? 

You will automatically be sent a notice that you need to sign the Contributor's License Agreement (CLA) if your pull request requires one. 

As a community member, **you must sign the CLA before you can contribute large submissions to this project**. You only need complete and submit the CLA document once. Carefully review the document. You may be required to have your employer sign the document.

### What happens with my contributions?

When you submit your changes, via a pull request, our team will be notified and will review your pull request. You will receive notifications about your pull request from GitHub; you may also be notified by someone from our team if we need more information. If your pull request is approved, we'll update the documentation on GitHub and on MSDN. We reserve the right to edit your submission for legal, style, clarity, or other issues.

### Who approves pull requests?

The owner of the sample repository approves pull requests. 

### How soon will I get a response about my change request?

Pull requests are typically reviewed within 10 business days.


## More resources

* To learn more about Markdown, go to the Git creator's site [Daring Fireball].
* To learn more about using Git and GitHub, check out the [GitHub Help section] [GitHub Help].

[GitHub Home]: http://github.com
[GitHub Help]: http://help.github.com/
[Set Up Git]: http://help.github.com/win-set-up-git/
[Markdown Home]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
