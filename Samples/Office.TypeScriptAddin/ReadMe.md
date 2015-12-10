# Office TypeScript Add-in #

### Summary ###
This is a sample project extending the Visual Studio 2015 template for an Office task pane add-in with TypeScript and TypeScript type definitions.

It’s a great way to help you ensure the quality and code maintainability of your project. Also, if you’re coming from a more strongly typed programming language (like C#, Java, etc.), TypeScript can be your way into the world of JavaScript.

Read more about this sample at: http://simonjaeger.com/use-typescript-in-a-visual-studio-office-add-in-project

### Applies to ###
-  Office Client (Excel, Word, PowerPoint, Project)
-  Office Online (Excel, Word, PowerPoint)

### Prerequisites ###
None

### Solution ###
Solution | Author(s)
---------|----------
Office.TypeScriptAddin | Simon Jäger (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | December 10th 2015 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Description #
This is a sample project extending the Visual Studio 2015 template for an Office task pane add-in with TypeScript and TypeScript type definitions.

Visual Studio 2015 generates JavaScript from your TypeScript files, that’s what makes it work everywhere. So you should continue to reference *.js files in your HTML. If you browse your file system, you can see that the *.js files (along with *.js.map files) will be located besides the *.ts files.

You can also customize your build phase and debugging experience in a few ways. If you get into the properties of your Visual Studio project and locate the TypeScript Build tab – you can select things such as the ECMAScript version (what’s ECMAScript you may ask – head to: <http://blogs.msdn.com/b/tess/archive/2015/11/12/mastering-asp-net-5-without-growing-a-beard.aspx>), source map location, output properties and more.

![](http://simonjaeger.com/wp-content/uploads/2015/12/tsprops.png)

When your code grows larger, investing in using something like TypeScript really proves to be valuable. At some point you will surely refactor code, hunt bugs and debug – and with TypeScript your life gets a bit easier! You can use Visual Studio to debug your TypeScript code and not have to deal with the plain JavaScript itself (this is what the *.js.map files are good for, mapping between the TypeScript and JavaScript).

## Source Code Files ##

The key source code files in this project are the following:

- `Office.TypeScriptAddinWeb\AddIn\App.ts` - contains the common add-in functionality.
- `Office.TypeScriptAddinWeb\AddIn\Home\Home.ts` - contains the logic for the home page of the add-in.

## Sub level 1.1 ##
Description:
Code snippet:
```C#
string textAndDate = String.Format("Some text to modify - {0}", DateTime.Now.Ticks);
```

## Sub level 1.2 ##

# Doc scenario 2 #

## Sub level 2.1 ##

## Sub level 2.2 ##

### Note: ###

## Sub level 2.3 ##

# Doc scenario 3#

