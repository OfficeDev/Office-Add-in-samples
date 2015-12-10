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

Below is the App.ts file included in the sample, it's a rewritten version of the default App.js file that comes along with the Visual Studio Office add-in template. 

```TS
/* Common app functionality */

class App {
    private initialized: boolean = false;

    // Common initialization function (to be called from each page)
    initialize() {
        $('body').append(
            '<div id="notification-message">' +
            '<div class="padding">' +
            '<div id="notification-message-close"></div>' +
            '<div id="notification-message-header"></div>' +
            '<div id="notification-message-body"></div>' +
            '</div>' +
            '</div>');

        // After initialization, enable the showNotification function
        this.initialized = true;
    }

    // Notification function, enabled after initialization
    showNotification(header: string, text: string) {
        if (!this.initialized) {
            console.log('Add-in has not yet been initialized.');
            return;
        }

        $('#notification-message-header').text(header);
        $('#notification-message-body').text(text);
        $('#notification-message').slideDown('fast');
    }
}

var app = new App();
```
Below is the generatds JavaScript code, compiled from the TypeScript file as App.js. Included is also a source mapping reference - used when debugging (mapping the two files).

```JS
/* Common app functionality */
var App = (function () {
    function App() {
        this.initialized = false;
    }
    // Common initialization function (to be called from each page)
    App.prototype.initialize = function () {
        $('body').append('<div id="notification-message">' +
            '<div class="padding">' +
            '<div id="notification-message-close"></div>' +
            '<div id="notification-message-header"></div>' +
            '<div id="notification-message-body"></div>' +
            '</div>' +
            '</div>');
        // After initialization, enable the showNotification function
        this.initialized = true;
    };
    // Notification function, enabled after initialization
    App.prototype.showNotification = function (header, text) {
        if (!this.initialized) {
            console.log('Add-in has not yet been initialized.');
            return;
        }
        $('#notification-message-header').text(header);
        $('#notification-message-body').text(text);
        $('#notification-message').slideDown('fast');
    };
    return App;
})();
var app = new App();
//# sourceMappingURL=App.js.map
```

You can also customize your build phase and debugging experience in a few ways. If you get into the properties of your Visual Studio project and locate the TypeScript Build tab – you can select things such as the ECMAScript version (what’s ECMAScript you may ask – head to: <http://blogs.msdn.com/b/tess/archive/2015/11/12/mastering-asp-net-5-without-growing-a-beard.aspx>), source map location, output properties and more.

![](http://simonjaeger.com/wp-content/uploads/2015/12/tsprops.png)

When your code grows larger, investing in using something like TypeScript really proves to be valuable. At some point you will surely refactor code, hunt bugs and debug – and with TypeScript your life gets a bit easier! You can use Visual Studio to debug your TypeScript code and not have to deal with the plain JavaScript itself (this is what the *.js.map files are good for, mapping between the TypeScript and JavaScript).

## Source Code Files ##

The key source code files in this project are the following:

- `Office.TypeScriptAddinWeb\AddIn\App.ts` - contains the common add-in functionality.
- `Office.TypeScriptAddinWeb\AddIn\Home\Home.ts` - contains the logic for the home page of the add-in.

## More Resouces ##
Find more information and resources at:
- Learn more about building for Office and Office 365 at: <http://dev.office.com/>
- Get started with TypeScript at: <http://www.typescriptlang.org/>
- Find lots of TypeScript type definitions at: <https://github.com/DefinitelyTyped/DefinitelyTyped>
- Read more about this sample at: <http://simonjaeger.com/use-typescript-in-a-visual-studio-office-add-in-project>
