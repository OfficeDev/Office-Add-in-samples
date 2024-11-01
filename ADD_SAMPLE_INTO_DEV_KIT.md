# How to add samples to sample gallery of Office Add-ins Development Kit
## Configure the sample for the Dev Kit
The Dev Kit sample gallery uses the [samples-config-v1.json](./samples-config-v1.json) file in this repo's `main` branch. It uses the sample configuration ID find the files, downloads a .zip file of the sample project from this repo, and unzips that file in the designated path.
```
"filterOptions": { // Consumed by the search filters
    "capabilities": [
    ],
    "languages": [
    ],
    "technologies": [
    ]
},
"samples": [ // All config in this config list are presented in the sample gallery
    {
        "id": "Excel-HelloWorld-TaskPane-JS",
        "onboardDate": "2024-07-01",
        "title": "Excel Hello World Task Pane Add-in",
        "shortId": "Excel Hello World",
        "shortDescription": "A simple Excel Task Pane Add-in using JavaScript.",
        "fullDescription": "This sample demonstrates how to create a simple Excel Task Pane Add-in using JavaScript.",
        "types": [
            "Excel"
        ],
        "tags": [
            "JS",
            "Hello World",
            "Excel"
        ],
        "thumbnailPath": "assets/thumbnail.png",
        "gifPath": "assets/sampleDemo.gif",
        "suggested": false
    },
]
```

![The Office Add-ins Dev Kit sample gallery with the parts of the UI labelled with corresponding JSON properties.](assets/config_definition.png)

## Branch
For new samples, please use `dev` as the target branch for pull requests. The `dev` branch will be merged into `main` after testing.

`main` branch is the release branch. All content `main` appears in the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger)'s sample gallery.
![The Office Add-ins Dev Kit sample gallery in VS Code.](assets/sample_gallery.png)

## Testing
TBD

## Check in new sample into Sample Gallery
1. Add the sample project folder is under the **Samples** directory.
    1. Make sure it runs and has been tested on the supported platforms.
    2. Add the **README.md** and **RUN_WITH_EXTENSION.md** files. Use the **README_TEMPLATE.md** under the root directory as a base.
2. Add a new config to the [config file](samples-config-v1.json) with following format.
    1. Fill in these values.
    2. Make sure **id** has identical value with the folder name you just created.
    ![alt text](assets/config_format.png)