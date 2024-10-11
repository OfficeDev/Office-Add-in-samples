# Contributing
## Branch
Default branch is `dev`. You should set `dev` as the target branch of PRs. `dev` will be merged into `main` regularly after testing.

`main` branch is the release branch, which means all content in this branch will be directly consumed by the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger)'s sample gallery:
![alt text](assets/sample_gallery.png)

## Testing
TBD

## Add new sample
1. Add a new project folder under the **Samples** directory.
    1. Make sure it's runnable and well tested.
    2. Add README.md and RUN_WITH_EXTENSION.md. You can use the README_TEMPLATE.md under root directory.
2. Add a new config to the [config file](samples-config-v1.json) with following format.
    1. Fill in these values.
    2. Make sure **id** has identical value with the folder name you just created
    ![alt text](assets/config_format.png)

### Config definition
![alt text](assets/config_definition.png)
