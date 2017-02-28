# VBA IDE CodeExport

[![The MIT License](https://img.shields.io/badge/license-MIT-orange.svg?style=flat-square)](http://opensource.org/licenses/MIT)
[![Gitter](https://img.shields.io/gitter/room/nwjs/nw.js.svg?style=flat-square)](https://gitter.im/VBA-IDE-Code-Export)
[![Stories in Ready](https://badge.waffle.io/spences10/VBA-IDE-Code-Export.png?label=ready&title=Ready)](https://waffle.io/spences10/VBA-IDE-Code-Export)

[![Throughput Graph](https://graphs.waffle.io/spences10/VBA-IDE-Code-Export/throughput.svg)](https://waffle.io/spences10/VBA-IDE-Code-Export/metrics/throughput)

<!-- TOC -->

- [VBA IDE CodeExport](#vba-ide-codeexport)
  - [Intro](#intro)
  - [Installing](#installing)
  - [Usage](#usage)
    - [The configuration file](#the-configuration-file)
    - [Importing](#importing)
    - [Exporting](#exporting)
  - [Build from source](#build-from-source)
    - [Clone the repo](#clone-the-repo)
    - [Build the binary](#build-the-binary)
    - [Add the code](#add-the-code)
  - [Contributing](#contributing)
  - [Roadmap](#roadmap)

<!-- /TOC -->

## Intro

Export your excel VBA project source code for use with Git (or any VCS) from the Excel Developer ribbon.

For a while now I have used this code so that all the associated VBA files used in a VBA project (*.cls, *.bas, *.frm files) can be effortlessly exported for use with a Version Control System.

This is specifically for Excel, although the VBIDE extensibility can be used for all the MS Office suite.

## Installing

1. [Download](https://github.com/spences10/VBA-IDE-Code-Export/releases) the Add-In or Add-In Workbook.
2. Put the add-in into your add-ins folder. The easiest way to do this is to open the Add-In workbook (.xlsm file) and save it as an Add-In. Excel will automatically direct you to the correct folder.
3. Enable the add-in in Excel.
4. Check the `Trust access to the VBA project model` check box located in `Trust Centre -> Trust Centre Settings -> Macro Settings -> Trust access to the VBA project model`.

## Usage

On installing the add-in you will have the buttons for CodeExport in the Excel Developer ribbon. The add-in will also create a menu in the VBA IDE (the VBE) called `Export for VCS`. All controls for the add-in are found in this menu.

![](/img/ribbon-buttons.png)

### The configuration file

A file named `CodeExport.config.json` in the same directory as an Excel file declares what gets imported into that Excel file. The `Make Config File` button in the `Export For VCS` menu will create or update the configuration file for the current active project based upon the contents of that project. The JSON file format is used as the file format for the configuration file.

The `Module Paths` property specifies a mapping of VBA modules to their location in the file system. File paths may be either relative or absolute. Relatives paths are relative to the directory of the configuration file and the Excel file. The `Base Path` property can be used to add a common prefix to all the file paths.

The `References` property declares the references to libraries that your VBA modules require. These will be imported when the import action is used and will be removed when the export action is used.

The `VBAProject Name` property declares the VBAProject name. This will be imported when the import action is used.

### Importing

The `Import` button in the `Export For VCS` menu will:

* Import all the modules specified in the configuration file from the file system into the Excel file. Existing modules will be overwritten.
* Add all library references declared in the configuration file. Existing library references will be overwritten.
* Set the VBAProject name as declared in the configuration file.

### Exporting

The `Export` button in the `Export For VCS` menu will:

* Export all the modules specified in the configuration file from the Excel file into the appropriate places in the file system. Existing files will be overwritten.
* Remove library references from the project which are declared in the configuration file.

## Build from source

### Clone the repo

Clone the repo `git clone https://github.com/spences10/VBA-IDE-Code-Export`, navigate to where you have cloned the code to, there you will find the `VBA-IDE-Code-Export.package` folder this is the 'unpacked' version of the VBA-IDE-Code-Export Excel `.xlsm` binary.

### Build the binary

All you need to do to build the binary is create an empty `.zip` file, name it `VBA-IDE-Code-Export` then open the empty `.zip` and drag and drop the **_contents_** of the `VBA-IDE-Code-Export.package` folder into the `.zip` file.

>Now [I'm assuming you're on windows] if you have the Folder Options setting for `Hide extensions for known file types` checked then this is the time to uncheck it.

![](/img/unhide-file-extensions.gif)

Rename the file extension on the newly created `VBA-IDE-Code-Export.zip` file from `.zip` to `.xlsm` acknowledge the dialog saying `if you change the extension bad things might happen` and you will have the binary ready to add the code to.

![](/img/build-from-source.gif)

### Add the code

Best way to do this, have the `VBA-IDE-Code-Export.xlsm` VBA IDE open in one window then the file explorer open at the `src` folder in another window.

Multi select the contents of the folder **_excluding `VBA-IDE-Code-Export.package` and `CodeExport.config.json`_** drag and drop into the `VBA-IDE-Code-Export.xlsm` VBA IDE.

Add in the following references: 

1. Microsoft Scripting Runtime
2. Microsoft Visual Basic for Applications Extensibility 5.3
3. Windows Script Host Object Model
4. Microsoft Shell Controls And Automation

Save, Debug>Compile the project then from the Immediate pane in the VBA IDE enter `auto_open` and hit return this should create the VBA IDE menu items, you're ready to move onto **usage.**

![](/img/add-code.gif)


## Contributing

Please fork this repository and contribute back using pull requests.

Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.

Please use the template file `VBA-IDE-Code-Export.xlsm` for working in, however don't commit the template file unless you are actually making a change to the template file. This helps with source control since merging an Excel file is not fun.

## Roadmap

- [x] Add pretty ribbon UI
- [ ] Save XL as XML
- [ ] Command line interface