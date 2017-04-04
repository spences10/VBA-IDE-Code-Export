# Add-in Installer Build Instructions

The following instructions can be followed to create an installer exe from
the add-in workbook `VBA-IDE-Code-Export.xlsm`:

1. Make sure that you have [Inno Setup] installed.
3. Save the add-in workbook as a `xlam` and `xla` file in the `src` folder.
4. Use Inno Setup to compile the `addin-installer.iss` file. You can quickly
   do this by right clicking on the file and clicking the `compile` menu option.
5. The final product (the installation exe) can be found in the `deploy` folder.

[Inno Setup]: http://www.jrsoftware.org/isinfo.php
