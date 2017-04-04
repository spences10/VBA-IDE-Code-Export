# Release Instructions

Steps to take to create a new release:

1. Choose a version number following the
   [semantic versioning](http://semver.org) guidelines.
2. Update version number in `config.iss` and git commit the change.
2. Build the `VBA-IDE-Code-Export.xlsm` workbook. Don't forget to set a
   VBA Project password of "123".
3. [Build the installer](installer-build-instructions.md) (`CodeExport_setup_1.2.3.exe`)
4. Create a new GitHub release:
    * Create a new tag following a format similar to `v1.2.3` (substituting the
      appropriate version number).
    * Title the release following a format similar to `Version 1.2.3`
      (substituting the appropriate version number).
    * Attach the following as downloadable binaries.
        * `CodeExport_setup_1.2.3.exe`
        * `VBA-IDE-Code-Export.xlam`
5. Cheer and celebrate loudly about another great release. Let it be known!
