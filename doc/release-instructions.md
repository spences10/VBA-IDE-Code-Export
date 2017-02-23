# Release Instructions

Steps to take to create a new release:

1. Build the `VBA-IDE-Code-Export.xlsm` workbook.
2. Save the workbook as an add-in `VBA-IDE-Code-Export.xlam`.
3. Create a new GitHub release:
    * Choose a version number following the
      [semantic versioning](http://semver.org) guidelines.
    * Create a new tag following a format similar to `v1.2.3` (substituting the
      appropriate version number).
    * Title the release following a format similar to `Version 1.2.3`
      (substituting the appropriate version number).
    * Attach both the `VBA-IDE-Code-Export.xlsm` workbook and
      `VBA-IDE-Code-Export.xlam` add-in as downloadable binaries.
