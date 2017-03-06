# CodeExport Comprehensive Test Project

This is an example project which uses CodeExport and uses as many features as
possible. It serves as both an example of how to use CodeExport and as a
test subject for integration tests.

## Quick start guide

To 'build' the workbook use the following procedure:

1. Open the template workbook `template-workbook.xlsm` in Excel.
2. Run CodeExport `Import` from the developer ribbon menu or the
   'Export for VCS' in the VBE menu.

That's it! The workbook should now be ready for use.

## Test procedures

Below are some examples of test procedures you may follow. These are guides
and don't cover everything. Use your intuition and imagination to try and make
CodeExport do something that it shouldn't do or something that you think is
unhelpful. If you find any problems, make sure to record your hard earned
discovery in the [issues list].

### Quick test

1. Open template workbook `template-workbook.xlsm` in Excel.
2. Run CodeExport `Import`. No errors should be shown.
3. Click `Test VBA code` button in `Sheet1` of the template workbook.
   The test success dialog should be shown.
4. Run CodeExport `Make config file`. No errors should be shown. The config file
   should be unchanged.
5. Run CodeExport `Export`. No errors should be shown. The files in the file
   file system should be unchanged. The workbook should return to it's
   "template" state where there is no VBA code, references, etc.

### Comprehensive test procedure

1. Open template workbook `template-workbook.xlsm` in Excel.

2. Run CodeExport `Import`.
3. Click the `Test VBA code` button in `Sheet1` of the template workbook.
   The test success dialog should be shown.
4. Run CodeExport `Import` a second time. Nothing should change and no errors
   should be raised.
5. Click the `Test VBA code` button again. Everything should run the same as
   before.

6. Run CodeExport `Make config file`. Confirm that the configuration file is
   exactly the same (Hint: use `git diff` or `git status`).

8. Run CodeExport `Export`.
9. Confirm that the "source file" in the file system are unchanged.
10. Click the `Test VBA code` button. This should fail. The workbook should have
    returned to the "template" state where there is no VBA code, references,
    etc.
11. Run CodeExport `Export` again. No errors should be shown. Source files in
    the file system should be unchanged.

[issues list]: https://github.com/spences10/VBA-IDE-Code-Export/issues
