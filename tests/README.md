# Tests

Unit tests for the payroll VBA modules. They use a small assertion module written in plain VBA, so no add-ins or extra tools are needed.

The tests are for development only. The loader never downloads this folder from GitHub, so users' workbooks do not get the tests.

## Files

- `TestHelpers.bas` – the assertion helpers that every test module uses.
- `Test_PRSI.bas` – tests for `PRSI_Subclass` and `PRSI_Charge`. Expected values come from the Department of Social Protection guide _Pay Related Social Insurance (PRSI) 2026 Contribution Rates and User Guide_.

## How it works

A test module has one public procedure that runs all of its tests, for example `RunPRSITests`:

```vba
Sub RunPRSITests()
    BeginTests
    Test_PRSI_Subclass_Weekly
    ' ...more test procedures
    EndTests
End Sub
```

`BeginTests` resets the counters, and each test procedure calls assertions from `TestHelpers`:

| Assertion                                                               | Use it for                                                                                                     |
| ----------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| `AssertEqual expected, actual, testName`                                | Strings and whole numbers. Numbers are compared by value, whatever their type. A number never equals a string. |
| `AssertNear expected, actual, testName[, tolerance]`                    | Money and other `Double` results. The default tolerance is half a cent (0.005).                                |
| `AssertTrue condition, testName`                                        | Any other check.                                                                                               |
| `AssertExpectedFailure expected, actual, testName, reason[, tolerance]` | A known discrepancy, for example a value printed in the guide that the code deliberately does not reproduce.   |

`AssertExpectedFailure` works like `xfail` in pytest:

- If the values differ, as expected, the check is logged as `XFAIL` and does not fail the run.
- If the values match, the check is reported as a failure (`XPASS`). This means the known problem has gone, and the check should become a normal assertion.

`EndTests` prints a summary such as `Passed: 53, failed: 0, expected failures: 3` to the Immediate window. It then does one of two things:

- If all tests passed, it says "All tests passed" aloud and shows a popup asking whether to open the Immediate window. **No** is the default button, and the popup closes by itself after 3 seconds.
- If any test failed, it says so aloud and opens the VBA editor with the Immediate window, which lists every `FAIL:` line.

## Running the tests

### On your computer, with the local repository (automatic)

When the repository is cloned to `C:\repositories\payroll`, the loader imports the modules from that folder instead of GitHub. In that case it also imports every `.bas` file from `tests`. Each time you open the workbook, the latest test modules are imported along with the payroll modules.

`loader.bas` does not update itself. If your workbook has an older loader, import the current `loader.bas` once by hand, as described in step 4 of [Option 1 in the main README](../README.md#option-1-set-up-the-workbook-manually).

### In any workbook (manual)

1. Download [`TestHelpers.bas`](https://github.com/SVDotsenko/payroll/blob/main/tests/TestHelpers.bas) and the test modules you need, for example [`Test_PRSI.bas`](https://github.com/SVDotsenko/payroll/blob/main/tests/Test_PRSI.bas). Open each link and click **Download raw file**. Make sure the saved files keep the `.bas` extension, for example `TestHelpers.bas`, not `TestHelpers.bas.txt`.
2. Open the workbook in desktop Excel and press `Alt+F11` to open the Visual Basic Editor. The workbook must already contain the modules under test, for example `PRSI`.
3. Select **File > Import File...** and import `TestHelpers.bas`.
4. Repeat for the test modules you downloaded, for example `Test_PRSI.bas`.
5. Run the test procedure, for example `RunPRSITests`. Either put the cursor inside it and press `F5`, or press `Alt+F8` in Excel and choose it from the list.
6. Read the result in the popup or in the Immediate window (`Ctrl+G` in the editor).

To update a test module after changing it, remove the old module first: right-click it in the **Project** pane and choose **Remove**, answering **No** when asked to export it. Then import the file again. Otherwise the VBA editor imports it as a second copy named, for example, `Test_PRSI1`.

To open the Immediate window automatically, `EndTests` needs **Trust access to the VBA project object model**. This is under **File > Options > Trust Center > Trust Center Settings... > Macro Settings**. Without it, the tests still run and print their results, but the window does not open by itself.
