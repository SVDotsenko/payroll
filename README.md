This project contains Excel automation files for calculating taxes and other payroll charges for employees in Ireland.

## PRSI Functions

`PRSI.bas` provides the following VBA functions for Class A PRSI calculations:

- `PRSI_Subclass(GrossPay, PayFrequency)` returns the applicable PRSI subclass (`M`, `A0`, `AX`, `AL`, or `A1`).
- `PRSI_Charge(GrossPay, PayFrequency)` calculates the Employee PRSI charge, including the applicable PRSI credit.
- `PRSI_ChargeSteps(GrossPay, PayFrequency)` returns a step-by-step explanation of the Employee PRSI calculation as text.

`PayFrequency` is optional: use `52` for weekly, `26` for fortnightly, or `12` for monthly pay periods. It defaults to weekly.

## USC Function

`USC.bas` provides `USC_Charge(GrossPay, PayFrequency)`, which calculates Universal Social Charge (USC) for the pay period using the standard rates: 0.5%, 2%, 3%, and 8%. The annual band amounts are €12,012, €16,688, and €41,344; any pay above these bands is charged at 8%. `PayFrequency` accepts `52` for weekly, `26` for fortnightly, or `12` for monthly pay periods. It is optional and defaults to weekly; any other value is treated as weekly.

## Using the VBA Module in Excel

You only need to set this up once per workbook. After that, the workbook downloads the latest payroll modules from GitHub every time it opens. Choose one of the three options below.

### Option 1: Set up the workbook manually

1. Download [`loader.bas`](https://github.com/SVDotsenko/payroll/blob/main/loader.bas): open the link and click **Download raw file**. Make sure the saved file is named `loader.bas`, not `loader.bas.txt`.
2. In Excel, go to **File > Options > Trust Center > Trust Center Settings... > Macro Settings** and tick **Trust access to the VBA project object model**. The loader needs this to update the modules.
3. Open your workbook in the desktop version of Excel and press `Alt+F11` to open the Visual Basic Editor.
4. Select **File > Import File...** and import the downloaded `loader.bas`.
5. In the **Project** pane, double-click **ThisWorkbook** and paste this code:

   ```vba
   Private Sub Workbook_Open()
       UpdatePayrollModules
   End Sub
   ```

6. Close the editor and save the workbook as an Excel Macro-Enabled Workbook (`.xlsm`).
7. Close the workbook, then open it again and enable macros when prompted. The `PRSI` and `USC` modules are downloaded and the function descriptions are registered automatically. In Excel's Insert Function wizard, the functions appear under the `payroll` category.

### Option 2: Create the workbook automatically with `setup.ps1`

This option creates a new, ready-to-use workbook on your Desktop and turns on the required Excel setting for you.

1. Download [`setup.ps1`](https://github.com/SVDotsenko/payroll/blob/main/setup.ps1): open the link and click **Download raw file**. Make sure the saved file is named `setup.ps1`, not `setup.ps1.txt`.
2. Close Excel.
3. Right-click the downloaded `setup.ps1` and select **Run with PowerShell**. If Windows shows a security warning about a file downloaded from the internet, clear the **Always ask before opening this file** checkbox and then click **Open**. If the checkbox stays ticked, nothing happens and the workbook is not created.
4. Wait for the message **The payroll workbook has been saved to your Desktop**.
5. Open `Payroll.xlsm` from your Desktop and click **Enable Content** when Excel asks about macros. The `PRSI` and `USC` modules are downloaded and the function descriptions are registered automatically.

If `Payroll.xlsm` already exists on your Desktop, the script stops without changing it. Rename or move the existing file and run the script again.

### Option 3: Run `setup.ps1` directly from GitHub

This option does the same as Option 2, but you don't need to download any file, and Windows does not show the security warning.

1. Close Excel.
2. Open the **Start** menu, type `PowerShell`, and open **Windows PowerShell**.
3. Paste this command and press `Enter`:

   ```powershell
   irm https://raw.githubusercontent.com/SVDotsenko/payroll/main/setup.ps1 | iex
   ```

4. Wait for the message **The payroll workbook has been saved to your Desktop**, then close the PowerShell window.
5. Open `Payroll.xlsm` from your Desktop and click **Enable Content** when Excel asks about macros. The `PRSI` and `USC` modules are downloaded and the function descriptions are registered automatically.

As with Option 2, if `Payroll.xlsm` already exists on your Desktop, the script stops without changing it.

### Using the functions

When the workbook opens with an internet connection, it automatically updates the payroll functions to the latest version. Without an internet connection you can still use the workbook as usual, but the functions may not be the latest version.

The very first time a new workbook is opened, an internet connection is required, because that is when the functions are downloaded for the first time. If there is no connection, Excel shows a message; connect to the internet and open the workbook again.

Use the functions in worksheet formulas, for example:
- `=PRSI_Subclass(A2,52)` returns the PRSI subclass for the pay in cell `A2`.
- `=PRSI_Charge(A2,52)` calculates the Employee PRSI charge.
- `=PRSI_ChargeSteps(A2,52)` returns the calculation steps as text.
- `=USC_Charge(A2,52)` calculates USC for the pay in cell `A2`.

Replace `52` with `26` for fortnightly pay or `12` for monthly pay. These functions require desktop Excel; VBA macros are not supported in Excel for the web.
