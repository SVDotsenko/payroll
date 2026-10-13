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

1. Open your workbook in the desktop version of Excel.
2. Press `Alt+F11` to open the Visual Basic Editor.
3. Select **File > Import File...** and import `PRSI.bas` and `USC.bas`.
4. Close the editor and save the workbook as an Excel Macro-Enabled Workbook (`.xlsm`). Enable macros when prompted.
5. Return to the main Excel window (not the Visual Basic Editor) and press `Alt+F8` to open the Macro dialog. Run `RegisterPRSIDescriptions` and `RegisterUSCDescription` to register the function and argument descriptions for Excel's Insert Function wizard. In the wizard, look for the functions under the `payroll` category.
6. Use the functions in worksheet formulas, for example:
   - `=PRSI_Subclass(A2,52)` returns the PRSI subclass for the pay in cell `A2`.
   - `=PRSI_Charge(A2,52)` calculates the Employee PRSI charge.
   - `=PRSI_ChargeSteps(A2,52)` returns the calculation steps as text.
   - `=USC_Charge(A2,52)` calculates USC for the pay in cell `A2`.

Replace `52` with `26` for fortnightly pay or `12` for monthly pay. These functions require desktop Excel; VBA macros are not supported in Excel for the web.
