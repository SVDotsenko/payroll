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

You only need to do this once. It creates a ready-to-use payroll workbook on your Desktop.

1. Close Excel.
2. Press `Win+R`, paste the following line and press **Enter**:

   ```
   powershell -ExecutionPolicy Bypass -c "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/SVDotsenko/payroll/main/setup.ps1)))"
   ```

3. Wait for the message **The payroll workbook has been saved to your Desktop**.
4. Open `Payroll.xlsm` from your Desktop and click **Enable Content** when Excel asks about macros.

Every time the workbook opens, it downloads the latest payroll functions automatically, so it needs an internet connection. In Excel's Insert Function wizard, the functions appear under the `payroll` category.

Use the functions in worksheet formulas, for example:
- `=PRSI_Subclass(A2,52)` returns the PRSI subclass for the pay in cell `A2`.
- `=PRSI_Charge(A2,52)` calculates the Employee PRSI charge.
- `=PRSI_ChargeSteps(A2,52)` returns the calculation steps as text.
- `=USC_Charge(A2,52)` calculates USC for the pay in cell `A2`.

Replace `52` with `26` for fortnightly pay or `12` for monthly pay. These functions require desktop Excel; VBA macros are not supported in Excel for the web.
