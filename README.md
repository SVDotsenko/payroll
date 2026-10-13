This project contains Excel automation files for calculating taxes and other payroll charges for employees in Ireland.

## PRSI Functions

`PRSI.bas` provides the following VBA functions for Class A PRSI calculations:

- `PRSI_Subclass(GrossPay, PayFrequency)` returns the applicable PRSI subclass (`M`, `A0`, `AX`, `AL`, or `A1`).
- `PRSI_Charge(GrossPay, PayFrequency)` calculates the Employee PRSI charge, including the applicable PRSI credit.
- `PRSI_ChargeSteps(GrossPay, PayFrequency)` returns a step-by-step explanation of the Employee PRSI calculation as text.

`PayFrequency` is optional: use `52` for weekly, `26` for fortnightly, or `12` for monthly pay periods. It defaults to weekly.
