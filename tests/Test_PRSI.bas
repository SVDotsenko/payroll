Attribute VB_Name = "Test_PRSI"
' Tests for the PRSI module. Requires TestHelpers.
' Expected values come from "PRSI 2026 Contribution Rates and User Guide":
' subclasses from the table "Earnings bands and subclasses", AX charges from the table on page 3.
' Run RunPRSITests (Ctrl+Shift+Q, Alt+F8 or F5 in the VBA editor) and check the Immediate window.

Sub RunPRSITests()
Attribute RunPRSITests.VB_Description = "Runs the PRSI unit tests; results are printed to the Immediate window (Ctrl+G)."
Attribute RunPRSITests.VB_ProcData.VB_Invoke_Func = "Q\n14"
    BeginTests
    Test_PRSI_Subclass_Weekly
    Test_PRSI_Subclass_Fortnightly
    Test_PRSI_Subclass_Monthly
    Test_PRSI_Subclass_DefaultFrequency
    Test_PRSI_Charge_AX_Table
    Test_PRSI_Charge_GuideDiscrepancies
    EndTests
End Sub

' Weekly bands: A0 38-352, AX 352.01-424, AL 424.01-552, A1 more than 552.
Private Sub Test_PRSI_Subclass_Weekly()
    CheckSubclass 38, 52, "A0"         ' A0 lower bound
    CheckSubclass 200, 52, "A0"
    CheckSubclass 352, 52, "A0"        ' A0 upper bound
    CheckSubclass 352.01, 52, "AX"     ' AX lower bound
    CheckSubclass 377, 52, "AX"        ' Worked example from the guide
    CheckSubclass 424, 52, "AX"        ' AX upper bound
    CheckSubclass 424.01, 52, "AL"     ' AL lower bound
    CheckSubclass 500, 52, "AL"
    CheckSubclass 552, 52, "AL"        ' AL upper bound
    CheckSubclass 552.01, 52, "A1"     ' A1 lower bound
    CheckSubclass 100000, 52, "A1"
End Sub

' Fortnightly bands: A0 76-704, AX 704.01-848, AL 848.01-1104, A1 more than 1104.
Private Sub Test_PRSI_Subclass_Fortnightly()
    CheckSubclass 76, 26, "A0"
    CheckSubclass 400, 26, "A0"
    CheckSubclass 704, 26, "A0"
    CheckSubclass 704.01, 26, "AX"
    CheckSubclass 800, 26, "AX"
    CheckSubclass 848, 26, "AX"
    CheckSubclass 848.01, 26, "AL"
    CheckSubclass 1000, 26, "AL"
    CheckSubclass 1104, 26, "AL"
    CheckSubclass 1104.01, 26, "A1"
    CheckSubclass 100000, 26, "A1"
End Sub

' Monthly bands: A0 165-1525, AX 1525.01-1837, AL 1837.01-2392, A1 more than 2392.
Private Sub Test_PRSI_Subclass_Monthly()
    CheckSubclass 165, 12, "A0"
    CheckSubclass 1000, 12, "A0"
    CheckSubclass 1525, 12, "A0"
    CheckSubclass 1525.01, 12, "AX"
    CheckSubclass 1700, 12, "AX"
    CheckSubclass 1837, 12, "AX"
    CheckSubclass 1837.01, 12, "AL"
    CheckSubclass 2000, 12, "AL"
    CheckSubclass 2392, 12, "AL"
    CheckSubclass 2392.01, 12, "A1"
    CheckSubclass 100000, 12, "A1"
End Sub

' Omitted or unsupported PayFrequency falls back to the weekly bands.
Private Sub Test_PRSI_Subclass_DefaultFrequency()
    AssertEqual "A0", PRSI_Subclass(352), "PRSI_Subclass(352) without frequency"
    AssertEqual "AX", PRSI_Subclass(352.01), "PRSI_Subclass(352.01) without frequency"
    AssertEqual "A1", PRSI_Subclass(552.01), "PRSI_Subclass(552.01) without frequency"
    ' 1000 is A0 monthly but A1 weekly, so this shows the weekly bands were used
    CheckSubclass 1000, 4, "A1"
End Sub

' Weekly AX charges from the guide's table "Class A PRSI charge for gross earnings between 352.01 and 424" (page 3).
' Charge = Round(GrossPay * 4.2%, 2) - Round(12 - (GrossPay - 352.01) / 6, 2).
Private Sub Test_PRSI_Charge_AX_Table()
    CheckCharge 352.01, 52, 2.78       ' AX lower bound: maximum credit of 12
    CheckCharge 355, 52, 3.41
    CheckCharge 360, 52, 4.45
    CheckCharge 365, 52, 5.49          ' Credit 9.835 rounds up to 9.84
    CheckCharge 370, 52, 6.54
    CheckCharge 375, 52, 7.58
    CheckCharge 380, 52, 8.62          ' Credit 7.335 rounds up to 7.34
    CheckCharge 385, 52, 9.67
    CheckCharge 390, 52, 10.71
    CheckCharge 395, 52, 11.75         ' Table says 11.76, but it uses 16.60 for 395 * 4.2% = 16.59
    CheckCharge 400, 52, 12.8
    CheckCharge 405, 52, 13.84
    CheckCharge 410, 52, 14.88         ' Credit 2.335 rounds up to 2.34
    CheckCharge 415, 52, 15.93
    CheckCharge 420, 52, 16.97
    CheckCharge 424, 52, 17.81         ' AX upper bound: table says 17.80, but 424 * 4.2% = 17.808 rounds to 17.81
End Sub

' Values printed in the guide that the code deliberately does not reproduce (expected failures).
' The code follows the method of the page 3 table: the credit is rounded as a whole, not its one-sixth part.
Private Sub Test_PRSI_Charge_GuideDiscrepancies()
    AssertExpectedFailure 8, PRSI_Charge(377, 52), "PRSI_Charge(377, 52) worked example on page 2", _
        "the example rounds one-sixth first (12 - 4.17 = 7.83); the table method gives a credit of 7.84 and 7.99"
    AssertExpectedFailure 11.76, PRSI_Charge(395, 52), "PRSI_Charge(395, 52) page 3 table", _
        "the table uses 16.60 for 395 * 4.2%, which is 16.59"
    AssertExpectedFailure 17.8, PRSI_Charge(424, 52), "PRSI_Charge(424, 52) page 3 table", _
        "the table gives 17.80 for 424 * 4.2% = 17.808, which rounds to 17.81"
End Sub

Private Sub CheckCharge(ByVal grossPay As Double, ByVal payFrequency, ByVal expected As Double)
    AssertNear expected, PRSI_Charge(grossPay, payFrequency), "PRSI_Charge(" & grossPay & ", " & payFrequency & ")"
End Sub

Private Sub CheckSubclass(ByVal grossPay As Double, ByVal payFrequency, ByVal expected As String)
    AssertEqual expected, PRSI_Subclass(grossPay, payFrequency), "PRSI_Subclass(" & grossPay & ", " & payFrequency & ")"
End Sub
