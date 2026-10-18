Attribute VB_Name = "Test_PRSI"
' Tests for the PRSI module. Requires TestHelpers.
' Expected values come from "PRSI 2026 Contribution Rates and User Guide", table "Earnings bands and subclasses".
' Run RunPRSITests (Alt+F8 or F5 in the VBA editor) and check the Immediate window.

Sub RunPRSITests()
    BeginTests
    Test_PRSI_Subclass_Weekly
    Test_PRSI_Subclass_Fortnightly
    Test_PRSI_Subclass_Monthly
    Test_PRSI_Subclass_DefaultFrequency
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

Private Sub CheckSubclass(ByVal grossPay As Double, ByVal payFrequency, ByVal expected As String)
    AssertEqual expected, PRSI_Subclass(grossPay, payFrequency), "PRSI_Subclass(" & grossPay & ", " & payFrequency & ")"
End Sub
