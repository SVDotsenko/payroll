Attribute VB_Name = "PRSI"
Sub RegisterPRSIDescriptions()
    Application.MacroOptions _
        Macro:="PRSI_Subclass", _
        Category:="payroll", _
        Description:="Returns the Class A PRSI subclass (M, A0, AX, AL, A1). Rates and thresholds as of 30 September 2026.", _
        ArgumentDescriptions:=Array( _
            "Reckonable pay for the pay period. Zero returns Class M.", _
            "Optional. Pay periods per year: 52 = weekly (default), 26 = fortnightly, 12 = monthly. Any other value is treated as weekly.")

    Application.MacroOptions _
        Macro:="PRSI_Charge", _
        Category:="payroll", _
        Description:="Returns the Employee PRSI charge for Class A. The subclass is determined automatically from the pay. Rates as of 30 September 2026.", _
        ArgumentDescriptions:=Array( _
            "Reckonable pay for the pay period.", _
            "Optional. Pay periods per year: 52 = weekly (default), 26 = fortnightly, 12 = monthly. Any other value is treated as weekly.")

    Application.MacroOptions _
        Macro:="PRSI_ChargeSteps", _
        Category:="payroll", _
        Description:="Returns a step-by-step explanation of the Employee PRSI charge for Class A as text, one step per line. The last step equals the PRSI_Charge result. Rates as of 30 September 2026.", _
        ArgumentDescriptions:=Array( _
            "Reckonable pay for the pay period.", _
            "Optional. Pay periods per year: 52 = weekly (default), 26 = fortnightly, 12 = monthly. Any other value is treated as weekly.")
End Sub

Function PRSI_Subclass(GrossPay As Double, Optional PayFrequency = 52) As String
Attribute PRSI_Subclass.VB_Description = "Returns the Class A PRSI subclass (M, A0, AX, AL, A1). Rates and thresholds as of 30 September 2026."
Attribute PRSI_Subclass.VB_ProcData.VB_Invoke_Func = " \n20"
    If GrossPay = 0 Then
        PRSI_Subclass = "M"'No pay means nil contribution liability (Class M)
        Exit Function
    End If

    Select Case PayFrequency'Column index: 0 = weekly, 1 = fortnightly, 2 = monthly
        Case 26: col = 1
        Case 12: col = 2
        Case Else: col = 0
    End Select

    If GrossPay < Array(38, 76, 165)(col) Then Exit Function'Below the Class A threshold (this would be Class J)

    Set maxPay = CreateObject("Scripting.Dictionary")'Upper limit of each subclass by frequency: weekly, fortnightly, monthly
    maxPay("A0") = Array(352, 704, 1525)
    maxPay("AX") = Array(424, 848, 1837)
    maxPay("AL") = Array(552, 1104, 2392)

    For Each code In maxPay.Keys'Dictionary keeps insertion order
        If GrossPay <= maxPay(code)(col) Then
            PRSI_Subclass = code
            Exit Function
        End If
    Next
    PRSI_Subclass = "A1"
End Function

Function PRSI_Charge(GrossPay As Double, Optional PayFrequency = 52)
Attribute PRSI_Charge.VB_Description = "Returns the Employee PRSI charge for Class A. The subclass is determined automatically from the pay. Rates as of 30 September 2026."
Attribute PRSI_Charge.VB_ProcData.VB_Invoke_Func = " \n20"
    Const EE_RATE = 0.042 ' Employee PRSI rate until 30 Sep 2026

    Dim prsi As Double
    Select Case PRSI_Subclass(GrossPay, PayFrequency)
        Case "AX", "AL", "A1" ' Full rate minus the PRSI Credit (the credit is 0 for AL and A1)
            prsi = GrossPay * EE_RATE - PRSI_Credit(GrossPay, PayFrequency)
    End Select

    PRSI_Charge = IIf(prsi < 0, 0, Round(prsi, 2))
End Function

Private Function PRSI_Credit(GrossPay As Double, PayFrequency) As Double
    If PRSI_Subclass(GrossPay, PayFrequency) <> "AX" Then Exit Function

    weeksPerPeriod = 1
    Select Case PayFrequency
        Case 12
            weeksPerPeriod = 52 / 12
        Case 26
            weeksPerPeriod = 2
    End Select

    Const MAX_CREDIT_WEEKLY = 12
    Const AX_LOWER_WEEKLY = 352.01

    weeklyPay = GrossPay / weeksPerPeriod
    weeklyCredit = MAX_CREDIT_WEEKLY - (weeklyPay - AX_LOWER_WEEKLY) / 6

    If weeklyCredit < 0 Then weeklyCredit = 0
    If weeklyCredit > MAX_CREDIT_WEEKLY Then weeklyCredit = MAX_CREDIT_WEEKLY

    PRSI_Credit =  weeksPerPeriod * Application.WorksheetFunction.Round(weeklyCredit, 2)
End Function
Function PRSI_ChargeSteps(GrossPay As Double, Optional PayFrequency = 52) As String
Attribute PRSI_ChargeSteps.VB_Description = "Returns a step-by-step explanation of the Employee PRSI charge for Class A as text, one step per line. The last step equals the PRSI_Charge result. Rates as of 30 September 2026."
Attribute PRSI_ChargeSteps.VB_ProcData.VB_Invoke_Func = " \n20"
    ' Returns a step-by-step explanation of the PRSI_Charge calculation as text,
    ' one step per line. The last step is the same number PRSI_Charge returns.
    ' Turn on "Wrap Text" in the cell to see all lines.
    Const EE_RATE = 0.042              ' Employee PRSI rate until 30 Sep 2026
    Const MAX_CREDIT_WEEKLY = 12       ' Maximum weekly PRSI Credit
    Const AX_LOWER_WEEKLY = 352.01     ' Weekly earnings where the credit starts to taper

    code = PRSI_Subclass(GrossPay, PayFrequency)
    shownCode = IIf(code = "", "(empty)", code)
    Dim prsi As Double
    n = 0
    Text = ""

    Select Case code
        Case "AX"
            ' Employee PRSI at the full rate
            fullPrsi = GrossPay * EE_RATE
            n = n + 1
            Text = Text & "Step " & n & ". Subclass AX: PRSI at the full rate = " & Format(GrossPay, "0.00") & " x " & Format(EE_RATE * 100, "0.0#") & "% = " & Format(fullPrsi, "0.0000") & "." & vbLf

            ' Number of weeks in the pay period
            weeksPerPeriod = 1
            Select Case PayFrequency
                Case 12     ' Monthly
                    weeksPerPeriod = 52 / 12
                Case 26     ' Fortnightly
                    weeksPerPeriod = 2
            End Select

            ' Weekly equivalent of the pay
            weeklyPay = GrossPay / weeksPerPeriod

            ' Weekly credit: the maximum minus one sixth of earnings above the lower limit
            excess = weeklyPay - AX_LOWER_WEEKLY
            n = n + 1
            Text = Text & "Step " & n & ". Weekly earnings above the lower limit = " & Format(weeklyPay, "0.00") & " - " & AX_LOWER_WEEKLY & " = " & Format(excess, "0.00") & "." & vbLf

            oneSixth = excess / 6
            n = n + 1
            Text = Text & "Step " & n & ". One sixth of that amount = " & Format(excess, "0.00") & " / 6 = " & Format(oneSixth, "0.0000") & "." & vbLf

            rawCredit = MAX_CREDIT_WEEKLY - oneSixth
            n = n + 1
            Text = Text & "Step " & n & ". Weekly credit before limits = " & MAX_CREDIT_WEEKLY & " - " & Format(oneSixth, "0.0000") & " = " & Format(rawCredit, "0.0000") & "." & vbLf

            ' Keep the credit between 0 and the maximum
            limitedCredit = rawCredit
            If rawCredit < 0 Then limitedCredit = 0
            If rawCredit > MAX_CREDIT_WEEKLY Then limitedCredit = MAX_CREDIT_WEEKLY

            ' Round the weekly credit to cents (as in the guide)
            weeklyCredit = Application.WorksheetFunction.Round(limitedCredit, 2)
            n = n + 1
            Text = Text & "Step " & n & ". Weekly credit rounded to cents = " & Format(weeklyCredit, "0.00") & "." & vbLf

            ' Credit for the whole pay period
            periodCredit = weeklyCredit * weeksPerPeriod

            ' Employee PRSI before the final checks
            prsi = fullPrsi - periodCredit
            n = n + 1
            Text = Text & "Step " & n & ". PRSI after credit = " & Format(fullPrsi, "0.0000") & " - " & Format(periodCredit, "0.00") & " = " & Format(prsi, "0.0000") & "." & vbLf

        Case "AL", "A1"
            ' Flat rate on all reckonable pay, no credit
            prsi = GrossPay * EE_RATE
            n = n + 1
            Text = Text & "Step " & n & ". Subclass " & code & ": PRSI at the full rate = " & Format(GrossPay, "0.00") & " x " & Format(EE_RATE * 100, "0.0#") & "% = " & Format(prsi, "0.0000") & "." & vbLf
            n = n + 1
            Text = Text & "Step " & n & ". No PRSI Credit applies to subclass " & code & "." & vbLf

        Case "", "M", "A0"
            n = n + 1
            Text = Text & "Step " & n & ". Subclass " & shownCode & ": no Employee PRSI is charged, so PRSI = 0.00." & vbLf
    End Select

    ' PRSI can never be negative, then round to cents
    If prsi < 0 Then
        result = 0
        n = n + 1
        Text = Text & "Step " & n & ". PRSI cannot be negative: " & Format(prsi, "0.0000") & " is below 0, so the Employee PRSI charge = 0.00."
    Else
        result = Round(prsi, 2)
        n = n + 1
        Text = Text & "Step " & n & ". Rounded to cents: Employee PRSI charge = " & Format(result, "0.00") & "."
    End If

    PRSI_ChargeSteps = Text
End Function
