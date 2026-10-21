Attribute VB_Name = "USC"
Private Sub RegisterUSCDescription()
    Application.MacroOptions _
        Macro:="USC_Charge", _
        Category:="payroll", _
        Description:="Returns the Universal Social Charge (USC) for the pay period at standard rates (0.5%, 2%, 3%, 8%).", _
        ArgumentDescriptions:=Array( _
            "Gross pay for the pay period.", _
            "Optional. Pay periods per year: 52 = weekly (default), 26 = fortnightly, 12 = monthly. Any other value is treated as weekly.")
End Sub

Function USC_Charge(GrossPay As Double, Optional PayFrequency = 52) As Double
Attribute USC_Charge.VB_Description = "Returns the Universal Social Charge (USC) for the pay period at standard rates (0.5%, 2%, 3%, 8%)."
Attribute USC_Charge.VB_ProcData.VB_Invoke_Func = " \n20"
    If GrossPay <= 0 Then Exit Function ' No USC on zero or negative pay (the function returns 0)
    ' Each pair: annual size of the PREVIOUS band, and the rate that starts after it.
    ' So 16688 next to 3% does not mean "3% from 16688": 3% starts at 12012 + 16688 = 28700.
    ' Bands as From - To (annual): 0 - 12012 at 0.5%, 12012 - 28700 at 2%,
    ' 28700 - 70044 at 3%, above 70044 at 8%.
    ' Since a pair only knows where its band starts, each loop step closes the PREVIOUS
    ' band: its end is now known, so it is taxed whole at its rate. The band where pay
    ' ends (including the last one, which has no upper limit) is taxed after the loop.
    ' Example, 300 a week (first band is 12012 / 52 = 231):
    '   0 - 231 at 0.5% = 1.155, then 231 - 300 at 2% = 1.38, total 2.535
    bands = Array( _
        Array(0, 0.005), _
        Array(12012, 0.02), _
        Array(16688, 0.03), _
        Array(41344, 0.08))

    periods = 52 ' Weekly by default; any other value is treated as weekly
    If IsNumeric(PayFrequency) Then
        Select Case PayFrequency
            Case 12, 26: periods = PayFrequency ' Monthly or fortnightly
        End Select
    End If

    Dim usc As Double, lower As Double, rate As Double ' Start and rate of the current band (none before the first pair)

    For Each band In bands
        upper = lower + band(0) / periods ' End of the current band = start of the next one
        If GrossPay <= upper Then Exit For ' Pay ends inside the current band
        usc = usc + (upper - lower) * rate ' Current band is full: tax it whole, From - To
        lower = upper
        rate = band(1)
    Next

    usc = usc + (GrossPay - lower) * rate ' Band where pay ends: from its start up to the pay

    USC_Charge = Application.WorksheetFunction.Round(usc, 2)
End Function
