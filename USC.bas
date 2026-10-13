Attribute VB_Name = "USC"
Sub RegisterUSCDescription()
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
    bands = Array(12012, 16688, 41344) ' Annual band sizes (the balance above them is taxed at the last rate)
    rates = Array(0.005, 0.02, 0.03, 0.08) ' Rates: one per band, plus the last one for the balance

    periods = 52
    If IsNumeric(PayFrequency) Then
        For Each f In Array(12, 26)
            If PayFrequency = f Then periods = f
        Next
    End If
    
    If GrossPay <= 0 Then Exit Function ' No USC on zero or negative pay (the function returns 0)

    Dim usc As Double
    remaining = GrossPay
    
    For i = 0 To UBound(bands) ' Go through the bands: take what fits into the band, tax it, carry the rest on
        bandSize = bands(i) / periods
        amount = IIf(remaining > bandSize, bandSize, remaining)
        usc = usc + amount * rates(i)
        remaining = remaining - amount
    Next
    
    usc = usc + remaining * rates(UBound(rates)) ' Balance: taxed at the last rate (remaining is 0 if pay did not reach this band)

    USC_Charge = Application.WorksheetFunction.Round(usc, 2)
End Function
