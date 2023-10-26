Attribute VB_Name = "Module1"
Sub stock_v2()
    ' Module 2 Challenge
    ' Loop through each worksheet
    For Each ws In Worksheets

        ' Variables needed
        Dim lr As Long
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim CurrentPrice As Double
        Dim OpenDate As Double
        Dim CloseDate As Double
        Dim CurrentDate As Double
        Dim Tracker As String
        Dim vol As Double
        Dim Summary_Table_Row As Integer
        Dim Maxvol As Double
        Dim MaxYc As Double
        Dim MinYc As Double
        Dim MaxTracker As String
        Dim MinTracker As String
        Dim MaxTrackerVol As String

        ' Determine the Last Row and set other values
        lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        vol = 0
        OpenDate = Empty

        ' Make headers for the results table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' Used for Testing
        ' ws.Range("M1").Value = "OpenPrice"
        ' ws.Range("N1").Value = "ClosePrice"
        ' ws.Range("O1").Value = "OpenDate"
        ' ws.Range("P1").Value = "CloseDate"
        ' ws.Range("Q1").Value = "CurrentDate"

        For i = 2 To lr

            CurrentDate = ws.Cells(i, 2).Value

            ' I made it so the dates don't have to be in order. This will find the min and max close date and set the price.
            If (CurrentDate < OpenDate Or OpenDate = Empty) Then
                OpenPrice = ws.Cells(i, 3).Value
                OpenDate = CurrentDate
                CloseDate = CurrentDate
            ElseIf CurrentDate >= CloseDate Then
                ClosePrice = ws.Cells(i, 6).Value
                CloseDate = CurrentDate
            End If

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                Tracker = ws.Cells(i, 1).Value

                ' Print the Tracker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Tracker

                ' Print the Brand Amount to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = vol

                ' Used for testing
                ' ws.Range("M" & Summary_Table_Row).Value = OpenPrice
                ' ws.Range("N" & Summary_Table_Row).Value = ClosePrice
                ' ws.Range("V" & Summary_Table_Row).Value = OpenDate
                ' ws.Range("W" & Summary_Table_Row).Value = CloseDate
                ' ws.Range("Q" & Summary_Table_Row).Value = CurrentDate

                ws.Range("J" & Summary_Table_Row).Value = ClosePrice - OpenPrice
                ws.Range("K" & Summary_Table_Row).Value = ((ClosePrice - OpenPrice) / OpenPrice)
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%" 'Conditional formatting is applied correctly and appropriately to the percent change column

                ' Make sure to use conditional formatting that will highlight positive change in green and negative change in red. Zero will remain white
                If ClosePrice > OpenPrice Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = VBA.ColorConstants.vbGreen
                ElseIf OpenPrice > ClosePrice Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = VBA.ColorConstants.vbRed
                End If

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                vol = 0
                OpenPrice = 0
                OpenDate = 0

            Else
                vol = vol + ws.Cells(i, 7).Value
            End If

        Next i

        lr = ws.Cells(Rows.Count, 10).End(xlUp).Row

        ' Getting the max/min values using the worksheet functions
        MaxYc = Application.WorksheetFunction.Max(ws.Range("K2:K" & lr))
        MinYc = Application.WorksheetFunction.Min(ws.Range("K2:K" & lr))
        Maxvol = Application.WorksheetFunction.Max(ws.Range("L2:L" & lr))

        ' Getting the tracker using the worksheet match index
        maxValRowIndex = Application.WorksheetFunction.Match(Maxvol, ws.Range("L2:L" & lr), 0) ' Max Volume
        MaxTrackerVol = Application.WorksheetFunction.Index(ws.Range("I2:I" & lr), maxValRowIndex)

        maxValRowIndex = Application.WorksheetFunction.Match(MaxYc, ws.Range("K2:K" & lr), 0) ' Max Yearly Change
        MaxTracker = Application.WorksheetFunction.Index(ws.Range("I2:I" & lr), maxValRowIndex)

        maxValRowIndex = Application.WorksheetFunction.Match(MinYc, ws.Range("K2:K" & lr), 0) ' Min Yearly Change
        MinTracker = Application.WorksheetFunction.Index(ws.Range("I2:I" & lr), maxValRowIndex)

        ws.Range("Q2").Value = MaxYc
        ws.Range("Q2").NumberFormat = "0.00%" 'Conditional formatting is applied correctly and appropriately to the percent change column
        ws.Range("Q3").Value = MinYc
        ws.Range("Q3").NumberFormat = "0.00%" 'Conditional formatting is applied correctly and appropriately to the percent change column
        ws.Range("Q4").Value = Maxvol

        ws.Range("P4").Value = MaxTrackerVol
        ws.Range("P2").Value = MaxTracker
        ws.Range("P3").Value = MinTracker

    Next ws
End Sub

