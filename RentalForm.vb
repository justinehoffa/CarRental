Option Strict On
Option Explicit On
Option Compare Binary

'Justine Hoffa
'RCET0265
'Fall2020
'Car Rental
'https://github.com/justinehoffa/Car-Rental

Public Class RentalForm
    Dim totalCustomers As Integer
    Dim totalDistance As Decimal
    Dim totalCharge As Decimal

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        SummaryButton.Enabled = False
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim Result As DialogResult
        Result = MessageBox.Show("Are you sure you want to exit?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
        If Result = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Dim summary As String
        summary = "Total Customers:      " & totalCustomers.ToString & vbNewLine &
            "Total Distance:         " & totalDistance.ToString & vbNewLine &
            "Total Charge:            " & totalCharge.ToString("C")
        MessageBox.Show(summary, "Detailed Summary")
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        MilesradioButton.Select()
        SummaryButton.Enabled = False
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click

        Dim ValidateCheckBox As Boolean = False
        Dim userMessage As String = ""
        Dim beginningOdometerReading As Decimal
        Dim beginningOdometer As Boolean = False
        Dim endingOdometerReading As Decimal
        Dim endingOdometer As Boolean = False
        Dim numberofDays As Integer
        Dim numberDays As Boolean = False
        Dim dayCharge As Decimal
        Dim miles As Decimal
        Dim mileageTotal As Decimal
        Dim amountOwed As Decimal

        If NameTextBox.Text = "" Then
            userMessage = "Please enter a Customer Name." & vbNewLine
        End If
        If AddressTextBox.Text = "" Then
            userMessage &= "Please enter an Address." & vbNewLine
        End If
        If CityTextBox.Text = "" Then
            userMessage &= "Please enter a City." & vbNewLine
        End If
        If StateTextBox.Text = "" Then
            userMessage &= "Please enter a State." & vbNewLine
        End If
        If ZipCodeTextBox.Text = "" Then
            userMessage &= "Please enter a Zip Code." & vbNewLine
        End If
        If BeginOdometerTextBox.Text = "" Then
            userMessage &= "Please enter a Beginning Odometer Reading." & vbNewLine
        End If
        If EndOdometerTextBox.Text = "" Then
            userMessage &= "Please enter an Ending Odometer Reading." & vbNewLine
        End If
        If DaysTextBox.Text = "" Then
            userMessage &= "Please enter a Number of Days." & vbNewLine
        End If

        If BeginOdometerTextBox.Text <> "" Or EndOdometerTextBox.Text <> "" Or DaysTextBox.Text <> "" Then
            Try
                beginningOdometerReading = CDec(BeginOdometerTextBox.Text)
                beginningOdometer = True
            Catch ex As Exception
                userMessage &= "Please enter a valid Beginning Odometer Reading." & vbNewLine
            End Try
            Try
                endingOdometerReading = CDec(EndOdometerTextBox.Text)
                endingOdometer = True
            Catch ex As Exception
                userMessage &= "Please enter a valid Ending Odometer Reading." & vbNewLine
            End Try
            Try
                numberofDays = CInt(DaysTextBox.Text)
                If numberofDays > 0 And numberofDays < 46 Then
                    numberDays = True
                Else
                    userMessage &= "Please enter a Number of Days between 0 and 45." & vbNewLine
                End If
            Catch ex As Exception
                userMessage &= "Please enter a valid Number of Days." & vbNewLine
            End Try
            If beginningOdometer = True And endingOdometer = True Then
                If beginningOdometerReading < endingOdometerReading Then
                Else
                    userMessage &= "Beginning Odometer Reading must be less than Ending Odometer Reading." & vbNewLine
                End If
            End If
        End If

        If userMessage <> "" Then
            MsgBox(userMessage)
            If beginningOdometer = False Then
                BeginOdometerTextBox.Text = ""
            End If
            If endingOdometer = False Then
                EndOdometerTextBox.Text = ""
            End If
            If numberDays = False Then
                DaysTextBox.Text = ""
            End If
        Else
            ValidateCheckBox = True
        End If

        If numberDays = False Then
            DaysTextBox.Select()
        End If
        If endingOdometer = False Then
            EndOdometerTextBox.Select()
        End If
        If beginningOdometer = False Then
            BeginOdometerTextBox.Select()
        End If
        If ZipCodeTextBox.Text = "" Then
            ZipCodeTextBox.Select()
        End If
        If StateTextBox.Text = "" Then
            StateTextBox.Select()
        End If
        If CityTextBox.Text = "" Then
            CityTextBox.Select()
        End If
        If AddressTextBox.Text = "" Then
            AddressTextBox.Select()
        End If
        If NameTextBox.Text = "" Then
            NameTextBox.Select()
        End If

        If ValidateCheckBox = True Then
            dayCharge = (numberofDays * 15D)
            DayChargeTextBox.Text = dayCharge.ToString("C")
            If MilesradioButton.Checked = True Then
                miles = endingOdometerReading - beginningOdometerReading
            Else
                miles = (endingOdometerReading - beginningOdometerReading) * 0.6214D
            End If
            TotalMilesTextBox.Text = miles.ToString
            Select Case miles
                Case <= 200
                    mileageTotal = miles * 0
                Case > 500
                    mileageTotal += (miles - 500D) * 0.1D + 36D
                Case Else
                    mileageTotal = (miles - 200D) * 0.12D
            End Select
            MileageChargeTextBox.Text = mileageTotal.ToString("C")

            Dim totalCharges As Decimal
            totalCharges = CDec(DayChargeTextBox.Text) + CDec(MileageChargeTextBox.Text)

            Dim totalDiscount As Decimal
            If AAAcheckbox.Checked = True Then
                totalDiscount = totalCharges * 0.05D
            End If
            If Seniorcheckbox.Checked = True Then
                totalDiscount += totalCharges * 0.03D
            End If
            TotalDiscountTextBox.Text = totalDiscount.ToString("C")
            TotalDiscountTextBox.Text = totalDiscount.ToString("C")

            amountOwed = CDec(totalCharges - totalDiscount)
            TotalChargeTextBox.Text = amountOwed.ToString("C")

            SummaryButton.Enabled = True

            totalCustomers += 1
            totalDistance += CDec(TotalMilesTextBox.Text)
            totalCharge += CDec(TotalChargeTextBox.Text)

        End If
    End Sub
End Class
