Module Module1
    Sub inputNumber(ByVal userInput As Button)
        Form1.TextBox1.Text = Val(Form1.TextBox1.Text & userInput.Text)
    End Sub
    Sub getOperator(ByVal selectOperator As Button)
        Form1.userInputNumber = Val(Form1.TextBox1.Text)
        Form1.buttonOperator = selectOperator.Text
        Form1.TextBox1.Text = ""
        Form1.TextBox3.Text = Form1.userInputNumber & " " & selectOperator.Text

    End Sub
    Sub calculateValue()

        Select Case Form1.buttonOperator
            Case "+"
                Form1.TextBox1.Text = Form1.userInputNumber + Val(Form1.TextBox1.Text)
            Case "-"
                Form1.TextBox1.Text = Form1.userInputNumber - Val(Form1.TextBox1.Text)
            Case "/"
                Form1.TextBox1.Text = Form1.userInputNumber / Val(Form1.TextBox1.Text)
            Case "*"
                Form1.TextBox1.Text = Form1.userInputNumber * Val(Form1.TextBox1.Text)

        End Select
    End Sub

    Sub addComboBoxItems()
        'adding items to combobox1 before loading the form1
        Form1.ComboBox1.Items.Add("Angle")
        Form1.ComboBox1.Items.Add("Area")
        Form1.ComboBox1.Items.Add("Energy")
        Form1.ComboBox1.Items.Add("Length")
        Form1.ComboBox1.Items.Add("Power")
        Form1.ComboBox1.Items.Add("Pressure")
        Form1.ComboBox1.Items.Add("Temperature")
        Form1.ComboBox1.Items.Add("Time")
        Form1.ComboBox1.Items.Add("Velocity")
        Form1.ComboBox1.Items.Add("Volume")
        Form1.ComboBox1.Items.Add("Weight/Mass")
        Form1.ComboBox1.SelectedIndex = 0

    End Sub

    Sub conversionType()
        'using select case if the conversion type of combobox1 is a type of conversion combobox2 and combobox3
        Select Case Form1.ComboBox1.SelectedItem
            Case "Angle"
                ' combobox 2 
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Degree")
                Form1.ComboBox2.Items.Add("Gradian")
                Form1.ComboBox2.Items.Add("Radian")
                Form1.ComboBox2.SelectedIndex = (0)

                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Degree")
                Form1.ComboBox3.Items.Add("Gradian")
                Form1.ComboBox3.Items.Add("Radian")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""



            Case "Area"
                ' combobox 2 
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Acres")
                Form1.ComboBox2.Items.Add("Hectares")
                Form1.ComboBox2.Items.Add("Square Centimeter")
                Form1.ComboBox2.Items.Add("Square feet")
                Form1.ComboBox2.Items.Add("Square inch")
                Form1.ComboBox2.Items.Add("Square kilometer")
                Form1.ComboBox2.Items.Add("Square meters")
                Form1.ComboBox2.Items.Add("Square miles")
                Form1.ComboBox2.Items.Add("Square millimeter")
                Form1.ComboBox2.Items.Add("Square Yard")
                Form1.ComboBox2.SelectedIndex = (0)

                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Acres")
                Form1.ComboBox3.Items.Add("Hectares")
                Form1.ComboBox3.Items.Add("Square Centimeter")
                Form1.ComboBox3.Items.Add("Square feet")
                Form1.ComboBox3.Items.Add("Square inch")
                Form1.ComboBox3.Items.Add("Square kilometer")
                Form1.ComboBox3.Items.Add("Square meters")
                Form1.ComboBox3.Items.Add("Square miles")
                Form1.ComboBox3.Items.Add("Square millimeter")
                Form1.ComboBox3.Items.Add("Square Yard")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""


            Case "Energy"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("British Thermal Unit")
                Form1.ComboBox2.Items.Add("Calorie")
                Form1.ComboBox2.Items.Add("Electron-Volts")
                Form1.ComboBox2.Items.Add("Foot-Pound")
                Form1.ComboBox2.Items.Add("Joule")
                Form1.ComboBox2.Items.Add("Kilocalorie")
                Form1.ComboBox2.Items.Add("Kilojoule")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("British Thermal Unit")
                Form1.ComboBox3.Items.Add("Calorie")
                Form1.ComboBox3.Items.Add("Electron-Volts")
                Form1.ComboBox3.Items.Add("Foot-Pound")
                Form1.ComboBox3.Items.Add("Joule")
                Form1.ComboBox3.Items.Add("Kilocalorie")
                Form1.ComboBox3.Items.Add("Kilojoule")
                Form1.ComboBox3.SelectedIndex = (0)

                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""


            Case "Length"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Angstrom")
                Form1.ComboBox2.Items.Add("Centimeters")
                Form1.ComboBox2.Items.Add("Chain")
                Form1.ComboBox2.Items.Add("Fathom")
                Form1.ComboBox2.Items.Add("Feet")
                Form1.ComboBox2.Items.Add("Hand")
                Form1.ComboBox2.Items.Add("Inch")
                Form1.ComboBox2.Items.Add("Kilometers")
                Form1.ComboBox2.Items.Add("Link")
                Form1.ComboBox2.Items.Add("Meter")
                Form1.ComboBox2.Items.Add("Microns")
                Form1.ComboBox2.Items.Add("Mile")
                Form1.ComboBox2.Items.Add("Millimeters")
                Form1.ComboBox2.Items.Add("Nanometer")
                Form1.ComboBox2.Items.Add("Nautical Miles")
                Form1.ComboBox2.Items.Add("PICA")
                Form1.ComboBox2.Items.Add("Rods")
                Form1.ComboBox2.Items.Add("Span")
                Form1.ComboBox2.Items.Add("Yards")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Angstrom")
                Form1.ComboBox3.Items.Add("Centimeters")
                Form1.ComboBox3.Items.Add("Chain")
                Form1.ComboBox3.Items.Add("Fathom")
                Form1.ComboBox3.Items.Add("Feet")
                Form1.ComboBox3.Items.Add("Hand")
                Form1.ComboBox3.Items.Add("Inch")
                Form1.ComboBox3.Items.Add("Kilometers")
                Form1.ComboBox3.Items.Add("Link")
                Form1.ComboBox3.Items.Add("Meter")
                Form1.ComboBox3.Items.Add("Microns")
                Form1.ComboBox3.Items.Add("Mile")
                Form1.ComboBox3.Items.Add("Millimeters")
                Form1.ComboBox3.Items.Add("Nanometer")
                Form1.ComboBox3.Items.Add("Nautical Miles")
                Form1.ComboBox3.Items.Add("PICA")
                Form1.ComboBox3.Items.Add("Rods")
                Form1.ComboBox3.Items.Add("Span")
                Form1.ComboBox3.Items.Add("Yards")
                Form1.ComboBox3.SelectedIndex = (0)

                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""



            Case "Power"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("BTU/minute")
                Form1.ComboBox2.Items.Add("Foot-Pound/minute")
                Form1.ComboBox2.Items.Add("Horsepower")
                Form1.ComboBox2.Items.Add("Kilowat")
                Form1.ComboBox2.Items.Add("Watt")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("BTU/minute")
                Form1.ComboBox3.Items.Add("Foot-Pound/minute")
                Form1.ComboBox3.Items.Add("Horsepower")
                Form1.ComboBox3.Items.Add("Kilowat")
                Form1.ComboBox3.Items.Add("Watt")
                Form1.ComboBox3.SelectedIndex = (0)

                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""


            Case "Pressure"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Atmosphere")
                Form1.ComboBox2.Items.Add("Bar")
                Form1.ComboBox2.Items.Add("Kilo Pascal")
                Form1.ComboBox2.Items.Add("Millimeter of mercury")
                Form1.ComboBox2.Items.Add("Pascal")
                Form1.ComboBox2.Items.Add("Pound per square inch (PSI)")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Atmosphere")
                Form1.ComboBox3.Items.Add("Bar")
                Form1.ComboBox3.Items.Add("Kilo Pascal")
                Form1.ComboBox3.Items.Add("Millimeter of mercury")
                Form1.ComboBox3.Items.Add("Pascal")
                Form1.ComboBox3.Items.Add("Pound per square inch (PSI)")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""



            Case "Temperature"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Degree Celsius")
                Form1.ComboBox2.Items.Add("Degree Fahrenheit")
                Form1.ComboBox2.Items.Add("Kelvin")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Degree Celsius")
                Form1.ComboBox3.Items.Add("Degree Fahrenheit")
                Form1.ComboBox3.Items.Add("Kelvin")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""

            Case "Time"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Day")
                Form1.ComboBox2.Items.Add("Hour")
                Form1.ComboBox2.Items.Add("Microsecond")
                Form1.ComboBox2.Items.Add("Millisecond")
                Form1.ComboBox2.Items.Add("Minute")
                Form1.ComboBox2.Items.Add("Second")
                Form1.ComboBox2.Items.Add("Week")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Day")
                Form1.ComboBox3.Items.Add("Hour")
                Form1.ComboBox3.Items.Add("Microsecond")
                Form1.ComboBox3.Items.Add("Millisecond")
                Form1.ComboBox3.Items.Add("Minute")
                Form1.ComboBox3.Items.Add("Second")
                Form1.ComboBox3.Items.Add("Week")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""



            Case "Velocity"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Centimeter per second")
                Form1.ComboBox2.Items.Add("Feet per second")
                Form1.ComboBox2.Items.Add("Kilmoter per hour")
                Form1.ComboBox2.Items.Add("Knots")
                Form1.ComboBox2.Items.Add("Mach (at std. atm)")
                Form1.ComboBox2.Items.Add("Meter per second")
                Form1.ComboBox2.Items.Add("Miles per hour")
                Form1.ComboBox2.SelectedIndex = (0)



                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Centimeter per second")
                Form1.ComboBox3.Items.Add("Feet per second")
                Form1.ComboBox3.Items.Add("Kilmoter per hour")
                Form1.ComboBox3.Items.Add("Knots")
                Form1.ComboBox3.Items.Add("Mach (at std. atm)")
                Form1.ComboBox3.Items.Add("Meter per second")
                Form1.ComboBox3.Items.Add("Miles per hour")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""



            Case "Volume"
                ' combobox 2
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Cubic centimer")
                Form1.ComboBox2.Items.Add("Cubic feet")
                Form1.ComboBox2.Items.Add("Cubic inch")
                Form1.ComboBox2.Items.Add("Cubic meter")
                Form1.ComboBox2.Items.Add("Cubic yard")
                Form1.ComboBox2.Items.Add("Fluid ounce (UK)")
                Form1.ComboBox2.Items.Add("Fluid ounce (US)")
                Form1.ComboBox2.Items.Add("Gallon (UK)")
                Form1.ComboBox2.Items.Add("Gallon (US)")
                Form1.ComboBox2.Items.Add("Liter")
                Form1.ComboBox2.Items.Add("Pint (UK)")
                Form1.ComboBox2.Items.Add("Pint (US)")
                Form1.ComboBox2.Items.Add("Quart (UK)")
                Form1.ComboBox2.Items.Add("Quart (UK)")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Cubic centimer")
                Form1.ComboBox3.Items.Add("Cubic feet")
                Form1.ComboBox3.Items.Add("Cubic inch")
                Form1.ComboBox3.Items.Add("Cubic meter")
                Form1.ComboBox3.Items.Add("Cubic yard")
                Form1.ComboBox3.Items.Add("Fluid ounce (UK)")
                Form1.ComboBox3.Items.Add("Fluid ounce (US)")
                Form1.ComboBox3.Items.Add("Gallon (UK)")
                Form1.ComboBox3.Items.Add("Gallon (US)")
                Form1.ComboBox3.Items.Add("Liter")
                Form1.ComboBox3.Items.Add("Pint (UK)")
                Form1.ComboBox3.Items.Add("Pint (US)")
                Form1.ComboBox3.Items.Add("Quart (UK)")
                Form1.ComboBox3.Items.Add("Quart (UK)")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""


            Case "Weight/Mass"
                ' combobox 3
                Form1.ComboBox2.Items.Clear()

                Form1.ComboBox2.Items.Add("Carat")
                Form1.ComboBox2.Items.Add("Centigram")
                Form1.ComboBox2.Items.Add("Decigram")
                Form1.ComboBox2.Items.Add("Dekagram")
                Form1.ComboBox2.Items.Add("Gram")
                Form1.ComboBox2.Items.Add("Hectogram")
                Form1.ComboBox2.Items.Add("Kilogram")
                Form1.ComboBox2.Items.Add("Long ton")
                Form1.ComboBox2.Items.Add("Milligram")
                Form1.ComboBox2.Items.Add("Ounce")
                Form1.ComboBox2.Items.Add("Pound")
                Form1.ComboBox2.Items.Add("Short ton")
                Form1.ComboBox2.Items.Add("Stone")
                Form1.ComboBox2.Items.Add("Tonne")
                Form1.ComboBox2.SelectedIndex = (0)


                ' combobox 3
                Form1.ComboBox3.Items.Clear()

                Form1.ComboBox3.Items.Add("Carat")
                Form1.ComboBox3.Items.Add("Centigram")
                Form1.ComboBox3.Items.Add("Decigram")
                Form1.ComboBox3.Items.Add("Dekagram")
                Form1.ComboBox3.Items.Add("Gram")
                Form1.ComboBox3.Items.Add("Hectogram")
                Form1.ComboBox3.Items.Add("Kilogram")
                Form1.ComboBox3.Items.Add("Long ton")
                Form1.ComboBox3.Items.Add("Milligram")
                Form1.ComboBox3.Items.Add("Ounce")
                Form1.ComboBox3.Items.Add("Pound")
                Form1.ComboBox3.Items.Add("Short ton")
                Form1.ComboBox3.Items.Add("Stone")
                Form1.ComboBox3.Items.Add("Tonne")
                Form1.ComboBox3.SelectedIndex = (0)


                Form1.TextBox4.Text = ""
                Form1.TextBox5.Text = ""




        End Select
    End Sub

    Sub calculateConversionType()

        Select Case Form1.ComboBox1.SelectedItem







            'conversion type of Angle






            Case "Angle"
                'Degree
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text

                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.1111111111111109
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0174532925199433
                End If
                'Gradian
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.9
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.015707963267949
                End If
                'Radian
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 57.29577951308233
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 63.66197723675814
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If










                'conversion type of Area








            Case "Area"
                'Acres
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.40468564224
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 40468564.224
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 43560
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6272640
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0040468564224
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4046.8564224
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0015625
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4046856422.4
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4840
                End If
                'Hectares
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.4710538146716532
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 107639.1041670972
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 15500031.000062
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0038610215854245
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 11959.9004630108
                End If
                'Square centimeter
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000024710538146716529
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000001
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001076391041671
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.15500031000062
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000038610215854244577
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000119599004630108
                End If
                'Square feet
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00002295684113865932
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000009290304
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 929.0304
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 144
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000009290304
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.09290304
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000035870064279155189
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 92903.04
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1111111111111111
                End If
                'Square inch
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000015942250790735641
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000064516
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.4516
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0069444444444444
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000064516
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00064516
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000024909766860524441
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 645.16
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0007716049382716049
                End If
                'Square kilometer
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 247.1053814671653
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10763910.416709719
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1550003100.0062
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.38610215854244578
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1195990.04630108
                End If
                'Square meters
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00024710538146716532
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10.76391041670972
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1550.0031000062
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000038610215854244582
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.19599004630108
                End If
                'Square miles
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 640
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 258.9988110336
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 25899881103.36
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 27878400
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4014489600
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.589988110336
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2589988.110336
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2589988110336
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3097600
                End If
                'Square millimeter
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000024710538146716532
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00001076391041670972
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0015500031000062
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000000038610215854244582
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000119599004630108
                End If
                'Square yard
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00020661157024793391
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000083612736
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 8361.2736
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox3.Text = Val(Form1.TextBox4.Text) * 1296
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000083612736
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.83612736
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000032283057851239672
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 836127.36
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If









                'conversion type of Energy










            Case "Energy"
                'British thermal unit
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 251.9957963122194
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.5851420255170006E+21
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 778.16937096787467
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1055.056
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.25199579631221941
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.055056
                End If
                'Calorie
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.003968320164996
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.61319518892216E+19
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.0880252065940561
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.1868
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0041868
                End If
                'Electron-Volts
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.518570132770204E-22
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.8267328986338009E-20
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.1817047649883909E-19
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.60217653E-19
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.8267328986338007E-23
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.60217653E-22
                End If
                'Foot found
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0012850672839464
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.32383155353286519
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 8.4623505771327212E+18
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.3558179483314
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00032383155353286522
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0013558179483314
                End If
                'Joule
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00094781698791343778
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.23884589662749589
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.2415094796077179E+18
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.7375621492772656
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00023884589662749589
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                'Kilocalorie
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.9683201649959812
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.61319518892216E+22
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3088.0252065940558
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4186.8
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.1868
                End If
                'Kilojoule
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.94781698791343783
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 238.8458966274959
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.2415094796077184E+21
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 737.56214927726558
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.23884589662749589
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If








                'conversion type of Length







            Case "Length"
                'Angstrom
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000001
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000000049709695378986723
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000054680664916885391
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000032808398950131231
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000984251968503937
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000039370078740157481
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000049709695378986718
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000000000621371192237334
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000001
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000000053995680345572351
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000023710630158366141
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000019883878151594689
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000043744531933508308
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000043744531933508308
                End If
                'Centimeters
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00049709695378986722
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0054680664916885
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0328083989501312
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0984251968503937
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.39370078740157483
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00001
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0497096953789867
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000621371192237334
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000053995680345572348
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.3710630158366142
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0019883878151595
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0437445319335083
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0109361329833771
                End If
                'Chain
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 201168000000
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2011.68
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 11
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 66
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 198
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 792
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0201168
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 20.1168
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 20116800
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0125
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 20116.8
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 20116800000
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0108622030237581
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4769.8200476982
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 88
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 22
                End If
                'Fathom
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 18288000000
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 182.88
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0909090909090909
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 18
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 72
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0018288
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9.0909090909090917
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.8288
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1828800
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0011363636363636
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1828.8
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1828800000
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00098747300215982722
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 433.6200043362
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.36363636363636359
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 8
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2
                End If
                'Feet
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3048000000
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 30.48
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0151515151515152
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.16666666666666671
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 12
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0003048
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.5151515151515149
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.3048
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 304800
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00018939393939393939
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 304.8
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 304800000
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00016457883369330449
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 72.2700007227
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0606060606060606
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.333333333333333
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.33333333333333331
                End If
                'Hand
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1016000000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10.16
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0050505050505051
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0555555555555556
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.33333333333333331
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001016
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.50505050505050508
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1016
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 101600
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000063131313131313131
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 101.6
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 101600000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000054859611231101511
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 24.0900002409
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0202020202020202
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.44444444444444442
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1111111111111111
                End If
                'Inch
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 254000000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.54
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0012626262626263
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0138888888888889
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0833333333333333
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.25
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000254
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1262626262626263
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0254
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 25400
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000015782828282828279
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 25.4
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 25400000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000013714902807775379
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.022500060225001
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0050505050505051
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1111111111111111
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0277777777777778
                End If
                'Kilometer
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 49.709695378986723
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 546.80664916885394
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3280.8398950131232
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9842.51968503937
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 39370.078740157478
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4970.9695378986717
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.621371192237334
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.5399568034557235
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 237106.30158366141
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 198.83878151594689
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4374.4531933508306
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1093.6132983377081
                End If
                'Link
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2011680000
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 20.1168
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.11
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.66
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.98
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 7.92
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000201168
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.201168
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 201168
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000125
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 201.168
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 201168000
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000108622030237581
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 47.698200476982
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.04
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.88
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.22
                End If
                'Meter
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0497096953789867
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.54680664916885391
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.280839895013123
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9.84251968503937
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 39.370078740157481
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.9709695378986716
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.9709695378986716
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000621371192237334
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00053995680345572347
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 237.10630158366141
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.19883878151594689
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.3744531933508313
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.093613298337708
                End If
                'Microns
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000049709695378986719
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000054680664916885388
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000032808398950131231
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000984251968503937
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000039370078740157478
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000049709695378986717
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000621371192237334
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000053995680345572349
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00023710630158366139
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000001988387815159469
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000004374453193350831
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001093613298337708
                End If
                'Mile
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 16093440000000
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 160934.4
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 80
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 880
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5280
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 15840
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 63360
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.609344
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 8000
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1609.344
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1609344000
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1609344
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1609344000000
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.86897624190064793
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 381585.603815856
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 320
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 7040
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1760
                End If
                'Millimeter
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000000
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000049709695378986721
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00054680664916885392
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0032808398950131
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0098425196850394
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0393700787401575
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0049709695378987
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000621371192237334
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000053995680345572355
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.23710630158366139
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00019883878151594691
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0043744531933508
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0010936132983377
                End If
                'Nanometer
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000001
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000049709695378986718
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000546806649168854
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000032808398950131229
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000984251968503937
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000039370078740157479
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000049709695378986716
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000001
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000000621371192237334
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000000053995680345572354
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000023710630158366141
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000001988387815159469
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000043744531933508308
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000010936132983377079
                End If
                'Nautical miles
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 18520000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 185200
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 92.0623558418834
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1012.685914260717
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6076.1154855643044
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 18228.346456692911
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 72913.385826771657
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.852
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9206.23558418834
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1852
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1852000000
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.1507794480235429
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1852000
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1852000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 439120.870532941
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 368.24942336753361
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 8101.4873140857389
                End If
                If Form1.ComboBox2.SelectedIndex = 14 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2025.371828521435
                End If
                'PICA
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 42175176
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.42175176
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00020965151515151521
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0023061666666667
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.013837
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.041511
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.166044
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000042175176
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0209651515151515
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0042175176
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4217.5176
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000026206439393939391
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.2175176
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4217517.6
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000022772773218142549
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00083860606060606063
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0184493333333333
                End If
                If Form1.ComboBox2.SelectedIndex = 15 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0046123333333333
                End If
                'Rods
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 50292000000
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 502.92
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.25
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.75
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 16.5
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 49.5
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 198
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0050292
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 25
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5.0292
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5029200
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.003125
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5029.2
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5029200000
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0027155507559395
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1192.45501192455
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 22
                End If
                If Form1.ComboBox2.SelectedIndex = 16 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5.5
                End If
                'Span
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2286000000
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 22.86
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0113636363636364
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.125
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.75
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.25
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0002286
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.136363636363636
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.2286
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 228600
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00014204545454545451
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 228.6
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 228600000
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001234341252699784
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 54.202500542025007
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0454545454545455
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 17 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.25
                End If
                'Yard
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9144000000
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 91.44
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0454545454545455
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.5
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 36
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0009144
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.545454545454545
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.9144
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 914400
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00056818181818181815
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 914.4
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 914400000
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 14 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00049373650107991361
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 15 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 216.8100021681
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 16 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1818181818181818
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 17 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4
                End If
                If Form1.ComboBox2.SelectedIndex = 18 And Form1.ComboBox3.SelectedIndex = 18 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If










                'conversion type of Power






            Case "Power"
                'BTU/minute
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 778.16937096787467
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0235808900293295
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0175842666666667
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 17.584266666666672
                End If
                'Foot-pound/minute
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0012850672839464
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000030303030303030289
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00002259696580552333
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0225969658055233
                End If
                'Horsepower
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 42.407220370232679
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 33000.000000000007
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.74569987158227025
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 745.69987158227025
                End If
                'Kilowatt
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 56.86901927480627
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 44253.728956635932
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.341022089595028
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                'Watt
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0568690192748063
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 44.253728956635932
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001341022089595
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If









                '
                'conversion type of pressure







            Case "Pressure"
                'Atmosphere
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.01325
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 101.325
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 760.12753188297074
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 101325
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 14.695949400392211
                End If
                'Bar
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.98692326671601283
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 750.18754688672175
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 14.50377438972831
                End If
                'Kilo Pascal
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0098692326671601
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 7.5018754688672171
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.14503774389728311
                End If
                'Millimeter of mercury
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0013155687145324
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001333
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1333
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 133.3
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0193335312615078
                End If
                'Pascal
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000098692326671601285
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00001
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0075018754688672
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00014503774389728309
                End If
                'Pound per square inch (PSI)
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.068045961016531
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.06894757
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.894757
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 51.723608402100531
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6894.757
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If











                'conversion type of Temperature









            Case "Temperature"
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then

                    Form1.TextBox5.Text = 32
                    Form1.TextBox5.Text += Val(Form1.TextBox4.Text) ' WRONG FORMULA
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) + 273.15
                End If
                'Degree Farenheit
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = -17.777777777777779
                    Form1.TextBox5.Text += Val(Form1.TextBox4.Text) / 2 ' WRONG FORMULA
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = 255.37222222222221
                    Form1.TextBox5.Text += Form1.TextBox4.Text  ' WRONG FORMULA
                End If
                'Kelvin
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) - 273.15
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = -459.67
                    Form1.TextBox5.Text += (Val(Form1.TextBox4.Text) + 0.8) ' WRONG FORMULA
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If











                'conversion type of Time






            Case "Time"

                'Day
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 24
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 86400000000
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 86400000
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1440
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 86400
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1428571428571429
                End If
                'Hour
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0416666666666667
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3600000000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3600000
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 60
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3600
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.005952380952381
                End If
                'Microsecond
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000001157407407407407
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000027777777777777782
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000001666666666666667
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000000001653439153439153
                End If
                'Millisecond
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000011574074074074071
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000027777777777777781
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000016666666666666671
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000016534391534391531
                End If
                'Minute
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00069444444444444436
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0166666666666667
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 60000000
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 60000
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 60
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000992063492063492
                End If
                'Second
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00001157407407407407
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00027777777777777778
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0166666666666667
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000016534391534391531
                End If
                'Week
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 7
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 168
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 604800000000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 604800000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10080
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 604800
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If











                ' Conversion type of Velocity







            Case "Velocity"
                'Centimeter per second
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0328083989501312
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.036
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.019438444924406
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000029386414601756779
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.022369362920544
                End If
                'Feet per second
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 30.48
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.09728
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.59248380129589628
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00089569791706154659
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.3048
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.68181818181818177
                End If
                'kilometer per hour
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 27.777777777777779
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.91134441528142318
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.5399568034557235
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00081628929449324391
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.27777777777777779
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.621371192237334
                End If
                'Knots
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 51.444444444444443
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.6878098571011959
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.852
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0015117677734015
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.51444444444444437
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.1507794480235429
                End If
                'Mach
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 34029.33
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1116.447834645669
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1225.05588
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 661.47725701943841
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 340.2933
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 761.2144327129563
                End If
                'Meter per second
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.280839895013123
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.6
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.9438444924406051
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0029386414601757
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.236936292054402
                End If
                'Miles per hour
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 44.704
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.466666666666667
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.609344
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.86897624190064793
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0013136902783569
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.44704
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If








                'conversion type of Volume







            Case "Volume"
                'Cubic centimeter
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000035314666721488593
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0610237440947323
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000013079506193143921
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.035195079727854
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.033814022701843
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00021996924829908779
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00026417205235814838
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0017597539863927
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0021133764188652
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00087987699319635115
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0010566882094326
                End If
                'Cubic feet
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 28316.846592
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1728
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.028316846592
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.037037037037037
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 996.61367344685209
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 957.50649350649348
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.2288354590428261
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 7.4805194805194812
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 28.316846592
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 49.830683672342609
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 59.844155844155843
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 24.9153418361713
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 29.922077922077921
                End If
                'Cubic inch
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 16.387064
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00057870370370370367
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000016387064
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00002143347050754458
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.57674402398544677
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.55411255411255411
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.003604650149909
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0043290043290043
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.016387064
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0288372011992723
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0346320346320346
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0144186005996362
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0173160173160173
                End If
                'Cubic meter
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 35.314666721488592
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 61023.744094732283
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.3079506193143919
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 33814.022701843
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 219.96924829908781
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 264.17205235814839
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1759.753986392702
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1759.753986392702
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2113.3764188651871
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 879.87699319635124
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1056.688209432594
                End If
                'Cubic yard
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 764554.857984
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 27
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 46656
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.764554857984
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 26908.56918306501
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 25852.675324675321
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 168.17855739415629
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 201.974025974026
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 764.554857984
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1345.42845915325
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1615.792207792208
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 672.71422957662514
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 807.89610389610391
                End If
                'Fluid ounce (UK)
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 28.4130625
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0010033978327243
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.7338714549476339
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000284130625
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000037162882693493528
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.96075994040388391
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00625
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0075059370344053
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0284130625
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.05
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0600474962752427
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.025
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0300237481376214
                End If
                'Fluid ounce (US)
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 29.5735295625
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0010443793402778
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.8046875
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000295735295625
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000038680716306584362
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.040842730786236
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.006505267067414
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0078125
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0295735295625
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0520421365393118
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0625
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0260210682696559
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.03125
                End If
                'Gallon (UK)
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4546.09
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.16054365323589209
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 277.41943279162149
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00454609
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.005946061230959
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 160
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 153.72159046462141
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.200949925504855
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.54609
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 8
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9.6075994040388384
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.80379970201942
                End If
                'Gallon (US)
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3785.411784
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.13368055555555561
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 231
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.003785411784
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0049511316872428
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 133.22786954063821
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 128
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.8326741846289889
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.785411784
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.6613934770319112
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 8
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.3306967385159552
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4
                End If
                'Liter
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0353146667214886
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 61.02374409473228
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0013079506193144
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 35.195079727854051
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 33.814022701843
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.21996924829908779
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.26417205235814839
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.7597539863927021
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.1133764188651871
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.87987699319635115
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.056688209432594
                End If
                'Pint (UK)
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 568.26125
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0200679566544865
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 34.677429098952693
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00056826125
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00074325765386987067
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 20
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 19.21519880807768
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.125
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.15011874068810691
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.56826125
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.200949925504855
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.5
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.60047496275242751
                End If
                'Pint (US)
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 473.176473
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0167100694444444
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 28.875
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000473176473
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0006188914609053498
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 16.65348369257978
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 16
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1040842730786236
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.125
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.473176473
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.8326741846289889
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.41633709231449439
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.5
                End If
                'Quart (UK)
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1136.5225
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.040135913308973
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 69.354858197905372
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0011365225
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0014865153077397
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 40
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 38.430397616155361
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.25
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.3002374813762137
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.1365225
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.40189985100971
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.200949925504855
                End If
                'Quart (US)
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 946.352946
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0334201388888889
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 57.75
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000946352946
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0012377829218107
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 33.306967385159552
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 32
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.2081685461572472
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.25
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.946352946
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.665348369257978
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.8326741846289889
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If








                'conversion type of Weight/Mass










            Case "Weight/Mass"
                'Carat
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 20
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.02
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.2
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.002
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0002
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000019684130552221209
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 200
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00044092452436975523
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.665348369257978
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00044092452436975523
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000031494608883553942
                End If
                If Form1.ComboBox2.SelectedIndex = 0 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000002
                End If
                'Centigram
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.05
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00001
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000009842065276110606
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00035273961949580407
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000022046226218487761
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000011023113109243881
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001574730444177697
                End If
                If Form1.ComboBox2.SelectedIndex = 1 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000001
                End If
                'Decigram
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.5
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000009842065276110606
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.003527396194958
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00022046226218487761
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000001102311310924388
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000015747304441776971
                End If
                If Form1.ComboBox2.SelectedIndex = 2 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000001
                End If
                'Dekagram
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 50
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000098420652761106056
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.35273961949580412
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0220462262184878
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000011023113109243881
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0015747304441777
                End If
                If Form1.ComboBox2.SelectedIndex = 3 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00001
                End If
                'Gram
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000098420652761106068
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0352739619495804
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0022046226218488
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001102311310924388
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00015747304441776969
                End If
                If Form1.ComboBox2.SelectedIndex = 4 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                'Hectogram
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 500
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000098420652761106056
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 3.5273961949580408
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.22046226218487761
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00011023113109243881
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.015747304441777
                End If
                If Form1.ComboBox2.SelectedIndex = 5 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                'Kilogram
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00098420652761106058
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 35.273961949580411
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.2046226218487761
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0011023113109244
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.15747304441776969
                End If
                If Form1.ComboBox2.SelectedIndex = 6 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                'Long ton
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5080234.544
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 101604690.88
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10160469.088
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 101604.69088
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1016046.9088
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10160.469088
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1016.0469088
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1016046908.8
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 35840
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2240
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.12
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 160
                End If
                If Form1.ComboBox2.SelectedIndex = 7 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.0160469088
                End If
                'Milligram
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.005
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.1
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.01
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0001
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.001
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00001
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000001
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00000000098420652761106052
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000035273961949580409
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000002204622621848776
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000000011023113109243879
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0000001574730444177697
                End If
                If Form1.ComboBox2.SelectedIndex = 8 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000000001
                End If
                'Ounce
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 141.747615625
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2834.9523125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 283.49523125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2.8349523125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 28.349523125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.28349523125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.028349523125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00002790178571428571
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 28349.523125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0625
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00003125
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0044642857142857
                End If
                If Form1.ComboBox2.SelectedIndex = 9 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.000028349523125
                End If
                'Pound
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2267.96185
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 45359.237
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4535.9237
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 45.359237
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 453.59237
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4.5359237
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.45359237
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00044642857142857141
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 453592.37
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 16
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0005
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.0714285714285714
                End If
                If Form1.ComboBox2.SelectedIndex = 10 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00045359237
                End If
                'Short ton
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 4535923.7
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 90718474
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9071847.4
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 90718.474
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 907184.74
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 9071.8474
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 907.18474
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.8928571428571429
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 907184740
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 32000
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2000
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 142.85714285714289
                End If
                If Form1.ComboBox2.SelectedIndex = 11 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.90718474
                End If
                'Stone
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 31751.4659
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 635029.318
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 63502.9318
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 635.029318
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6350.29318
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 63.5029318
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6.35029318
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00625
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 6350293.18
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 224
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 14
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.007
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If
                If Form1.ComboBox2.SelectedIndex = 12 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.00635029318
                End If
                'Tonne
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 0 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 5000000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 1 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 2 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 3 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 100000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 4 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 5 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 10000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 6 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 7 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 0.98420652761106064
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 8 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1000000000
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 9 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 35273.961949580407
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 10 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 2204.6226218487759
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 11 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 1.1023113109243881
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 12 Then
                    Form1.TextBox5.Text = Val(Form1.TextBox4.Text) * 157.4730444177697
                End If
                If Form1.ComboBox2.SelectedIndex = 13 And Form1.ComboBox3.SelectedIndex = 13 Then
                    Form1.TextBox5.Text = Form1.TextBox4.Text
                End If




        End Select
    End Sub


End Module

