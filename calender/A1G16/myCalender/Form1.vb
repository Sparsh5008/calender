Public Class Form1
    
    Function isLeap(year As Integer)

        'IF THE YEAR IS DIVISIBLE BY 400 THEN IT IS LEAP YEAR
        'A NON CENTURY YEAR DIVISIBLE BY 4 IS ALSO A LEAP YEAR

        If year Mod 400 = 0 Then
            Return True
        ElseIf year Mod 100 = 0 Then
            Return False
        ElseIf year Mod 4 = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Sub fillBoxes(txt() As TextBox, start As Integer, str As String)

        'THIS FUNCTION WHEN CALLED FILLS AN ARRAY
        'OF TEXTBOXES CONSECUTIVELY FROM WEEK DAY AT START


        Dim i As Integer = -1
        For Each tb As TextBox In txt
            i = i + 1
            tb.Text = weekDay((start + i) Mod 7)
            If (tb.Text = str) Then
                tb.ForeColor = Color.Red
            End If
        Next
    End Sub
    Function dayOfWeek(year As Integer, month As Integer, day As Integer)
        'Index     Day 
        '0         Sunday 
        '1         Monday 
        '2         Tuesday 
        '3         Wednesday 
        '4         Thursday 
        '5         Friday 
        '6         Saturday


        Dim array() As Integer = {0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4}
        If month < 3 Then
            year = year - 1
        End If
        Dim weekday As Integer
        'THIS FORMULA CALCULATES THE WEEKDAY FROM A GIVEN DATE
        weekday = (year + (year \ 4) - (year \ 100) + (year \ 400) + array(month - 1) + day) Mod 7

        If weekday < 0 Then
            weekday = weekday + 7
        End If
        Return weekday
    End Function
    Function weekDay(num As Integer)
        If num = 0 Then
            Return "SUN"
        ElseIf num = 1 Then
            Return "MON"
        ElseIf num = 2 Then
            Return "TUE"
        ElseIf num = 3 Then
            Return "WED"
        ElseIf num = 4 Then
            Return "THU"
        ElseIf num = 5 Then
            Return "FRI"
        ElseIf num = 6 Then
            Return "SAT"
        Else
            Return "NONE"
        End If
    End Function
    Function monthNum(month As String)

        'TAKES MONTH AS INPUT AND RETURNS THE CORRESPONDING NUMBER

        If month = "January" Then
            Return 1
        ElseIf month = "February" Then
            Return 2
        ElseIf month = "March" Then
            Return 3
        ElseIf month = "April" Then
            Return 4
        ElseIf month = "May" Then
            Return 5
        ElseIf month = "June" Then
            Return 6
        ElseIf month = "July" Then
            Return 7
        ElseIf month = "August" Then
            Return 8
        ElseIf month = "September" Then
            Return 9
        ElseIf month = "October" Then
            Return 10
        ElseIf month = "November" Then
            Return 11
        ElseIf month = "December" Then
            Return 12
        Else
            Return -1
        End If
    End Function
    Sub initializeColor(txt() As TextBox)

        'INITIALIZING THE COLOR OF ALL THE TEXTBOXES BEFORE FILLING THEM

        For Each tb As TextBox In txt
            tb.ForeColor = Color.Olive
        Next
    End Sub
    Function monthFromNumber(month As Integer)

        'TAKES MONTH AS INPUT AND RETURNS THE NAME
        If month = 1 Then
            Return "January"
        ElseIf month = 2 Then
            Return "February"
        ElseIf month = 3 Then
            Return "March"
        ElseIf month = 4 Then
            Return "April"
        ElseIf month = 5 Then
            Return "May"
        ElseIf month = 6 Then
            Return "June"
        ElseIf month = 7 Then
            Return "July"
        ElseIf month = 8 Then
            Return "August"
        ElseIf month = 9 Then
            Return "September"
        ElseIf month = 10 Then
            Return "October"
        ElseIf month = 11 Then
            Return "November"
        ElseIf month = 12 Then
            Return "December"
        Else
            Return "F"
        End If
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'STORING THE ROWS IN ARRAYS
        Dim txt1() As TextBox = {TextBox59, TextBox60, TextBox61, TextBox62, TextBox63, TextBox64, TextBox65}
        Dim txt2() As TextBox = {TextBox79, TextBox78, TextBox77, TextBox76, TextBox75, TextBox74, TextBox66}
        Dim txt3() As TextBox = {TextBox86, TextBox85, TextBox84, TextBox83, TextBox82, TextBox81, TextBox80}
        Dim txt4() As TextBox = {TextBox93, TextBox92, TextBox91, TextBox90, TextBox89, TextBox88, TextBox87}
        Dim txt5() As TextBox = {TextBox100, TextBox99, TextBox98, TextBox97, TextBox96, TextBox95, TextBox94}
        Dim txt6() As TextBox = {TextBox107, TextBox106, TextBox105, TextBox104, TextBox103, TextBox102, TextBox101}
        Dim txt7() As TextBox = {TextBox114, TextBox113, TextBox112, TextBox111, TextBox110, TextBox109, TextBox108}
        Dim year, firstDay As Integer
        Try
            Dim str = "SUN"
            year = CInt(TextBox67.Text)
            firstDay = dayOfWeek(year, 1, 1) 'calculate firstday
            If year < 0 Then
                MessageBox.Show("Year can't be negative ! ", "ERROR")
            ElseIf TextBox67.Text.Contains(".") Then
                MessageBox.Show("Decimals not allowed ")
            ElseIf isLeap(year) = False Then
                'if the year is not leap then
                TextBox68.Text = "NON LEAP"
                TextBox1.Text = "CALENDER OF " & year
                TextBox3.Text = "JAN"
                TextBox10.Text = "OCT"
                TextBox4.Text = "MAY"
                TextBox5.Text = "AUG"
                TextBox6.Text = "FEB"
                TextBox13.Text = "MAR"
                TextBox20.Text = "NOV"
                TextBox7.Text = "JUN"
                TextBox8.Text = "SEP"
                TextBox15.Text = "DEC"
                TextBox9.Text = "APR"
                TextBox16.Text = "JUL"
                initializeColor(txt1)
                initializeColor(txt2)
                initializeColor(txt3)
                initializeColor(txt4)
                initializeColor(txt5)
                initializeColor(txt6)
                initializeColor(txt7)
                fillBoxes(txt1, firstDay, str)
                fillBoxes(txt2, firstDay + 1, str)
                fillBoxes(txt3, firstDay + 2, str)
                fillBoxes(txt4, firstDay + 3, str)
                fillBoxes(txt5, firstDay + 4, str)
                fillBoxes(txt6, firstDay + 5, str)
                fillBoxes(txt7, firstDay + 6, str)
            Else
                'if the year is leap then..
                TextBox68.Text = "LEAP"
                TextBox1.Text = "CALENDER OF " & year
                TextBox3.Text = "JAN"
                TextBox10.Text = "APR"
                TextBox17.Text = "JULY"
                TextBox4.Text = "OCT"
                TextBox5.Text = "MAY"
                TextBox6.Text = "FEB"
                TextBox13.Text = "AUG"
                TextBox7.Text = "MAR"
                TextBox14.Text = "NOV"
                TextBox8.Text = "JUN"
                TextBox9.Text = "SEP"
                TextBox16.Text = "DEC"
                initializeColor(txt1)
                initializeColor(txt2)
                initializeColor(txt3)
                initializeColor(txt4)
                initializeColor(txt5)
                initializeColor(txt6)
                initializeColor(txt7)
                fillBoxes(txt1, firstDay, str)
                fillBoxes(txt2, firstDay + 1, str)
                fillBoxes(txt3, firstDay + 2, str)
                fillBoxes(txt4, firstDay + 3, str)
                fillBoxes(txt5, firstDay + 4, str)
                fillBoxes(txt6, firstDay + 5, str)
                fillBoxes(txt7, firstDay + 6, str)
            End If
        Catch ex As Exception
            MessageBox.Show("INVALID INPUT", "ERROR")
        End Try

       
    End Sub
    Sub HighLight(day As String, txt() As TextBox)
        'HIGHLIGHTING WEEKDAY   
        For Each tb As TextBox In txt
            If tb.Text = day Then
                tb.ForeColor = Color.Red
            End If
        Next
    End Sub
    Function printCal(month As Integer, year As Integer)

        'calculate the first day of the month
        Dim cal As String
        Dim firstDay As Integer
        Dim totalDays As Integer
        firstDay = dayOfWeek(year, month, 1)
        Form2.Label1.Text = monthFromNumber(month) & " " & year

        cal = " S    M     T    W    Th    F    S" & Environment.NewLine

        'number of spaces is twice that of the number corresponding to the first day
        For i As Integer = 1 To firstDay
            cal = cal & "       "
        Next
        totalDays = numOfDays(month, year)

        For i = 1 To totalDays
            If (i + firstDay) Mod 7 = 1 And i > 1 Then
                cal = cal & Environment.NewLine
            End If
            If (i > 9) Then
                cal = cal & i & "   "
            Else
                cal = cal & "  " & i & "   "
            End If
        Next

        Return cal

    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'storing rows as arrays
        Dim txt1() As TextBox = {TextBox59, TextBox60, TextBox61, TextBox62, TextBox63, TextBox64, TextBox65}
        Dim txt2() As TextBox = {TextBox79, TextBox78, TextBox77, TextBox76, TextBox75, TextBox74, TextBox66}
        Dim txt3() As TextBox = {TextBox86, TextBox85, TextBox84, TextBox83, TextBox82, TextBox81, TextBox80}
        Dim txt4() As TextBox = {TextBox93, TextBox92, TextBox91, TextBox90, TextBox89, TextBox88, TextBox87}
        Dim txt5() As TextBox = {TextBox100, TextBox99, TextBox98, TextBox97, TextBox96, TextBox95, TextBox94}
        Dim txt6() As TextBox = {TextBox107, TextBox106, TextBox105, TextBox104, TextBox103, TextBox102, TextBox101}
        Dim txt7() As TextBox = {TextBox114, TextBox113, TextBox112, TextBox111, TextBox110, TextBox109, TextBox108}
        'initialize colors first
        initializeColor(txt1)
        initializeColor(txt2)
        initializeColor(txt3)
        initializeColor(txt4)
        initializeColor(txt5)
        initializeColor(txt6)
        initializeColor(txt7)

        Dim weekDay As String
        weekDay = ComboBox1.Text
        Try
            If weekDay = "" Then
                MessageBox.Show("You didnot enter any day")
            Else
                'highlighting the given weekday
                HighLight(weekDay, txt1)
                HighLight(weekDay, txt2)
                HighLight(weekDay, txt3)
                HighLight(weekDay, txt4)
                HighLight(weekDay, txt5)
                HighLight(weekDay, txt6)
                HighLight(weekDay, txt7)
            End If

        Catch ex As Exception
            MessageBox.Show("Some error occured")
        End Try
      
    End Sub
    Function numOfDays(month As Integer, year As Integer)
        If month = 1 Or month = 3 Or month = 5 Or month = 7 Or month = 8 Or month = 10 Or month = 12 Then
            Return 31
        ElseIf month = 4 Or month = 6 Or month = 9 Or month = 11 Then
            Return 30
        Else
            If isLeap(year) = True Then
                Return 29
            Else
                Return 28
            End If
        End If

    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim year, month As Integer
        Try

            month = monthNum(ComboBox2.Text)
            If TextBox67.Text = "" Then
                'if no year is entered display the message to enter year
                MessageBox.Show("enter year first")
            Else
                year = CInt(TextBox67.Text)
                If year < 0 Then
                    MessageBox.Show("year cannot be negative")
                ElseIf TextBox67.Text.Contains(".") Then
                    MessageBox.Show("decimals not allowed")
                Else
                    If month = -1 Then
                        'it means that nothing was selected
                        MessageBox.Show("No month selected")
                    Else

                        Form2.RichTextBox1.Text = printCal(month, year)
                        Form2.Show()
                    End If

                End If
            End If
           
        Catch ex As Exception
            MessageBox.Show("SORRY! INVALID INPUTS")
        End Try

    End Sub
End Class
