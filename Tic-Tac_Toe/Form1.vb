'Project Name: Tic Tac Toe
'Public Class Main
'Author: ds

'Possible Fixes:
'Clean up interface
'Use loop to parse CheckWin function
'Use global error handler
'less redundant code in each button click handler


Public Class Form1
    '---------------------------------
    ' x,o,- as current set,default unset
    'Dim PB_chars() As Char = {"_", "_", "_", "_", "_", "_", "_", "_", "_"}
    Dim PB_chars(9) As Char

    ' 1-2-3
    ' 4-5-6
    ' 7-8-9

    'Which players turn,also known as 'side' in functions
    Dim turn As Char = "x"

    'Integer to check for Draw
    Dim enderCount As Integer = 0
    '---------------------------------
    '-------------------------------------
    'CheckWin - Checks rows for possible win
    'Horizontal,Vertical,2 diagonals
    'calls EndGame if side returns true for HVD checks
    'side is current sides turn, X or O

    Public Sub CheckWin(ByVal side As Char)


        'Check Horizontal
        If (PB_chars(1) = side) And (PB_chars(2) = side) And (PB_chars(3) = side) Then
            EndGame(side)
        ElseIf (PB_chars(4) = side) And (PB_chars(5) = side) And (PB_chars(6) = side) Then
            EndGame(side)
        ElseIf (PB_chars(7) = side) And (PB_chars(8) = side) And (PB_chars(9) = side) Then
            EndGame(side)
        End If

        'Vertical
        If (PB_chars(1) = side) And (PB_chars(4) = side) And (PB_chars(7) = side) Then
            EndGame(side)
        ElseIf (PB_chars(2) = side) And (PB_chars(5) = side) And (PB_chars(8) = side) Then
            EndGame(side)
        ElseIf (PB_chars(3) = side) And (PB_chars(6) = side) And (PB_chars(9) = side) Then
            EndGame(side)
        End If

        'Diagonal 
        If (PB_chars(1) = side) And (PB_chars(5) = side) And (PB_chars(9) = side) Then
            EndGame(side)
        ElseIf (PB_chars(3) = side) And (PB_chars(5) = side) And (PB_chars(7) = side) Then
            EndGame(side)
        End If

        '---------------------------------------
        'To check for possible draw,
        'will only check for draw when all 9 images are clicked
        '
        For j As Integer = 1 To 9

            If ((PB_chars(j) = "x") Or (PB_chars(j) = "o")) Then
                enderCount = enderCount + 1
            End If
        Next j
        'MessageBox.Show(enderCount & "= ender count", "Window")
        If (enderCount = 9) Then
            'game declared draw
            enderCount = 0
            EndGame("-")
        Else
            enderCount = 0
        End If
        '----------------------------------------
    End Sub

    '-------------------------------------------
    'End of Game function
    'Shows MessageBox with winner
    'resets variables

    Public Sub EndGame(ByVal turn As Char)
        If (turn = "-") Then
            MessageBox.Show("No one Wins! It's a draw.", "Winner")
        Else
            MessageBox.Show(turn & " Wins!", "Winner")
            Label2.Text = (turn & " Won the last game.")
        End If

        For i As Integer = 1 To 9
            PB_chars(i) = "-"
        Next
        turn = "x"

        PictureBox1.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox2.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox3.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox4.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox5.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox6.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox7.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox8.Image = Tic_Tac_Toe.My.Resources.Resources.n
        PictureBox9.Image = Tic_Tac_Toe.My.Resources.Resources.n


    End Sub




    '----------------------------------
    'Handles pictureboxes 1-9

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If ((PB_chars(1) = "x") Or (PB_chars(1) = "o")) Then
            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(1) = "x"
                PictureBox1.Image = Tic_Tac_Toe.My.Resources.Resources.x

                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")
            ElseIf (turn = "o") Then
                PB_chars(1) = "o"
                PictureBox1.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        If ((PB_chars(2) = "x") Or (PB_chars(2) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(2) = "x"
                PictureBox2.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(2) = "o"
                PictureBox2.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        If ((PB_chars(3) = "x") Or (PB_chars(3) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(3) = "x"
                PictureBox3.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(3) = "o"
                PictureBox3.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        If ((PB_chars(4) = "x") Or (PB_chars(4) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(4) = "x"
                PictureBox4.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(4) = "o"
                PictureBox4.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        If ((PB_chars(5) = "x") Or (PB_chars(5) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(5) = "x"
                PictureBox5.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(5) = "o"
                PictureBox5.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        If ((PB_chars(6) = "x") Or (PB_chars(6) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(6) = "x"
                PictureBox6.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(6) = "o"
                PictureBox6.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        If ((PB_chars(7) = "x") Or (PB_chars(7) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(7) = "x"
                PictureBox7.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(7) = "o"
                PictureBox7.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        If ((PB_chars(8) = "x") Or (PB_chars(8) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(8) = "x"
                PictureBox8.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(8) = "o"
                PictureBox8.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        If ((PB_chars(9) = "x") Or (PB_chars(9) = "o")) Then

            MessageBox.Show("This Square has already been set!", "Error")
        Else
            If (turn = "x") Then
                PB_chars(9) = "x"
                PictureBox9.Image = Tic_Tac_Toe.My.Resources.Resources.x
                CheckWin(turn)
                turn = "o"
                Label1.Text = (turn & "'s turn.")

            ElseIf (turn = "o") Then
                PB_chars(9) = "o"
                PictureBox9.Image = Tic_Tac_Toe.My.Resources.Resources.o
                CheckWin(turn)
                turn = "x"
                Label1.Text = (turn & "'s turn.")

            End If
        End If
    End Sub
End Class
