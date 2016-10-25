Public Class Form1
    Dim x(50) As Single
    Dim y(50) As Single
    Dim z(50) As Single
    Dim redon As Integer
    Dim ex(50) As Single
    Dim ey(50) As Single
    Dim ez(50) As Single
    Dim ec As Single
    Dim i As Integer
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Calcular_Click(sender As Object, e As EventArgs) Handles Calcular.Click
        tabla.Text = " "
        ex(i) = 1
        ey(i) = 1
        ex(i) = 1
        ''x(i) = 1
        ''y(i) = 2
        ''z(i) = 2
        ec = 0.5 * 10 ^ (-c.Text)
        redon = c.Text + 2
        ''redon = 3
        tabla.Text = tabla.Text & "i" & Chr(9) & "X" & Chr(9) & "Y" & Chr(9) & "Z" & Chr(9) & Chr(9) & "EX" & Chr(9) & "EY" & Chr(9) & "EZ" & vbCrLf & vbCrLf

        tabla.Text = tabla.Text & i & Chr(9) & Math.Round(x(i), redon) & Chr(9) & Math.Round(y(i), redon) & Chr(9) & Math.Round(z(i), redon) &
                                        Chr(9) & Chr(9) & "--" & Chr(9) & "--" & Chr(9) & "--" & vbCrLf

        Do While ex(i) > ec Or ey(i) > ec Or ez(i) > ec

            x(i + 1) = (b1.Text - a12.Text * y(i) - a13.Text * z(i)) / a11.Text
            y(i + 1) = (b2.Text - a21.Text * x(i) - a23.Text * z(i)) / a22.Text
            z(i + 1) = (b3.Text - a31.Text * x(i) - a32.Text * y(i)) / a33.Text

            ex(i + 1) = Math.Abs((x(i + 1) - x(i)) / x(i + 1))
            ey(i + 1) = Math.Abs((y(i + 1) - y(i)) / y(i + 1))
            ez(i + 1) = Math.Abs((z(i + 1) - z(i)) / z(i + 1))

            tabla.Text = tabla.Text & i + 1 & Chr(9) & Math.Round(x(i + 1), redon) & Chr(9) & Math.Round(y(i + 1), redon) & Chr(9) & Math.Round(z(i + 1), redon) &
                                            Chr(9) & Chr(9) & Math.Round(ex(i + 1), redon) & Chr(9) & Math.Round(ey(i + 1), redon) & Chr(9) & Math.Round(ez(i + 1), redon) & vbCrLf
            i = i + 1
        Loop
        tabla.Text = tabla.Text & vbCrLf & "La solución es: X=    " & Math.Round(x(i), redon) & ", Y= " &
           Math.Round(y(i), redon) & ", Z=" & Math.Round(z(i), redon)
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub
End Class
