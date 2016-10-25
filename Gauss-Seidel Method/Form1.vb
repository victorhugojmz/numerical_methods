Public Class Jacobi
    Dim i As Integer
    Dim x(50) As Single
    Dim y(50) As Single
    Dim z(50) As Single
    Dim ex(50) As Single
    Dim ey(50) As Single
    Dim ez(50) As Single

    Dim a11 As Single
    Dim a12 As Single
    Dim a13 As Single
    Dim a21 As Single
    Dim a22 As Single
    Dim a23 As Single
    Dim a31 As Single
    Dim a32 As Single
    Dim a33 As Single
    Dim b1 As Single
    Dim b2 As Single
    Dim b3 As Single

    Dim ec As Single
    Dim c As Single
    Dim redon As Integer

    Private Sub salir_Click(sender As Object, e As EventArgs)
        End
    End Sub

    Private Sub calcular_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub salir_Click_1(sender As Object, e As EventArgs) Handles salir.Click
        End
    End Sub

    Private Sub calcular_Click_1(sender As Object, e As EventArgs) Handles calcular.Click
        a11 = ta11.Text
        a12 = ta12.Text
        a13 = ta13.Text
        a21 = ta21.Text
        a22 = ta22.Text
        a23 = ta23.Text
        a31 = ta31.Text
        a32 = ta32.Text
        a33 = ta33.Text
        b1 = tb1.Text
        b2 = tb2.Text
        b3 = tb3.Text

        ex(i) = 1
        ey(i) = 1
        ex(i) = 1
        c = tc.Text
        ec = 0.5 * 10 ^ (-c)
        redon = c + 2
        tr.Text = tr.Text & "i" & Chr(9) & "x" & Chr(9) & "y" & Chr(9) & "z" & Chr(9) & "ex" & Chr(9) &
            "ey" & Chr(9) & "ez" & vbCrLf

        tr.Text = tr.Text & i & Chr(9) & Math.Round(x(i), redon) & Chr(9) & Math.Round(y(i), redon) & Chr(9) &
            Math.Round(z(i), redon) & Chr(9) & "--" & Chr(9) & "--" & Chr(9) & "--" & vbCrLf

        Do While ex(i) > ec Or ey(i) > ec Or ez(i) > ec
            i = i + 1

            x(i) = (b1 - a12 * y(i - 1) - a13 * z(i - 1)) / a11
            y(i) = (b2 - a21 * x(i) - a23 * z(i - 1)) / a22
            z(i) = (b3 - a31 * x(i) - a32 * y(i)) / a33

            ''x(i) = (b1 - a12 * y(i - 1) - a13 * z(i - 1)) / a11
            ''y(i) = (b2 - a21 * x(i - 1) - a23 * z(i - 1)) / a22
            ''z(i) = (b3 - a31 * x(i - 1) - a32 * y(i - 1)) / a33

            ex(i) = Math.Abs((x(i) - x(i - 1)) / x(i))
            ey(i) = Math.Abs((y(i) - y(i - 1)) / y(i))
            ez(i) = Math.Abs((z(i) - z(i - 1)) / z(i))

            tr.Text = tr.Text & i & Chr(9) & Math.Round(x(i), redon) & Chr(9) & Math.Round(y(i), redon) &
                Chr(9) & Math.Round(z(i), redon) & Chr(9) & Math.Round(ex(i), redon) &
                Chr(9) & Math.Round(ey(i), redon) & Chr(9) & Math.Round(ez(i), redon) & vbCrLf
        Loop
        tr.Text = tr.Text & vbCrLf & "La solución es: X=" & Math.Round(x(i), redon) & ", Y= " &
            Math.Round(y(i), redon) & ", Z=" & Math.Round(z(i), redon)
    End Sub

    Private Sub Jacobi_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
