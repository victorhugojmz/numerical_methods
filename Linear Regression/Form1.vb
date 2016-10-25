Public Class Form1
    'chart command 
    Dim x(50), y(50) As Single
    Dim sx, sy, sxy, sxc, xm, ym As Single
    Dim i = 0, n, redon As Integer
    Dim a, b As Double

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub agregar_Click(sender As Object, e As EventArgs) Handles agregar.Click
        If i < n Then

            x(i) = tx.Text
            y(i) = ty.Text
            tr.Text = tr.Text & i & Chr(9) & x(i) & Chr(9) & y(i) & vbCrLf
            table.Text = table.Text & x(i) * x(i) & Chr(9) & y(i) * y(i) & Chr(9) & x(i) * y(i) & vbCrLf
            i += 1
            If i = n Then
                tx.Enabled = False
                ty.Enabled = False
                calcular.Enabled = True
                calcular.Focus()
            End If
            tx.Text = ""
            ty.Text = ""
            tx.Focus()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles asignar.Click
        n = numparejas.Text
        tr.Text = "i" & Chr(9) & "X" & Chr(9) & "Y" & vbCrLf
        table.Text = "X2" & Chr(9) & "Y2" & Chr(9) & "XY" & vbCrLf
    End Sub

    Private Sub Chart1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Chart1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles table.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles ty.TextChanged

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        End
    End Sub
    Private Sub calcular_Click(sender As Object, e As EventArgs) Handles calcular.Click
        i = 0
        Do While i < n
            sx += x(i)
            sy += y(i)
            sxy += x(i) * y(i)
            sxc += Math.Pow(x(i), 2)
            i = i + 1
        Loop
        xm = sx / n
        ym = sy / n
        b = (n * sxy - (sx * sy)) / (n * sxc - (Math.Pow(sx, 2)))
        br.Text = b
        a = ((sy - b * sx) / n)
        ar.Text = a
    End Sub
    Private Sub Label9_Click(sender As Object, e As EventArgs)

    End Sub
End Class
