Public Class Form1
    Dim x(50), y(50) As Single
    Dim n, i, id, idaux, valorAInterpolar, redon As Integer

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
    Private Sub bts_Click(sender As Object, e As EventArgs) Handles bts.Click
        Try
            redon = tn.Text + 2
            Try
                n = tbn.Text
                If n < 2 Then
                    MessageBox.Show("Por lo menos deben ser 2 parejas de números ", "Informe", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Information)
                    tbn.Focus()
                Else
                    Try
                        valorAInterpolar = tbx.Text

                        tbr.Text = "i" & Chr(9) & "X" & Chr(9) & "Y" & vbCrLf

                        tn.Enabled = False
                        tbn.Enabled = False
                        tbx.Enabled = False
                        bts.Enabled = False

                        tbxi.Enabled = True
                        tbyi.Enabled = True
                        btin.Enabled = True
                        tbxi.Focus()
                    Catch ex As Exception

                        MessageBox.Show("El valor a interpolar no es válido", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        tbx.Focus()
                    End Try
                End If
            Catch ex As Exception
                MessageBox.Show("El número de parejas no es válido", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tbn.Focus()
            End Try
        Catch ex As Exception
            MessageBox.Show("El número de cifras significativas no es válido", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tn.Focus()
        End Try
    End Sub

    Private Sub btin_Click(sender As Object, e As EventArgs) Handles btin.Click

        If i < n Then

            Try
                x(i) = tbxi.Text
                Try
                    y(i) = tbyi.Text

                    tbr.Text = tbr.Text & i & Chr(9) & x(i) & Chr(9) & y(i) & vbCrLf
                    i = i + 1

                    If i = n Then

                        tbxi.Enabled = False
                        tbyi.Enabled = False
                        btin.Enabled = False

                        btc.Enabled = True
                        btc.Focus()
                    End If

                    tbyi.Text = ""
                    tbxi.Text = ""
                    tbxi.Focus()
                Catch ex As Exception

                    MessageBox.Show("El valor de y(i) no es válido", "¡Atención!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    tbyi.Focus()
                End Try
            Catch ex As Exception
                MessageBox.Show("El valor de x(i) no es válido", "¡Atención!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                tbxi.Focus()
            End Try
        End If
    End Sub

    Private Sub btc_Click(sender As Object, e As EventArgs) Handles btc.Click
        valorAInterpolar = tbx.Text
        i = 0

        Do While x(i) <= tbx.Text
            id = i

            idaux = id
            i = i + 1
        Loop
        i = i - 1
        If cbm.SelectedIndex = 0 Then

            tbr2.Text = tbr2.Text & "Indice = " & i & vbCrLf
            tbr2.Text = tbr2.Text & "X(0) = " & tbx.Text & vbCrLf
            tbr2.Text = tbr2.Text & "X(1) =" & x(i) & vbCrLf
            tbr2.Text = tbr2.Text & "X(1+1) =" & x(i + 1) & vbCrLf
            tbr2.Text = tbr2.Text & "Y(1) =" & y(i) & vbCrLf
            tbr2.Text = tbr2.Text & "Y(1+1) =" & y(i + 1) & vbCrLf
            tbs.Text = "y(" & tbx.Text & ") = " & Math.Round(((((tbx.Text - x(i + 1)) / (x(i) - x(i + 1))) * y(i)) + (((tbx.Text - x(i)) / (x(i + 1) - x(i))) * y(i + 1))), redon)

        Else

            tbr2.Text = tbr2.Text & "indice = " & i & vbCrLf
            tbr2.Text = tbr2.Text & "X(0) = " & tbx.Text & vbCrLf
            tbr2.Text = tbr2.Text & "X(1) =" & x(i) & vbCrLf
            tbr2.Text = tbr2.Text & "X(1+1) =" & x(i + 1) & vbCrLf
            tbr2.Text = tbr2.Text & "X(1+2) =" & x(i + 2) & vbCrLf
            tbr2.Text = tbr2.Text & "Y(1)=" & y(i) & vbCrLf
            tbr2.Text = tbr2.Text & "Y(1+1)=" & y(i + 1) & vbCrLf
            tbr2.Text = tbr2.Text & "Y(1+2)=" & y(i + 2) & vbCrLf
            tbs.Text = "y(" & tbx.Text & ") = " & Math.Round((((((tbx.Text - x(i + 1)) * (tbx.Text - x(i + 2))) / ((x(i) - x(i + 1)) * (x(i) - x(i + 2)))) * y(i)) + ((((tbx.Text - x(i)) * (tbx.Text - x(i + 2))) / ((x(i + 1) - x(i)) * (x(i + 1) - x(i + 2)))) * y(i + 1)) + ((((tbx.Text - x(i)) * (tbx.Text - x(i + 1))) / ((x(i + 2) - x(i)) * (x(i + 2) - x(i + 1)))) * y(i + 2))), redon)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles bta.Click
        If cbm.SelectedIndex = 0 Then

            MessageBox.Show("Interpolación de primer grado", "¡Atención!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            bta.Enabled = False
            cbm.Enabled = False
            tn.Enabled = True
            tbn.Enabled = True
            tbx.Enabled = True
            bts.Enabled = True
            tn.Focus()
        ElseIf cbm.SelectedIndex = 1 Then

            MessageBox.Show("interpolación de segundo grado", "¡Atención!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            bta.Enabled = False
            cbm.Enabled = False
            tn.Enabled = True
            tbn.Enabled = True
            tbx.Enabled = True
            bts.Enabled = True
            tn.Focus()
        Else
            MessageBox.Show("No se ha seleccionado el tipo de interpolación", "¡Atención!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            cbm.Focus()
        End If
    End Sub
End Class
