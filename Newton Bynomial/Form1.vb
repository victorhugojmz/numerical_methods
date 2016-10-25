Public Class Form1
    Dim x(50), y(50) As Single
    Dim decr, dec2, factorial, sumatoria As Single

    Dim n, i, id, idaux As Integer

    Private Sub tbr_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub tbvalores_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ln_Click(sender As Object, e As EventArgs) Handles ln.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub tbX_TextChanged(sender As Object, e As EventArgs) Handles tbX.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Reinicio_Click(sender As Object, e As EventArgs)

    End Sub

    Dim redon As Integer

    Private Sub bts_Click(sender As Object, e As EventArgs) Handles btS.Click
        btS.Enabled = False
        tbXi.Enabled = True
        tbYi.Enabled = True
        btV.Enabled = True
        tbresultado.Text = "i" & Chr(9) & "X" & Chr(9) & "Y" & vbCrLf
        n = tbN.Text
        i = 0
    End Sub

    Private Sub btin_Click(sender As Object, e As EventArgs) Handles btV.Click
        If i < n Then
            x(i) = tbXi.Text
            y(i) = tbYi.Text
            tbXi.Focus()
            tbresultado.Text = tbresultado.Text & i & Chr(9) & x(i) & Chr(9) & y(i) & vbCrLf
            i = i + 1
            If i = n Then
                btV.Enabled = False
                btC.Enabled = True
            End If
            tbYi.Text = ""
            tbXi.Text = ""
        End If
    End Sub

    Private Sub btc_Click(sender As Object, e As EventArgs) Handles btC.Click
        Dim s, m, j, auxy As Single
        Dim dif(50), combinaciones(50), combsm(50), ant, sup(50), inf(50) As Single
        Dim col(50, 50) As Single
        Dim diferencias(50) As Single



        redon = tcs.Text + 2
        i = 0
        Do While x(i) <= tbX.Text
            id = i
            idaux = id
            i = i + 1
        Loop

        tbresultado.Text = tbresultado.Text & "Parejas = " & n & vbCrLf
        m = n - (id + 1)
        tbresultado.Text = tbresultado.Text & "m = " & m & vbCrLf
        tbresultado.Text = tbresultado.Text & "Indice = " & id & vbCrLf
        tbresultado.Text = tbresultado.Text & "Valor a Interpolar= " & tbX.Text & vbCrLf
        tbresultado.Text = tbresultado.Text & "x(0) = " & x(id) & Chr(9) & "y(i) = " & y(id) & vbCrLf
        tbresultado.Text = tbresultado.Text & "x(0+1) = " & x(id + 1) & Chr(9) & "y(i+1) = " & y(id + 1) & vbCrLf
        s = (tbX.Text - x(id)) / (x(id + 1) - x(id))
        tbresultado.Text = tbresultado.Text & "S = " & Math.Round(s, redon) & vbCrLf
        i = 0
        dif(i) = y(id)
        auxy = dif(i)
        tbresultado.Text = tbresultado.Text & "==========================Valores diagonal==============================" & vbCrLf
        tbresultado.Text = tbresultado.Text & dif(i) & vbCrLf
        j = 0
        'Diagonal
        Do While j < m
            Do While i <= m
                dif(i) = y(id + 1) - y(id)
                y(id) = dif(i)
                If i = 0 Then
                    diferencias(j) = Math.Round(y(id), redon)
                    tbresultado.Text = tbresultado.Text & diferencias(j) & vbCrLf
                End If
                id = id + 1
                i = i + 1
            Loop
            i = 0
            id = idaux
            j = j + 1
        Loop
        'factorial
        inf(0) = 1
        inf(1) = s
        i = 2
        factorial = 1
        decr = m
        dec2 = m
        j = m
        Do While j <> 0
            Do While dec2 <> 0
                factorial = factorial * dec2
                inf(j - 1) = factorial
                dec2 = dec2 - 1
            Loop
            j = j - 1
            dec2 = j
            factorial = 1
        Loop

        j = 0
        ant = 1
        Do While j < m
            combsm(j) = ant * (s - j)
            'tbdivision.Text = tbdivision.Text & combsm(j) & "/" & inf(j) & vbCrLf
            j = j + 1
            ant = combsm(j - 1)
        Loop

        j = 0
        i = m
        tbresultado.Text = tbresultado.Text & "=============== Valores de S/I ==========" & vbCrLf
        Do While j < i
            combinaciones(j) = combsm(j) / inf(j)

            tbresultado.Text = tbresultado.Text & Math.Round(combinaciones(j), redon) & vbCrLf
            j = j + 1
        Loop
        j = 0
        i = m
        tbresultado.Text = tbresultado.Text & "y(" & tbX.Text & ") = (" & auxy & ") * (1) "
        sumatoria = auxy
        Do While j < i
            tbresultado.Text = tbresultado.Text & "+ (" & diferencias(j) & ") * (" & combinaciones(j) & ") "
            sumatoria = sumatoria + (Math.Round(diferencias(j), redon) * Math.Round(combinaciones(j), redon))
            j = j + 1
        Loop
        tbresultado.Text = tbresultado.Text & "= " & (Math.Round(sumatoria, redon))
    End Sub



End Class
