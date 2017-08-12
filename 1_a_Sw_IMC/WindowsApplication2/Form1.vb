Public Class Form1

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim peso, altura As Double
        Dim imc As Double
        peso = Val(TextBox1.Text)
        altura = Val(TextBox2.Text)
        imc = peso / (altura * altura)
        If IsNumeric(TextBox1.Text) And Val(TextBox1.Text) > 0 Then
            peso = Val(TextBox1.Text)
        Else : MsgBox("Error, ingrese un número válido")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox2.Text) And Val(TextBox2.Text) > 0 Then
            altura = Val(TextBox2.Text)
        Else : MsgBox("Error, ingrese un número válido")
            TextBox1.Focus()
            Exit Sub
        End If
        TextBox3.Text = Str(imc)
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) < 18.5 Then
            MsgBox("Clasificación: Bajo peso")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) = 18.5 Then
            MsgBox("Clasificación: Peso Promedio")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) < 25 Then
            MsgBox("Clasificación: Peso Promedio")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) = 25 Then
            MsgBox("Clasificación: Aumentado")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) < 30 Then
            MsgBox("Clasificación: Aumentado")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) = 30 Then
            MsgBox("Clasificación: Moderado")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) < 35 Then
            MsgBox("Clasificación: Moderado")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) = 35 Then
            MsgBox("Clasificación: Severo")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) < 40 Then
            MsgBox("Clasificación: Severo")
            TextBox1.Focus()
            Exit Sub
        End If
        If IsNumeric(TextBox3.Text) And Val(TextBox3.Text) >= 40 Then
            MsgBox("Clasificación: Muy Severo")
            TextBox1.Focus()
            Exit Sub
        End If
    End Sub
End Class
