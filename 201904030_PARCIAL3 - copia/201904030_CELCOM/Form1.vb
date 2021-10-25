Public Class Form1
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click

        'BOTÓN:MOSTRAR, aqui se muestran los datos en los listbox's
        'mostrando datos en listbox's
        ListBox1.Items.Add(registro(fila))
        ListBox2.Items.Add(nombre(fila))
        ListBox3.Items.Add(sueldo(fila))
        ListBox4.Items.Add(ventas(fila))
        ListBox5.Items.Add(lugar(fila))
        ListBox6.Items.Add(bono(fila))
        ListBox7.Items.Add(viaticos(fila))
        ListBox8.Items.Add(sueldofinal(fila))



    End Sub

    Private Sub OperarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OperarToolStripMenuItem.Click

        'BOTÓN: OPERAR, aqui se guardan los valores en los vectores y se realizan los cálculos correspondientes
        'Asignando valores a vectores

        While fila < 5
            registro(fila) = Val(TextBox1.Text)
            nombre(fila) = TextBox2.Text
            sueldo(fila) = Val(TextBox3.Text)
            ventas(fila) = Val(TextBox4.Text)
            lugar(fila) = ComboBox1.SelectedItem

            'Cálculo del bono:
            Select Case ventas(fila)
                Case 2000 To 10000
                    bono(fila) = Val(ventas(fila)) * 3 / 100
                Case 10001 To 20000
                    bono(fila) = Val(ventas(fila)) * 5 / 100
                Case > 20001
                    bono(fila) = Val(ventas(fila)) * 6 / 100
            End Select

            'Cálculo de viáticos:
            Select Case ComboBox1.SelectedItem
                Case "Norte"
                    viaticos(fila) = Val(ventas(fila)) * 5 / 100
                Case "Sur"
                    viaticos(fila) = Val(ventas(fila)) * 5.5 / 100
                Case "Occidente"
                    viaticos(fila) = Val(ventas(fila)) * 6 / 100
                Case "Oriente"
                    viaticos(fila) = Val(ventas(fila)) * 6.5 / 100
            End Select

            'Cálculo del sueldo final:
            sueldofinal(fila) = ventas(fila) + bono(fila) + viaticos(fila)

            'SI INGRESA UN REGISTRO REPETIDO:
            For fila As Integer = 0 To 5
                If (registro(fila) <> Nothing) Then
                    If (registro(fila) = TextBox1.Text) Then
                        MsgBox("Número de registro repetido")
                        Exit Sub
                    End If
                End If
            Next

            If fila = 6 Then
                MsgBox("Vectores llenos :(")
            End If
        End While


    End Sub

    Private Sub ConsultarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultarToolStripMenuItem.Click

        'BOTÓN: CONSULTAR
        'Realizar consulta con el registro 
        If TextBox1.Text = "" Then
            MsgBox("No ingresó un número de registro")
            Exit Sub
        End If

        For fila As Integer = 0 To 5
            If (registro(fila) <> Nothing) Then
                If (registro(fila) = TextBox1.Text) Then
                    registro(fila) = Val(TextBox1.Text)
                    nombre(fila) = TextBox2.Text
                    sueldo(fila) = Val(TextBox3.Text)
                    ventas(fila) = Val(TextBox4.Text)
                    lugar(fila) = ComboBox1.SelectedItem

                    Exit Sub
                End If
            End If
        Next


    End Sub

    Private Sub EstadísticasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EstadísticasToolStripMenuItem.Click

        'BOTÓN: ESTADÍSTICAS
        'estadísticas


        Dim est1 As Integer
        Dim est2 As Integer
        Dim est3 As Integer

        est1 = 0
        est2 = 0
        est3 = 0

        While fila <= 5

            'cuantos vendedores son de la región norte
            If lugar(fila) = "norte" Then
                est1 = est1 + 1
            End If

            'monto (Q) de los empleados de la región oriente
            If lugar(fila) = "oriente" Then
                est2 = est2 + Val(sueldofinal(fila))
            End If

            'Total viáticos 
            est3 = est3 + (viaticos(fila))

        End While

        'mostrando valores de estadísticas en las textbox's
        TextBox5.Text = Str(est1)
        TextBox6.Text = Str(est2)
        TextBox7.Text = Str(est3)



    End Sub

    Private Sub LimpiarEntradasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarEntradasToolStripMenuItem.Click

        'BOTÓN: LIMPIAR ENTRADAS
        'dejar los datos limpios para un nuevo ingreso 

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.SelectedIndex = -1



    End Sub

    Private Sub LimpiarVectoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarVectoresToolStripMenuItem.Click

        'BOTÓN: LIMPIAR VECTORES
        'borra datos almacenados y del historial donde se muestran
        fila = 0
        registro(fila) = Nothing
        nombre(fila) = Nothing
        sueldo(fila) = 0
        ventas(fila) = 0
        lugar(fila) = Nothing
        bono(fila) = 0
        viaticos(fila) = 0
        sueldofinal(fila) = 0

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        ListBox8.Items.Clear()




    End Sub

    Private Sub LimpiarEstadísticasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarEstadísticasToolStripMenuItem.Click

        'BOTÓN: LIMPIAR ESTADÍSTICAS
        'limpiar estadisticas

        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()

    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click

        'BOTÓN:SALIR
        'salir con pregunta
        Form2.Show()


    End Sub
End Class
