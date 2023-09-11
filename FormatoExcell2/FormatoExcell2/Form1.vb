Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1

    Inherits Form
    Private dataTable As System.Data.DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim connectionString As String = "Data Source=(localdb)\AIPR;Initial Catalog=MBDD1;Integrated Security=true;"

        Using connection As New SqlConnection(connectionString)
            connection.Open()
            ' Ejecuta tu consulta SQL y llena el DataTable.
            Dim query As String = "SELECT * FROM Clientes" ' Utiliza el nombre de tu tabla "Clientes".
            Dim adapter As New SqlDataAdapter(query, connection)
            dataTable = New System.Data.DataTable()
            adapter.Fill(dataTable)
        End Using

        ' Crear una instancia de Excel
        Dim excelApp As New Excel.Application()
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim worksheet As Excel.Worksheet = DirectCast(workbook.Worksheets(1), Excel.Worksheet)

        ' Supongamos que tienes un DataTable llamado "dataTable" con tus datos.
        Dim fila As Integer = 1

        For Each row As DataRow In dataTable.Rows
            worksheet.Cells(fila, 1) = row("IDCliente")
            worksheet.Cells(fila, 2) = row("Nombre")
            worksheet.Cells(fila, 3) = row("Apellido")
            worksheet.Cells(fila, 4) = row("Monto")
            ' Continúa colocando los datos en las celdas correspondientes.
            fila += 1
        Next

        ' Cálculo de sumatorias
        worksheet.Cells(fila, 1) = "Total:"
        worksheet.Cells(fila, 4).Formula = "SUM(D1:D" & (fila - 1) & ")"

        ' Mostrar números de línea
        Dim contador As Integer = 1

        For Each row As DataRow In dataTable.Rows
            worksheet.Cells(fila, 1) = contador
            worksheet.Cells(fila, 2) = row("Nombre")
            worksheet.Cells(fila, 3) = row("Apellido")
            ' Continúa colocando los datos en las celdas correspondientes.
            fila += 1
            contador += 1
        Next

        ' Guardar el archivo Excel
        workbook.SaveAs("C:\Users\venta\source\repos\FormatoExcell2\ruta_del_archivo.xlsx")

        ' Abrir el archivo en modo solo lectura
        workbook.ReadOnlyRecommended = True
        workbook.Close(False)
        excelApp.Quit()
        DataGridView1.DataSource = dataTable
    End Sub
End Class
