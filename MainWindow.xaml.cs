using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace TextoAExcelWPF;

public partial class MainWindow
{
    private string? _filePath;

    public MainWindow()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Controla el evento de hacer clic en el control
    /// y abre un cuadro de diálogo para seleccionar un archivo
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Archivos de texto (*.txt)|*.txt"
        };
        if (openFileDialog.ShowDialog() != true) return;
        _filePath = openFileDialog.FileName;
        // Actualiza el texto para mostrar el nombre del archivo seleccionado
        TextBlockMensaje.Text = Path.GetFileName(_filePath);
    }

    /// <summary>
    /// Comportamiento de arrastrar y soltar para el mouse
    /// cuando se arrastra un archivo sobre el control
    /// para dar feedback visual al usuario
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Border_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }


    /// <summary>
    /// Controla el evento de soltar un archivo en el control
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Border_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        var files = e.Data.GetData(DataFormats.FileDrop) as string?[];
        _filePath = files?[0];
        // Actualiza el texto para mostrar el nombre del archivo seleccionado
        TextBlockMensaje.Text = Path.GetFileName(_filePath);
    }


    /// <summary>
    /// Procesa el archivo de texto seleccionado y genera un archivo Excel
    /// pidiendo al usuario la ubicación donde se guardará y el nombre del archivo
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Procesar_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_filePath) || !File.Exists(_filePath))
        {
            MessageBox.Show("Por favor, selecciona un archivo válido.", "Archivo no encontrado",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            // Leer el archivo de texto
            var lines = File.ReadAllLines(_filePath);
            var data = (from line in lines
                    select line.Split(' ')
                    into parts
                    where parts.Length >= 2
                    let nombre = parts[0]
                    let apellido = parts[1]
                    let nombreCompleto = $"{nombre} {apellido}"
                    select new Persona { Nombre = nombre, Apellido = apellido, NombreCompleto = nombreCompleto })
                .ToList();

            // Seleccionar ubicación para guardar el archivo Excel
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
            };
            if (saveFileDialog.ShowDialog() != true) return;
            ExportarAExcel(data, saveFileDialog.FileName);
            MessageBox.Show("Archivo Excel generado exitosamente.", "Proceso completado", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ocurrió un error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// Método para exportar los datos a un archivo Excel
    /// </summary>
    /// <param name="data"></param>
    /// <param name="excelFilePath"></param>
    private static void ExportarAExcel(List<Persona> data, string excelFilePath)
    {
        // Se requiere instalar ClosedXML a través de NuGet
        // Install-Package ClosedXML

        using var workbook = new ClosedXML.Excel.XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Datos");
        worksheet.Cell(1, 1).Value = "Nombre";
        worksheet.Cell(1, 2).Value = "Apellido";
        worksheet.Cell(1, 3).Value = "Nombre Completo";

        for (var i = 0; i < data.Count; i++)
        {
            worksheet.Cell(i + 2, 1).Value = data[i].Nombre;
            worksheet.Cell(i + 2, 2).Value = data[i].Apellido;
            worksheet.Cell(i + 2, 3).Value = data[i].NombreCompleto;
        }

        workbook.SaveAs(excelFilePath);
    }
}