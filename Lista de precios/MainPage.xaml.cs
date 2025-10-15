using ExcelDataReader;
using System.Data;
using Microsoft.Maui.ApplicationModel;

namespace Lista_de_precios
{
    public partial class MainPage : ContentPage
    {
        private DataTable excelDataTable;
        private AudioRecorderService audioRecorder;
        private bool isRecording = false;

        public MainPage()
        {
            InitializeComponent();
            audioRecorder = new AudioRecorderService();
        }//fin del constructor

        private async void CargarExcelClicked(object sender, EventArgs e)
        {
            try
            {
                var result = await FilePicker.PickAsync(new PickOptions
                {
                    PickerTitle = "Seleccionar archivo Excel",
                    FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                    {
                        { DevicePlatform.WinUI, new[] { ".xlsx", ".xls" } },
                        { DevicePlatform.Android, new[] { "application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
                        { DevicePlatform.iOS, new[] { "com.microsoft.excel.xls", "org.openxmlformats.spreadsheetml.sheet" } }
                    })
                });

                if (result != null)
                {
                    await LoadExcelFile(result.FullPath);
                }
            }
            catch (Exception ex)
            {
                await DisplayAlert("Error", $"No se pudo cargar el archivo: {ex.Message}", "OK");
            }
        
        }//fin cargar excel clicked

        private async Task LoadExcelFile(string filePath)
        {
            try
            {
                // Configurar el codificador para ExcelDataReader
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Leer el archivo Excel como DataSet
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true // Usar primera fila como encabezados
                            }
                        });

                        // Tomar la primera hoja
                        excelDataTable = result.Tables[0];

                    }
                }
            }
            catch (Exception ex)
            {
                await DisplayAlert("Error", $"Error al leer el archivo: {ex.Message}", "OK");
            }
        }//fin de load excel




        private async void OnRecordButtonClicked(object sender, EventArgs e)
        {
            if (!isRecording)
            {
                // Iniciar grabación
                var success = await audioRecorder.StartRecordingAsync();
                if (success)
                {
                    isRecording = true;
                    BusquedaVoz.Text = "🛑 Detener Grabación";
                    await DisplayAlert("Grabación", "Grabación iniciada", "OK");
                }
                else
                {
                    await DisplayAlert("Error", "No se pudo iniciar la grabación. Verifica los permisos.", "OK");
                }
            }
            else
            {
                // Detener grabación
                var filePath = await audioRecorder.StopRecordingAsync();
                isRecording = false;
                BusquedaVoz.Text = "🎤 Grabar Audio";

                if (!string.IsNullOrEmpty(filePath))
                {
                    await DisplayAlert("Grabación", $"Audio guardado en: {filePath}", "OK");
                }
                else
                {
                    await DisplayAlert("Error", "No se pudo guardar el audio", "OK");
                }
            }
        }// fin del metodo OnRecordButtonClicked

        /*
        private void BusquedaClicked (object sender, EventArgs e)
        {
            BuscarEnExcel(Busqueda.Text, true, "2");
            MostrarResultadosBusqueda(excelDataTable);
        }


        public DataTable BuscarEnExcel(string textoBusqueda, bool buscarEnTodasLasColumnas = true, string columnaEspecifica = "")
        {
            if (excelDataTable == null || string.IsNullOrWhiteSpace(textoBusqueda))
            {
                return excelDataTable?.Clone() ?? new DataTable();
            }

            try
            {
                // Crear filtro para DataTable
                string filtro = CrearFiltroBusqueda(textoBusqueda, buscarEnTodasLasColumnas, columnaEspecifica);

                // Aplicar filtro
                DataRow[] filasFiltradas = excelDataTable.Select(filtro);

                // Convertir resultado a DataTable
                return FilasToDataTable(filasFiltradas);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en búsqueda: {ex.Message}");
                return excelDataTable?.Clone() ?? new DataTable();
            }
        }//BuscarEnExcel

        private string CrearFiltroBusqueda(string texto, bool todasLasColumnas, string columnaEspecifica)
        {
            if (todasLasColumnas)
            {
                // Buscar en todas las columnas
                var filtros = new List<string>();
                foreach (DataColumn columna in excelDataTable.Columns)
                {
                    if (columna.DataType == typeof(string))
                    {
                        filtros.Add($"[{columna.ColumnName}] LIKE '%{EscapeLikeValue(texto)}%'");
                    }
                }
                return string.Join(" OR ", filtros);
            }
            else if (!string.IsNullOrEmpty(columnaEspecifica))
            {
                // Buscar en columna específica
                return $"[{columnaEspecifica}] LIKE '%{EscapeLikeValue(texto)}%'";
            }

            return "";
        }//CrearFiltroBusqueda

        private string EscapeLikeValue(string value)
        {
            // Escapar caracteres especiales para LIKE
            return value.Replace("[", "[[]").Replace("%", "[%]").Replace("_", "[_]");
        }//EscapeLikeValue

        private DataTable FilasToDataTable(DataRow[] filas)
        {
            if (filas == null || filas.Length == 0)
                return excelDataTable?.Clone() ?? new DataTable();

            DataTable tablaResultado = excelDataTable.Clone();
            foreach (DataRow fila in filas)
            {
                tablaResultado.ImportRow(fila);
            }
            return tablaResultado;
        }//FilasToDataTable

        private async void MostrarResultadosBusqueda(DataTable resultados)
        {
            if (resultados.Rows.Count == 0)
            {
                await DisplayAlert("Búsqueda", "No se encontraron resultados", "OK");
                return;
            }

            // Aquí puedes actualizar tu DataGridView, ListView, o cualquier control
            // con los resultados filtrados

            await DisplayAlert("Búsqueda", $"Se encontraron {resultados.Rows.Count} resultados", "OK");
        }//MostrarResultadosBusqueda*/

    }//fin de la clase
}//fin del namespace
