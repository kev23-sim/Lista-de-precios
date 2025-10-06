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

                        // Mostrar los datos en pantalla
                        DisplayExcelData();
                    }
                }
            }
            catch (Exception ex)
            {
                await DisplayAlert("Error", $"Error al leer el archivo: {ex.Message}", "OK");
            }
        }//fin de load excel

        private void DisplayExcelData()
        {
            // Crear un CollectionView para mostrar los datos
            var collectionView = new CollectionView
            {
                ItemsSource = excelDataTable.DefaultView,
                ItemTemplate = new DataTemplate(() =>
                {
                    var grid = new Grid();

                    // Crear una etiqueta para cada celda en la fila
                    var label = new Label
                    {
                        Margin = new Thickness(5),
                        FontSize = 14,
                        LineBreakMode = LineBreakMode.WordWrap
                    };

                    // Vincular el texto de la etiqueta a los datos de la fila
                    label.SetBinding(Label.TextProperty, new Binding(".[0]")); // Muestra la primera columna

                    grid.Children.Add(label);
                    return grid;
                })
            };

            // Limpiar el grid y agregar el CollectionView

            MainGrid.Children.Clear();
            Grid.SetRow(collectionView, 3); // Fila 3 para mostrar los datos
            Grid.SetColumnSpan(collectionView, 2);
            MainGrid.Children.Add(collectionView);
        }



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


    }//fin de la clase
}//fin del namespace
