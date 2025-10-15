using ExcelDataReader;
using System.Data;
using Microsoft.Maui.ApplicationModel;
using ZXing.Net.Maui.Controls;

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
            Busqueda.Focused += OnBusquedaFocused;
            audioRecorder = new AudioRecorderService();
        }//fin del constructor


        private void OnBusquedaFocused(object sender, FocusEventArgs e)
        {
            if (Busqueda.Text == "Ingresa el nombre del producto")
            {
                Busqueda.Text = "";
                Busqueda.Opacity = 1.0; // Cambiar opacidad a normal
            }
        }//fin del OnBusquedaFocused
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
                        await DisplayAlert("Aviso", $"Excel cargado correctamente:","OK");

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


        private void BuscarProducto_clicked(object sender, EventArgs e)
        {
            string buscar = Busqueda.Text;
            BuscarProducto(buscar);
        }//fin del EscanearCodigoBarras_clicked
        private void BuscarProducto(string textoBusqueda)
        {

            if (excelDataTable == null || excelDataTable.Rows.Count == 0)
            {
                DisplayAlert("Error", "No hay datos cargados del Excel", "OK");
                return;
            }

            try
            {
                // Usar índice numérico en lugar de nombre (columna 1 = "DETALLE")
                var resultados = excelDataTable.AsEnumerable()
                    .Where(row =>
                    {
                        var valor = row.Field<string>(1); // Índice 1 para la segunda columna
                        return valor?.IndexOf(textoBusqueda, StringComparison.OrdinalIgnoreCase) >= 0;
                    })
                    .ToList();

                if (resultados.Count == 0)
                {
                    DisplayAlert("Búsqueda", "No se encontraron productos con ese nombre", "OK");
                    return;
                }

                // Mostrar los resultados filtrados
                MostrarResultadosBusqueda(resultados);
            }
            catch (Exception ex)
            {
                DisplayAlert("Error", $"Error al buscar: {ex.Message}", "OK");
            }
        }//fin del metodo BuscarProducto


        private void MostrarResultadosBusqueda(List<DataRow> resultados)
        {
            // Configurar el CollectionView que ya existe en el XAML
            ResultadosCollectionView.ItemsSource = resultados;
            ResultadosCollectionView.ItemTemplate = new DataTemplate(() =>
            {
                var grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Star });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                // Referencia
                var lblReferencia = new Label
                {
                    Margin = new Thickness(5),
                    FontSize = 12,
                    FontAttributes = FontAttributes.Bold,
                    WidthRequest = 80
                };
                lblReferencia.SetBinding(Label.TextProperty, new Binding("[0]")); // Columna REFERENCIA

                // Detalle
                var lblDetalle = new Label
                {
                    Margin = new Thickness(5),
                    FontSize = 14,
                    LineBreakMode = LineBreakMode.WordWrap
                };
                lblDetalle.SetBinding(Label.TextProperty, new Binding("[1]")); // Columna DETALLE

                // Valor
                var lblValor = new Label
                {
                    Margin = new Thickness(5),
                    FontSize = 14,
                    FontAttributes = FontAttributes.Bold,
                    TextColor = Colors.Green,
                    WidthRequest = 80,
                    HorizontalTextAlignment = TextAlignment.End
                };
                lblValor.SetBinding(Label.TextProperty, new Binding("[2]")); // Columna VALOR POS

                // Agregar controles al grid
                Grid.SetColumn(lblReferencia, 0);
                Grid.SetColumn(lblDetalle, 1);
                Grid.SetColumn(lblValor, 2);

                grid.Children.Add(lblReferencia);
                grid.Children.Add(lblDetalle);
                grid.Children.Add(lblValor);

                return grid;
            });

            // Hacer visible el CollectionView
            ResultadosCollectionView.IsVisible = true;

        }//fin del MostrarResultadosBusqueda


    }//fin de la clase
}//fin del namespace
