using Microsoft.Win32;

using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace IMLoader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ExcelFileModel? _masterFile;
        private ObservableCollection<ExcelFileModel> _additionalFiles = new ObservableCollection<ExcelFileModel>();

        public MainWindow()
        {
            InitializeComponent();
            ListFiles.ItemsSource = _additionalFiles;
            ListFiles.SelectionChanged += ListFiles_SelectionChanged;
            ListFiles.ItemTemplate = CreateFileTemplate();
        }

        private void BtnSelectMaster_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*" };
            if (dlg.ShowDialog() == true)
            {
                var sheets = ExcelHelper.GetSheetNames(dlg.FileName);
                var defaultSheet = sheets.FirstOrDefault(s => s.Equals("Inspection_Task", StringComparison.OrdinalIgnoreCase)) ?? sheets.FirstOrDefault();
                _masterFile = new ExcelFileModel
                {
                    FilePath = dlg.FileName,
                    Sheets = sheets,
                    SelectedSheet = defaultSheet,
                    IsMaster = true
                };
                TxtMasterFilePath.Text = dlg.FileName;
                CmbMasterSheet.ItemsSource = sheets;
                CmbMasterSheet.SelectedItem = defaultSheet;
            }
        }

        private void CmbMasterSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_masterFile != null && CmbMasterSheet.SelectedItem is string sheet)
            {
                _masterFile.SelectedSheet = sheet;
            }
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*", Multiselect = true };
            if (dlg.ShowDialog() == true)
            {
                foreach (var file in dlg.FileNames)
                {
                    if (_additionalFiles.Any(f => f.FilePath == file)) continue;
                    var sheets = ExcelHelper.GetSheetNames(file);
                    var defaultSheet = sheets.FirstOrDefault(s => s.Equals("Inspection_Task", StringComparison.OrdinalIgnoreCase)) ?? sheets.FirstOrDefault();
                    _additionalFiles.Add(new ExcelFileModel
                    {
                        FilePath = file,
                        Sheets = sheets,
                        SelectedSheet = defaultSheet
                    });
                }
            }
        }

        private void BtnMergeAndSave_Click(object sender, RoutedEventArgs e)
        {
            if (_masterFile == null || string.IsNullOrEmpty(_masterFile.SelectedSheet))
            {
                TxtStatus.Text = "Please select a master file and sheet.";
                return;
            }
            if (_additionalFiles.Count == 0)
            {
                TxtStatus.Text = "Please add at least one file to merge.";
                return;
            }
            var filesToMerge = _additionalFiles
                .Where(f => !string.IsNullOrEmpty(f.SelectedSheet))
                .Select(f => (f.FilePath, f.SelectedSheet!)).ToList();
            if (filesToMerge.Count == 0)
            {
                TxtStatus.Text = "No valid files/sheets to merge.";
                return;
            }
            var saveDlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FileName = "Merged.xlsx"
            };
            if (saveDlg.ShowDialog() == true)
            {
                try
                {
                    ExcelHelper.MergeFiles(_masterFile.FilePath, _masterFile.SelectedSheet!, filesToMerge, saveDlg.FileName);
                    TxtStatus.Text = $"Merge complete! Saved to: {saveDlg.FileName}";
                }
                catch (Exception ex)
                {
                    TxtStatus.Text = $"Error: {ex.Message}";
                }
            }
        }

        private void ListFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // No-op for now
        }

        private void BtnRemoveFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is ExcelFileModel fileModel)
            {
                _additionalFiles.Remove(fileModel);
            }
        }

        private DataTemplate CreateFileTemplate()
        {
            var template = new DataTemplate(typeof(ExcelFileModel));
            var spFactory = new FrameworkElementFactory(typeof(StackPanel));
            spFactory.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);
            spFactory.SetValue(StackPanel.MarginProperty, new Thickness(0, 0, 0, 5));

            var tbFactory = new FrameworkElementFactory(typeof(TextBlock));
            tbFactory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("FilePath") { Converter = new FileNameConverter() });
            tbFactory.SetValue(TextBlock.WidthProperty, 280.0);
            tbFactory.SetValue(TextBlock.MarginProperty, new Thickness(0, 0, 10, 0));
            tbFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            spFactory.AppendChild(tbFactory);

            var cbFactory = new FrameworkElementFactory(typeof(ComboBox));
            cbFactory.SetBinding(ComboBox.ItemsSourceProperty, new System.Windows.Data.Binding("Sheets"));
            cbFactory.SetBinding(ComboBox.SelectedItemProperty, new System.Windows.Data.Binding("SelectedSheet") { Mode = System.Windows.Data.BindingMode.TwoWay });
            cbFactory.SetValue(ComboBox.WidthProperty, 180.0);
            cbFactory.SetValue(ComboBox.MarginProperty, new Thickness(0, 0, 10, 0));
            spFactory.AppendChild(cbFactory);

            var btnFactory = new FrameworkElementFactory(typeof(Button));
            btnFactory.SetValue(Button.ContentProperty, "✕");
            btnFactory.SetValue(Button.WidthProperty, 30.0);
            btnFactory.SetValue(Button.HeightProperty, 30.0);
            btnFactory.SetValue(Button.BackgroundProperty, new SolidColorBrush(Color.FromRgb(231, 76, 60)));
            btnFactory.SetValue(Button.ForegroundProperty, Brushes.White);
            btnFactory.SetValue(Button.FontWeightProperty, FontWeights.Bold);
            btnFactory.SetValue(Button.FontSizeProperty, 14.0);
            btnFactory.SetValue(Button.BorderThicknessProperty, new Thickness(0));
            btnFactory.SetValue(Button.CursorProperty, Cursors.Hand);
            btnFactory.SetValue(Button.ToolTipProperty, "Remove this file");
            btnFactory.AddHandler(Button.ClickEvent, new RoutedEventHandler(BtnRemoveFile_Click));

            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(15));
            borderFactory.SetValue(Border.BackgroundProperty, new SolidColorBrush(Color.FromRgb(231, 76, 60)));
            borderFactory.SetValue(Border.WidthProperty, 30.0);
            borderFactory.SetValue(Border.HeightProperty, 30.0);
            borderFactory.AppendChild(btnFactory);

            spFactory.AppendChild(borderFactory);

            template.VisualTree = spFactory;
            return template;
        }
    }

    public class FileNameConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value is string s ? System.IO.Path.GetFileName(s) : string.Empty;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}