using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace WWPTools.WpfUI.Views
{
    public partial class AreaKeyplanImportWindow : Window
    {
        private readonly ObservableCollection<ColumnMapping> _columnMappings = new ObservableCollection<ColumnMapping>();
        private readonly ObservableCollection<string> _parameterOptions = new ObservableCollection<string>();

        public AreaKeyplanImportWindow()
        {
            InitializeComponent();
            DataContext = this;
            DialogStyles.ApplyPrimaryButtonStyle(OkButton);
            DialogStyles.ApplyPrimaryButtonStyle(CancelButton);
            DialogStyles.ApplyPrimaryButtonStyle(BrowseButton);
            DialogStyles.ApplyPrimaryButtonStyle(LoadButton);

            OkButton.Click += (_, __) =>
            {
                LoadRequested = false;
                DialogResult = true;
            };
            CancelButton.Click += (_, __) => DialogResult = false;
            BrowseButton.Click += (_, __) => BrowseForFile();
            LoadButton.Click += (_, __) =>
            {
                LoadRequested = true;
                DialogResult = true;
            };

            MappingGrid.ItemsSource = _columnMappings;
            var comboColumn = MappingGrid.Columns.OfType<DataGridComboBoxColumn>().FirstOrDefault();
            if (comboColumn != null)
                comboColumn.ItemsSource = _parameterOptions;
            LoadLogo();
        }

        public bool LoadRequested { get; private set; }

        public IList<string> ParameterOptions => _parameterOptions;

        public string FilePath => FilePathBox.Text ?? "";

        public void Initialize(
            string filePath,
            IList<string> columnNames,
            IList<string> previewLines,
            IList<string> parameterOptions,
            IList<string> defaultSelections)
        {
            FilePathBox.Text = filePath ?? "";

            _parameterOptions.Clear();
            if (parameterOptions != null)
            {
                foreach (var option in parameterOptions)
                    _parameterOptions.Add(option);
            }

            _columnMappings.Clear();
            if (columnNames != null)
            {
                for (var i = 0; i < columnNames.Count; i++)
                {
                    var column = columnNames[i] ?? "";
                    var mapping = new ColumnMapping { ColumnName = column };
                    if (defaultSelections != null && i < defaultSelections.Count)
                    {
                        var selection = defaultSelections[i];
                        if (!string.IsNullOrWhiteSpace(selection) && _parameterOptions.Contains(selection))
                            mapping.SelectedOption = selection;
                    }
                    if (string.IsNullOrWhiteSpace(mapping.SelectedOption) && _parameterOptions.Count > 0)
                        mapping.SelectedOption = _parameterOptions[0];
                    _columnMappings.Add(mapping);
                }
            }

            var previewText = "";
            if (previewLines != null && previewLines.Count > 0)
                previewText = string.Join(Environment.NewLine, previewLines);
            PreviewBox.Text = previewText;

            var hasData = _columnMappings.Count > 0;
            MappingGrid.IsEnabled = hasData;
            OkButton.IsEnabled = hasData;
        }

        public IList<ColumnMapping> GetMappings()
        {
            return _columnMappings.ToList();
        }

        private void BrowseForFile()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Excel File",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                CheckFileExists = true
            };

            var result = dialog.ShowDialog();
            if (result == true)
                FilePathBox.Text = dialog.FileName ?? "";
        }

        private void LoadLogo()
        {
            try
            {
                var assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (string.IsNullOrWhiteSpace(assemblyPath))
                    return;
                var logoPath = Path.Combine(assemblyPath, "WWPtools-logo.png");
                if (!File.Exists(logoPath))
                    return;

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(logoPath);
                bitmap.EndInit();
                LogoImage.Source = bitmap;
            }
            catch
            {
            }
        }
    }

    public class ColumnMapping
    {
        public string ColumnName { get; set; } = "";
        public string SelectedOption { get; set; } = "";
    }
}
