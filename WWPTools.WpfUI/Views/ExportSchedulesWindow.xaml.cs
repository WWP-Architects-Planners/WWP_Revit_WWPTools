using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace WWPTools.WpfUI.Views
{
    public partial class ExportSchedulesWindow : Window
    {
        public ExportSchedulesWindow()
        {
            InitializeComponent();
            DialogStyles.ApplyPrimaryButtonStyle(OkButton);
            DialogStyles.ApplyPrimaryButtonStyle(CancelButton);
            DialogStyles.ApplyToggleStyle(ModeExcel);
            DialogStyles.ApplyToggleStyle(ModeCsv);
            OkButton.Click += (_, __) => DialogResult = true;
            CancelButton.Click += (_, __) => DialogResult = false;
            ModeExcel.Checked += (_, __) => SetMode(true);
            ModeCsv.Checked += (_, __) => SetMode(false);
            ModeExcel.Unchecked += (_, __) => EnsureModeSelected();
            ModeCsv.Unchecked += (_, __) => EnsureModeSelected();
            BrowseExcelButton.Click += (_, __) => BrowseExcel();
            BrowseCsvButton.Click += (_, __) => BrowseCsv();
            LoadLogo();
        }

        public void Initialize(
            IList<string> items,
            string prompt,
            string modeLabelExcel,
            string modeLabelCsv,
            int defaultMode,
            ISet<int> prechecked,
            string excelPath,
            string csvFolder,
            string csvDelimiter,
            bool csvQuoteAll)
        {
            PromptText.Text = prompt ?? "";
            ModeExcel.Content = string.IsNullOrWhiteSpace(modeLabelExcel) ? "Export to Excel" : modeLabelExcel;
            ModeCsv.Content = string.IsNullOrWhiteSpace(modeLabelCsv) ? "Export to CSV" : modeLabelCsv;

            ItemsList.Items.Clear();
            if (items != null)
            {
                foreach (var item in items)
                    ItemsList.Items.Add(item);
            }

            if (prechecked != null)
            {
                foreach (var index in prechecked)
                {
                    if (index >= 0 && index < ItemsList.Items.Count)
                        ItemsList.SelectedItems.Add(ItemsList.Items[index]);
                }
            }

            ExcelPathBox.Text = excelPath ?? "";
            CsvFolderBox.Text = csvFolder ?? "";
            QuoteAllCheck.IsChecked = csvQuoteAll;

            SetDelimiterSelection(csvDelimiter);

            if (defaultMode == 1)
            {
                ModeCsv.IsChecked = true;
                ModeExcel.IsChecked = false;
            }
            else
            {
                ModeExcel.IsChecked = true;
                ModeCsv.IsChecked = false;
            }

            UpdateModeVisibility();
        }

        public int[] GetSelectedIndices()
        {
            var selected = new List<int>();
            foreach (var item in ItemsList.SelectedItems)
            {
                var index = ItemsList.Items.IndexOf(item);
                if (index >= 0)
                    selected.Add(index);
            }
            selected.Sort();
            return selected.ToArray();
        }

        public int GetSelectedMode()
        {
            return ModeCsv.IsChecked == true ? 1 : 0;
        }

        public string GetExcelPath()
        {
            return ExcelPathBox.Text ?? "";
        }

        public string GetCsvFolder()
        {
            return CsvFolderBox.Text ?? "";
        }

        public string GetCsvDelimiter()
        {
            if (DelimiterCombo.SelectedItem is ComboBoxItem item && item.Tag != null)
                return item.Tag.ToString();
            return ",";
        }

        public bool GetCsvQuoteAll()
        {
            return QuoteAllCheck.IsChecked == true;
        }

        private void SetDelimiterSelection(string delimiter)
        {
            var target = delimiter ?? ",";
            foreach (var obj in DelimiterCombo.Items)
            {
                if (obj is ComboBoxItem item && item.Tag != null && item.Tag.ToString() == target)
                {
                    DelimiterCombo.SelectedItem = item;
                    return;
                }
            }
            DelimiterCombo.SelectedIndex = 0;
        }

        private void SetMode(bool excelMode)
        {
            if (excelMode)
            {
                if (ModeCsv.IsChecked == true)
                    ModeCsv.IsChecked = false;
                ModeExcel.IsChecked = true;
            }
            else
            {
                if (ModeExcel.IsChecked == true)
                    ModeExcel.IsChecked = false;
                ModeCsv.IsChecked = true;
            }
            UpdateModeVisibility();
        }

        private void EnsureModeSelected()
        {
            if (ModeExcel.IsChecked != true && ModeCsv.IsChecked != true)
                ModeExcel.IsChecked = true;
            UpdateModeVisibility();
        }

        private void UpdateModeVisibility()
        {
            var isCsv = ModeCsv.IsChecked == true;
            ExcelPanel.Visibility = isCsv ? Visibility.Collapsed : Visibility.Visible;
            CsvOptionsPanel.Visibility = isCsv ? Visibility.Visible : Visibility.Collapsed;
            CsvFolderPanel.Visibility = isCsv ? Visibility.Visible : Visibility.Collapsed;
        }

        private void BrowseExcel()
        {
            var initialDir = SafeDirectory(ExcelPathBox.Text);
            var path = DialogService.SaveFileDialog(
                "Export Schedules",
                "Excel Workbook (*.xlsx)|*.xlsx",
                "xlsx",
                initialDir,
                Path.GetFileName(ExcelPathBox.Text ?? ""));
            if (!string.IsNullOrWhiteSpace(path))
                ExcelPathBox.Text = path;
        }

        private void BrowseCsv()
        {
            var initialDir = SafeDirectory(CsvFolderBox.Text);
            var path = DialogService.SelectFolderDialog("Select CSV Folder", initialDir);
            if (!string.IsNullOrWhiteSpace(path))
                CsvFolderBox.Text = path;
        }

        private static string SafeDirectory(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path))
                {
                    var dir = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
                    if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir))
                        return dir;
                }
            }
            catch
            {
            }
            return "";
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
}
