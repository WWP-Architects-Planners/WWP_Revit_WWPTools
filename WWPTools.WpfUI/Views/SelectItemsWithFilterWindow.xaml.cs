using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace WWPTools.WpfUI.Views
{
    public partial class SelectItemsWithFilterWindow : Window
    {
        private class ItemEntry
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public override string ToString() => Text ?? "";
        }

        private List<ItemEntry> _allItems = new List<ItemEntry>();
        private List<ItemEntry> _filteredItems = new List<ItemEntry>();
        private Dictionary<string, List<string>> _valuesByParam = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        private HashSet<int> _selectedIndices = new HashSet<int>();
        private bool _isUpdating;

        public string SelectedFilterParameter { get; private set; } = "";
        public string SelectedFilterValue { get; private set; } = "";

        public SelectItemsWithFilterWindow()
        {
            InitializeComponent();
            DialogStyles.ApplyPrimaryButtonStyle(OkButton);
            DialogStyles.ApplyPrimaryButtonStyle(CancelButton);
            DialogStyles.ApplyPrimaryButtonStyle(ApplyFilterButton);
            DialogStyles.ApplyPrimaryButtonStyle(ClearFilterButton);
            DialogStyles.ApplyPrimaryButtonStyle(SelectAllButton);
            DialogStyles.ApplyPrimaryButtonStyle(SelectNoneButton);
            DialogStyles.ApplyPrimaryButtonStyle(InvertSelectionButton);
            OkButton.Click += (_, __) => DialogResult = true;
            CancelButton.Click += (_, __) => DialogResult = false;
            ApplyFilterButton.Click += (_, __) => ApplyFilter();
            ClearFilterButton.Click += (_, __) => ClearFilter();
            FilterParamCombo.SelectionChanged += (_, __) => UpdateValueOptions();
            ItemsList.SelectionChanged += (_, __) => UpdateSelectedIndices();
            SearchBox.TextChanged += (_, __) => ApplySearch();
            SelectAllButton.Click += (_, __) => SelectAllVisible();
            SelectNoneButton.Click += (_, __) => SelectNoneVisible();
            InvertSelectionButton.Click += (_, __) => InvertVisible();
            LoadLogo();
        }

        public void Initialize(
            IList<string> items,
            string prompt,
            IList<string> filterParams,
            IList<IList<string>> filterValuesByParam,
            ISet<int> prechecked,
            string defaultFilterParam,
            string defaultFilterValue)
        {
            PromptText.Text = prompt ?? "";

            _allItems = new List<ItemEntry>();
            if (items != null)
            {
                for (var i = 0; i < items.Count; i++)
                {
                    _allItems.Add(new ItemEntry { Index = i, Text = items[i] ?? "" });
                }
            }

            _valuesByParam = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            if (filterParams != null && filterValuesByParam != null)
            {
                var count = Math.Min(filterParams.Count, filterValuesByParam.Count);
                for (var i = 0; i < count; i++)
                {
                    var name = filterParams[i] ?? "";
                    var values = filterValuesByParam[i] ?? new List<string>();
                    _valuesByParam[name] = values.Select(v => v ?? "").ToList();
                }
            }

            FilterParamCombo.Items.Clear();
            if (filterParams != null)
            {
                foreach (var p in filterParams)
                {
                    FilterParamCombo.Items.Add(p ?? "");
                }
            }
            if (FilterParamCombo.Items.Count > 0)
            {
                var defaultIndex = -1;
                if (!string.IsNullOrWhiteSpace(defaultFilterParam))
                    defaultIndex = FilterParamCombo.Items.IndexOf(defaultFilterParam);
                FilterParamCombo.SelectedIndex = defaultIndex >= 0 ? defaultIndex : 0;
            }

            _selectedIndices = new HashSet<int>(prechecked ?? new HashSet<int>());
            UpdateValueOptions();
            if (FilterValueCombo.Items.Count > 0 && !string.IsNullOrWhiteSpace(defaultFilterValue))
            {
                var valueIndex = FilterValueCombo.Items.IndexOf(defaultFilterValue);
                if (valueIndex >= 0)
                    FilterValueCombo.SelectedIndex = valueIndex;
            }
            _filteredItems = new List<ItemEntry>(_allItems);
            RenderItems(_filteredItems);
        }

        public int[] GetSelectedIndices()
        {
            var selected = _selectedIndices.ToList();
            selected.Sort();
            return selected.ToArray();
        }

        private void RenderItems(IEnumerable<ItemEntry> items)
        {
            _isUpdating = true;
            ItemsList.Items.Clear();
            foreach (var item in items)
                ItemsList.Items.Add(item);

            foreach (ItemEntry item in ItemsList.Items)
            {
                if (_selectedIndices.Contains(item.Index))
                    ItemsList.SelectedItems.Add(item);
            }
            _isUpdating = false;
            UpdateCountLabel();
        }

        private void UpdateSelectedIndices()
        {
            if (_isUpdating)
                return;
            _selectedIndices.Clear();
            foreach (ItemEntry item in ItemsList.SelectedItems)
                _selectedIndices.Add(item.Index);
            UpdateCountLabel();
        }

        private void UpdateValueOptions()
        {
            var param = FilterParamCombo.SelectedItem as string ?? "";
            FilterValueCombo.Items.Clear();
            if (string.IsNullOrWhiteSpace(param) || !_valuesByParam.ContainsKey(param))
                return;
            var values = _valuesByParam[param] ?? new List<string>();
            var unique = values.Distinct().OrderBy(v => v).ToList();
            foreach (var v in unique)
                FilterValueCombo.Items.Add(v);
            if (FilterValueCombo.Items.Count > 0)
                FilterValueCombo.SelectedIndex = 0;
        }

        private void ApplyFilter()
        {
            var param = FilterParamCombo.SelectedItem as string ?? "";
            var value = FilterValueCombo.SelectedItem as string ?? "";
            if (string.IsNullOrWhiteSpace(param) || !_valuesByParam.ContainsKey(param))
            {
                _filteredItems = new List<ItemEntry>(_allItems);
                ApplySearch();
                return;
            }

            SelectedFilterParameter = param;
            SelectedFilterValue = value ?? "";

            var values = _valuesByParam[param] ?? new List<string>();
            _filteredItems = _allItems.Where(i =>
            {
                if (i.Index < 0 || i.Index >= values.Count)
                    return false;
                var v = values[i.Index] ?? "";
                return string.Equals(v, value ?? "", StringComparison.OrdinalIgnoreCase);
            }).ToList();

            ApplySearch();
        }

        private void ClearFilter()
        {
            SelectedFilterParameter = "";
            SelectedFilterValue = "";
            _filteredItems = new List<ItemEntry>(_allItems);
            ApplySearch();
        }

        private void ApplySearch()
        {
            var term = (SearchBox.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(term))
            {
                RenderItems(_filteredItems);
                return;
            }
            var filtered = _filteredItems.Where(i =>
                i.Text != null &&
                i.Text.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            RenderItems(filtered);
        }

        private void SelectAllVisible()
        {
            foreach (ItemEntry item in ItemsList.Items)
                _selectedIndices.Add(item.Index);
            RenderItems(ItemsList.Items.Cast<ItemEntry>());
        }

        private void SelectNoneVisible()
        {
            foreach (ItemEntry item in ItemsList.Items)
                _selectedIndices.Remove(item.Index);
            RenderItems(ItemsList.Items.Cast<ItemEntry>());
        }

        private void InvertVisible()
        {
            foreach (ItemEntry item in ItemsList.Items)
            {
                if (_selectedIndices.Contains(item.Index))
                    _selectedIndices.Remove(item.Index);
                else
                    _selectedIndices.Add(item.Index);
            }
            RenderItems(ItemsList.Items.Cast<ItemEntry>());
        }

        private void UpdateCountLabel()
        {
            var total = _allItems.Count;
            var visible = ItemsList.Items.Count;
            var selected = _selectedIndices.Count;
            CountLabel.Text = $"{selected} selected | {visible} visible | {total} total";
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
