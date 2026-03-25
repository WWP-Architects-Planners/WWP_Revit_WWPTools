using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace WWPTools.WpfUI.Views
{
    public partial class RoundAnglesWindow : Window
    {
        private readonly ObservableCollection<RoundAnglesItem> _visibleItems = new ObservableCollection<RoundAnglesItem>();
        private List<RoundAnglesItem> _allItems = new List<RoundAnglesItem>();
        private int _matchedWarnings;
        private string _skippedText = "Skipped: None";
        private string _logPath = "";
        private int? _groupFilterId;
        private string _groupFilterLabel = "";

        internal RoundAnglesController Controller { get; set; }

        public RoundAnglesWindow()
        {
            InitializeComponent();
            ResultsList.ItemsSource = _visibleItems;

            DialogStyles.ApplyPrimaryButtonStyle(BtnRefresh);
            DialogStyles.ApplySecondaryButtonStyle(BtnRemoveSelected);
            DialogStyles.ApplySecondaryButtonStyle(BtnResetList);
            DialogStyles.ApplyPrimaryButtonStyle(BtnApply);
            DialogStyles.ApplySecondaryButtonStyle(BtnClose);
            DialogStyles.ApplyPrimaryButtonStyle(BtnSelect);
            DialogStyles.ApplySecondaryButtonStyle(BtnGoToView);
            DialogStyles.ApplySecondaryButtonStyle(BtnEditGroup);
            DialogStyles.ApplySecondaryButtonStyle(BtnEditBoundary);
            DialogStyles.ApplySecondaryButtonStyle(BtnZoom);
            DialogStyles.ApplySecondaryButtonStyle(BtnIsolate);

            BtnRefresh.Click += (_, __) => Controller?.RequestRefresh();
            BtnRemoveSelected.Click += (_, __) => RemoveSelected();
            BtnResetList.Click += (_, __) => ResetList();
            BtnApply.Click += (_, __) => Controller?.RequestApply(GetApplicableItems());
            BtnClose.Click += (_, __) => Close();
            BtnSelect.Click += (_, __) => Controller?.RequestSelect(GetCurrentItem());
            BtnGoToView.Click += (_, __) => Controller?.RequestGoToView(GetCurrentItem());
            BtnEditGroup.Click += (_, __) => ApplyGroupFilterAndEdit(GetCurrentItem());
            BtnEditBoundary.Click += (_, __) => Controller?.RequestEditBoundary(GetCurrentItem());
            BtnZoom.Click += (_, __) => Controller?.RequestZoomTo(GetCurrentItem());
            BtnIsolate.Click += (_, __) => Controller?.RequestIsolate(GetCurrentItem());
            ResultsList.SelectionChanged += (_, __) => OnSelectionChanged();
            Closed += (_, __) => Controller?.NotifyWindowClosed();

            LoadLogo();
            UpdateButtons();
        }

        internal void SetOwnerHandle(IntPtr handle)
        {
            if (handle == IntPtr.Zero)
                return;

            new WindowInteropHelper(this) { Owner = handle };
        }

        internal void LoadItems(int matchedWarnings, string skippedText, IEnumerable<RoundAnglesItem> items, string logPath)
        {
            _matchedWarnings = matchedWarnings;
            _skippedText = string.IsNullOrWhiteSpace(skippedText) ? "Skipped: None" : skippedText;
            _logPath = logPath ?? "";
            _allItems = (items ?? Enumerable.Empty<RoundAnglesItem>()).ToList();

            RepopulateVisibleItems();

            TxtLogPath.Text = string.IsNullOrWhiteSpace(_logPath) ? "(log file is written after apply)" : _logPath;
            if (_visibleItems.Count > 0)
                ResultsList.SelectedIndex = 0;
            else
                TxtDetail.Text = "";

            UpdateSummary();
            UpdateButtons();
        }

        internal void RefreshPresentation(string logPath)
        {
            _logPath = logPath ?? "";
            TxtLogPath.Text = string.IsNullOrWhiteSpace(_logPath) ? "(log file is written after apply)" : _logPath;
            TxtDetail.Text = GetCurrentItem()?.DetailText ?? "";
            ResultsList.Items.Refresh();
            UpdateSummary();
            UpdateButtons();
        }

        internal void ShowInfo(string message, string title)
        {
            MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        internal void ShowError(string message, string title)
        {
            MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void OnSelectionChanged()
        {
            TxtDetail.Text = GetCurrentItem()?.DetailText ?? "";
            UpdateButtons();
            Controller?.RequestSelect(GetCurrentItem());
        }

        private RoundAnglesItem GetCurrentItem()
        {
            return ResultsList.SelectedItem as RoundAnglesItem;
        }

        private IList<RoundAnglesItem> GetApplicableItems()
        {
            return _visibleItems.Where(item => item.CanApply).ToList();
        }

        private void RemoveSelected()
        {
            var selected = ResultsList.SelectedItems.Cast<RoundAnglesItem>().ToList();
            if (selected.Count == 0)
                return;

            var nextSelection = _visibleItems.Except(selected).FirstOrDefault();
            foreach (var item in selected)
                _visibleItems.Remove(item);

            if (nextSelection != null)
                ResultsList.SelectedItem = nextSelection;
            else
                TxtDetail.Text = "";

            UpdateSummary();
            UpdateButtons();
        }

        private void ResetList()
        {
            _groupFilterId = null;
            _groupFilterLabel = "";
            FooterText.Text = "Selection follows the highest parent element available.";
            RepopulateVisibleItems();

            if (_visibleItems.Count > 0)
                ResultsList.SelectedIndex = 0;
            else
                TxtDetail.Text = "";

            UpdateSummary();
            UpdateButtons();
        }

        private void UpdateSummary()
        {
            var pending = _visibleItems.Count(item => item.Status == RoundAnglesItemStatus.Pending && !item.IsGroupMember);
            var updated = _visibleItems.Count(item => item.Status == RoundAnglesItemStatus.Updated);
            var failed = _visibleItems.Count(item => item.Status == RoundAnglesItemStatus.Failed);
            var grouped = _visibleItems.Count(item => item.IsGroupMember);

            TxtSummary.Text = string.Format(
                "Matched warnings: {0}    Pending fixes: {1}    Updated: {2}    Failed: {3}    Grouped skips: {4}",
                _matchedWarnings,
                pending,
                updated,
                failed,
                grouped);
            TxtSkipped.Text = string.IsNullOrWhiteSpace(_groupFilterLabel)
                ? _skippedText
                : _skippedText + " | Active group filter: " + _groupFilterLabel;
        }

        private void UpdateButtons()
        {
            var current = GetCurrentItem();
            var hasCurrent = current != null;
            var hasVisible = _visibleItems.Count > 0;

            BtnRemoveSelected.IsEnabled = ResultsList.SelectedItems.Count > 0;
            BtnResetList.IsEnabled = _visibleItems.Count != _allItems.Count;
            BtnApply.IsEnabled = _visibleItems.Any(item => item.CanApply);

            BtnSelect.IsEnabled = hasCurrent && current.PreferredUiElementIdValue.HasValue;
            BtnGoToView.IsEnabled = hasCurrent;
            BtnEditGroup.IsEnabled = hasCurrent && current.GroupElementIdValue.HasValue;
            BtnEditBoundary.IsEnabled = hasCurrent;
            BtnZoom.IsEnabled = hasCurrent && current.PreferredUiElementIdValue.HasValue;
            BtnIsolate.IsEnabled = hasCurrent && current.PreferredUiElementIdValue.HasValue;
            BtnClose.IsEnabled = true;
            BtnRefresh.IsEnabled = true;
            ResultsList.IsEnabled = hasVisible;
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
                // Ignore logo loading issues.
            }
        }

        private void ApplyGroupFilterAndEdit(RoundAnglesItem item)
        {
            if (item?.GroupElementIdValue != null)
            {
                _groupFilterId = item.GroupElementIdValue;
                _groupFilterLabel = item.HostGroupText;
                FooterText.Text = string.IsNullOrWhiteSpace(_groupFilterLabel)
                    ? "Filtered to active group."
                    : "Filtered to group: " + _groupFilterLabel;
                RepopulateVisibleItems();
                if (_visibleItems.Count > 0)
                    ResultsList.SelectedIndex = 0;
            }

            Controller?.RequestEditGroup(item);
        }

        private void RepopulateVisibleItems()
        {
            _visibleItems.Clear();
            foreach (var item in FilteredItems())
                _visibleItems.Add(item);
        }

        private IEnumerable<RoundAnglesItem> FilteredItems()
        {
            if (!_groupFilterId.HasValue)
                return _allItems;

            return _allItems.Where(item => item.GroupElementIdValue == _groupFilterId.Value);
        }
    }
}
