using System.Collections.Generic;
using System.Windows;
using WWPTools.WpfUI.Views;

namespace WWPTools.WpfUI.TestHarness;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        OpenDialogButton.Click += (_, __) => ShowSelectItemsDialog();
    }

    private void ShowSelectItemsDialog()
    {
        var items = new List<string>
        {
            "Schedule - Doors",
            "Schedule - Rooms",
            "Schedule - Windows",
            "Schedule - Finishes"
        };

        var window = new SelectItemsWithModeWindow
        {
            Title = "Export Schedules",
            Width = 720,
            Height = 620,
            Owner = this
        };

        window.Initialize(
            items,
            "Select schedules to export:",
            "Export to Excel",
            "Export to CSV",
            0,
            new HashSet<int> { 0, 2 });

        if (window.ShowDialog() == true)
        {
            var selected = string.Join(", ", window.GetSelectedIndices());
            var mode = window.GetSelectedMode() == 0 ? "Excel" : "CSV";
            MessageBox.Show(this, $"Selected indices: {selected}\nMode: {mode}", "Result");
        }
    }
}
