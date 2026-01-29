using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Interop;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WWPTools.WpfUI
{
    public static class DialogService
    {
        private const string DefaultFont = "Arial Narrow";
        private const double DefaultFontSize = 14;

        public static int[] SelectIndices(
            IEnumerable items,
            string title,
            string prompt,
            bool multiselect,
            int width,
            int height)
        {
            return SelectIndices(AsStringList(items), title, prompt, multiselect, width, height);
        }

        public static void Alert(string message, string title)
        {
            var window = CreateMessageWindow(message, title, showCancel: false);
            window.ShowDialog();
        }

        public static bool Confirm(string message, string title)
        {
            var window = CreateMessageWindow(message, title, showCancel: true);
            return window.ShowDialog() == true;
        }

        public static string OpenFileDialog(string title, string filter, bool multiselect, string initialDirectory)
        {
            EnsureApplication();
            var dialog = new OpenFileDialog
            {
                Title = title ?? "Open File",
                Filter = string.IsNullOrWhiteSpace(filter) ? "All files (*.*)|*.*" : filter,
                Multiselect = multiselect,
                InitialDirectory = string.IsNullOrWhiteSpace(initialDirectory) ? null : initialDirectory
            };
            var result = dialog.ShowDialog();
            if (result != true)
                return null;
            if (multiselect)
                return string.Join("|", dialog.FileNames ?? Array.Empty<string>());
            return dialog.FileName;
        }

        public static string SaveFileDialog(string title, string filter, string defaultExtension, string initialDirectory, string fileName)
        {
            EnsureApplication();
            var dialog = new SaveFileDialog
            {
                Title = title ?? "Save File",
                Filter = string.IsNullOrWhiteSpace(filter) ? "All files (*.*)|*.*" : filter,
                DefaultExt = string.IsNullOrWhiteSpace(defaultExtension) ? null : defaultExtension,
                InitialDirectory = string.IsNullOrWhiteSpace(initialDirectory) ? null : initialDirectory,
                FileName = fileName ?? ""
            };
            var result = dialog.ShowDialog();
            if (result != true)
                return null;
            return dialog.FileName;
        }

        public static string SelectFolderDialog(string title, string initialDirectory)
        {
            EnsureApplication();
            return BrowseForFolder(title ?? "Select Folder", initialDirectory);
        }

        public static string PromptText(
            string title,
            string prompt,
            string defaultValue,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };
            content.Children.Add(CreatePrompt(prompt));
            var input = CreateTextBox();
            input.Text = defaultValue ?? "";
            content.Children.Add(input);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            if (window.ShowDialog() != true)
                return null;
            return input.Text ?? "";
        }

        public static bool ShowTextReport(
            string title,
            string text,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            var textBox = new TextBox
            {
                Text = text ?? "",
                IsReadOnly = true,
                TextWrapping = TextWrapping.NoWrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                MinHeight = Math.Max(200, height - 140)
            };
            content.Children.Add(textBox);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            if (buttons.Cancel != null)
                buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            return window.ShowDialog() == true;
        }

        public static int[] SelectIndices(
            IList<string> items,
            string title,
            string prompt,
            bool multiselect,
            int width,
            int height)
        {
            if (items == null)
                items = Array.Empty<string>();

            var window = CreateWindow(title, width, height);
            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var promptBlock = CreatePrompt(prompt);
            Grid.SetRow(promptBlock, 0);
            grid.Children.Add(promptBlock);

            var listBox = new ListBox
            {
                SelectionMode = multiselect ? SelectionMode.Extended : SelectionMode.Single,
                MinHeight = 200,
                Margin = new Thickness(0, 0, 0, 8)
            };
            ScrollViewer.SetVerticalScrollBarVisibility(listBox, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(listBox, ScrollBarVisibility.Disabled);
            foreach (var item in items)
            {
                listBox.Items.Add(item);
            }
            Grid.SetRow(listBox, 1);
            grid.Children.Add(listBox);

            var buttons = CreateOkCancelButtons("OK", "Cancel");
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            Grid.SetRow(buttons.Panel, 2);
            grid.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(grid);

            if (window.ShowDialog() != true)
                return Array.Empty<int>();

            if (multiselect)
                return GetSelectedIndices(listBox);

            return listBox.SelectedIndex >= 0 ? new[] { listBox.SelectedIndex } : Array.Empty<int>();
        }

        public static SelectionWithModeResult SelectItemsWithMode(
            IEnumerable items,
            string title,
            string prompt,
            string modeLabelA,
            string modeLabelB,
            int defaultMode,
            IEnumerable precheckedIndices,
            int width,
            int height)
        {
            return SelectItemsWithMode(
                AsStringList(items),
                title,
                prompt,
                modeLabelA,
                modeLabelB,
                defaultMode,
                AsIntArray(precheckedIndices),
                width,
                height);
        }

        public static SelectionWithModeResult SelectItemsWithMode(
            IList<string> items,
            string title,
            string prompt,
            string modeLabelA,
            string modeLabelB,
            int defaultMode,
            int[] precheckedIndices,
            int width,
            int height)
        {
            if (items == null)
                items = Array.Empty<string>();

            EnsureApplication();
            var prechecked = new HashSet<int>(precheckedIndices ?? Array.Empty<int>());
            var window = new Views.SelectItemsWithModeWindow
            {
                Title = title ?? "Select Items",
                Width = width > 0 ? width : 720,
                Height = height > 0 ? height : 620
            };
            var owner = GetOwnerHandle();
            if (owner != IntPtr.Zero)
                new WindowInteropHelper(window) { Owner = owner };

            window.Initialize(items, prompt, modeLabelA, modeLabelB, defaultMode, prechecked);

            if (window.ShowDialog() != true)
                return null;

            return new SelectionWithModeResult
            {
                SelectedIndices = window.GetSelectedIndices(),
                Mode = window.GetSelectedMode()
            };
        }

        public static SheetRenumberInputsResult SelectSheetRenumberInputs(
            IEnumerable categories,
            IEnumerable printSets,
            string title,
            string categoryLabel,
            string printSetLabel,
            string startingLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            return SelectSheetRenumberInputs(
                AsStringList(categories),
                AsStringList(printSets),
                title,
                categoryLabel,
                printSetLabel,
                startingLabel,
                okText,
                cancelText,
                width,
                height);
        }

        public static SheetRenumberInputsResult SelectSheetRenumberInputs(
            IList<string> categories,
            IList<string> printSets,
            string title,
            string categoryLabel,
            string printSetLabel,
            string startingLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            ComboBox categoryCombo = null;
            ComboBox printsetCombo = null;

            if (categories != null)
            {
                content.Children.Add(CreateLabel(categoryLabel));
                categoryCombo = CreateCombo(categories);
                content.Children.Add(categoryCombo);
            }

            if (printSets != null)
            {
                content.Children.Add(CreateLabel(printSetLabel));
                printsetCombo = CreateCombo(printSets);
                content.Children.Add(printsetCombo);
            }

            content.Children.Add(CreateLabel(startingLabel));
            var startInput = CreateTextBox();
            content.Children.Add(startInput);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);

            if (window.ShowDialog() != true)
                return null;

            return new SheetRenumberInputsResult
            {
                Category = categoryCombo != null && categoryCombo.SelectedItem != null ? categoryCombo.SelectedItem.ToString() : "",
                PrintSet = printsetCombo != null && printsetCombo.SelectedItem != null ? printsetCombo.SelectedItem.ToString() : "",
                StartingNumber = startInput.Text ?? ""
            };
        }

        public static SheetRenumberInputsWithListResult SelectSheetRenumberInputsWithList(
            IEnumerable items,
            string title,
            string prompt,
            string startingLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            return SelectSheetRenumberInputsWithList(
                AsStringList(items),
                title,
                prompt,
                startingLabel,
                okText,
                cancelText,
                width,
                height);
        }

        public static SheetRenumberInputsWithListResult SelectSheetRenumberInputsWithList(
            IList<string> items,
            string title,
            string prompt,
            string startingLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            if (items == null)
                items = Array.Empty<string>();

            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };
            content.Children.Add(CreatePrompt(prompt));

            var listBox = new ListBox
            {
                SelectionMode = SelectionMode.Extended,
                Height = Math.Max(200, height - 220),
                Margin = new Thickness(0, 0, 0, 8)
            };
            ScrollViewer.SetVerticalScrollBarVisibility(listBox, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(listBox, ScrollBarVisibility.Disabled);
            for (var i = 0; i < items.Count; i++)
            {
                listBox.Items.Add(items[i]);
            }
            content.Children.Add(listBox);

            content.Children.Add(CreateLabel(startingLabel));
            var startInput = CreateTextBox();
            content.Children.Add(startInput);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);

            if (window.ShowDialog() != true)
                return null;

            return new SheetRenumberInputsWithListResult
            {
                SelectedIndices = GetSelectedIndices(listBox),
                StartingNumber = startInput.Text ?? ""
            };
        }

        public static ViewnameReplaceInputsResult ViewnameReplaceInputs(
            string title,
            string findLabel,
            string replaceLabel,
            string prefixLabel,
            string suffixLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            content.Children.Add(CreateLabel(findLabel));
            var findInput = CreateTextBox();
            content.Children.Add(findInput);

            content.Children.Add(CreateLabel(replaceLabel));
            var replaceInput = CreateTextBox();
            content.Children.Add(replaceInput);

            content.Children.Add(CreateLabel(prefixLabel));
            var prefixInput = CreateTextBox();
            content.Children.Add(prefixInput);

            content.Children.Add(CreateLabel(suffixLabel));
            var suffixInput = CreateTextBox();
            content.Children.Add(suffixInput);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);

            if (window.ShowDialog() != true)
                return null;

            return new ViewnameReplaceInputsResult
            {
                Find = findInput.Text ?? "",
                Replace = replaceInput.Text ?? "",
                Prefix = prefixInput.Text ?? "",
                Suffix = suffixInput.Text ?? ""
            };
        }

        public static DuplicateSheetInputsResult DuplicateSheetInputs(
            IEnumerable items,
            string title,
            string prompt,
            string optionsLabel,
            string duplicateWithViewsLabel,
            string prefixLabel,
            string suffixLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            return DuplicateSheetInputs(
                AsStringList(items),
                title,
                prompt,
                optionsLabel,
                duplicateWithViewsLabel,
                prefixLabel,
                suffixLabel,
                okText,
                cancelText,
                width,
                height);
        }

        public static DuplicateSheetInputsResult DuplicateSheetInputs(
            IList<string> items,
            string title,
            string prompt,
            string optionsLabel,
            string duplicateWithViewsLabel,
            string prefixLabel,
            string suffixLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            if (items == null)
                items = Array.Empty<string>();

            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };
            content.Children.Add(CreatePrompt(prompt));

            var listBox = new ListBox
            {
                SelectionMode = SelectionMode.Extended,
                Height = Math.Max(240, height - 320),
                Margin = new Thickness(0, 0, 0, 8)
            };
            ScrollViewer.SetVerticalScrollBarVisibility(listBox, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(listBox, ScrollBarVisibility.Disabled);
            for (var i = 0; i < items.Count; i++)
            {
                listBox.Items.Add(items[i]);
            }
            content.Children.Add(listBox);

            var viewsCheckbox = new CheckBox { Content = duplicateWithViewsLabel, IsChecked = true, Margin = new Thickness(0, 8, 0, 2) };
            content.Children.Add(viewsCheckbox);

            content.Children.Add(CreateLabel(optionsLabel));
            var optionsCombo = CreateCombo(new[]
            {
                "Duplicate Views",
                "Duplicate Views w/Details",
                "Duplicate Views AsDependent"
            });
            content.Children.Add(optionsCombo);

            content.Children.Add(CreateLabel(prefixLabel));
            var prefixInput = CreateTextBox();
            content.Children.Add(prefixInput);

            content.Children.Add(CreateLabel(suffixLabel));
            var suffixInput = CreateTextBox();
            content.Children.Add(suffixInput);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);

            if (window.ShowDialog() != true)
                return null;

            return new DuplicateSheetInputsResult
            {
                SelectedIndices = GetSelectedIndices(listBox),
                DuplicateWithViews = viewsCheckbox.IsChecked == true,
                DuplicateOption = optionsCombo.SelectedIndex >= 0 ? optionsCombo.SelectedIndex : 0,
                Prefix = prefixInput.Text ?? "",
                Suffix = suffixInput.Text ?? ""
            };
        }

        public static ProjectUpgraderOptionsResult ProjectUpgraderOptions(
            string title,
            string description,
            string includeSubfoldersLabel,
            string okText,
            string cancelText,
            int width,
            int height,
            string initialFolder)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            if (!string.IsNullOrWhiteSpace(description))
                content.Children.Add(CreatePrompt(description));

            content.Children.Add(CreateLabel("Folder"));
            var folderPanel = new DockPanel { LastChildFill = false, Margin = new Thickness(0, 0, 0, 8) };
            var folderInput = new TextBox { MinWidth = 320, Text = initialFolder ?? "" };
            DockPanel.SetDock(folderInput, Dock.Left);
            folderPanel.Children.Add(folderInput);

            var browseButton = new Button { Content = "Browse", MinWidth = 90, Margin = new Thickness(8, 0, 0, 0) };
            browseButton.Click += (_, __) =>
            {
                var picked = BrowseForFolder(title ?? "Select Folder", folderInput.Text);
                if (!string.IsNullOrWhiteSpace(picked))
                    folderInput.Text = picked;
            };
            DockPanel.SetDock(browseButton, Dock.Right);
            folderPanel.Children.Add(browseButton);
            content.Children.Add(folderPanel);

            var subfolders = new CheckBox
            {
                Content = string.IsNullOrWhiteSpace(includeSubfoldersLabel) ? "Include subfolders" : includeSubfoldersLabel,
                IsChecked = true,
                Margin = new Thickness(0, 4, 0, 0)
            };
            content.Children.Add(subfolders);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);

            if (window.ShowDialog() != true)
                return null;

            return new ProjectUpgraderOptionsResult
            {
                Folder = folderInput.Text ?? "",
                IncludeSubfolders = subfolders.IsChecked == true
            };
        }

        public static FindReplaceResult FindReplaceDialog(
            string title,
            string findLabel,
            string replaceLabel,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            content.Children.Add(CreateLabel(findLabel));
            var findInput = CreateTextBox();
            content.Children.Add(findInput);

            content.Children.Add(CreateLabel(replaceLabel));
            var replaceInput = CreateTextBox();
            content.Children.Add(replaceInput);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            if (window.ShowDialog() != true)
                return null;

            return new FindReplaceResult
            {
                FindText = findInput.Text ?? "",
                ReplaceText = replaceInput.Text ?? ""
            };
        }

        public static RandomTreeSettingsResult RandomTreeSettings(
            string title,
            string rotationLabel,
            string sizeLabel,
            string percentLabel,
            double defaultPercent,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            var rotation = new CheckBox
            {
                Content = string.IsNullOrWhiteSpace(rotationLabel) ? "Random Rotation" : rotationLabel,
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 4)
            };
            content.Children.Add(rotation);

            var size = new CheckBox
            {
                Content = string.IsNullOrWhiteSpace(sizeLabel) ? "Random Size" : sizeLabel,
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 8)
            };
            content.Children.Add(size);

            content.Children.Add(CreateLabel(string.IsNullOrWhiteSpace(percentLabel) ? "Size variance (%)" : percentLabel));
            var percentInput = new TextBox { Text = defaultPercent.ToString("0") };
            content.Children.Add(percentInput);

            var buttons = CreateOkCancelButtons("OK", "Cancel");
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            if (window.ShowDialog() != true)
                return null;

            double percentValue;
            if (!double.TryParse(percentInput.Text, out percentValue))
                percentValue = defaultPercent;

            return new RandomTreeSettingsResult
            {
                RandomRotation = rotation.IsChecked == true,
                RandomSize = size.IsChecked == true,
                Percent = percentValue
            };
        }

        public static ParameterCopyInputsResult ParameterCopyInputs(
            IEnumerable paramNames,
            string title,
            string sourceDefault,
            string targetDefault,
            string findDefault,
            string replaceDefault,
            string prefixDefault,
            string suffixDefault,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };
            var names = AsStringList(paramNames);

            content.Children.Add(CreateLabel("Source Parameter"));
            var sourceCombo = CreateCombo(names);
            if (!string.IsNullOrWhiteSpace(sourceDefault))
                sourceCombo.Text = sourceDefault;
            content.Children.Add(sourceCombo);

            content.Children.Add(CreateLabel("Target Parameter"));
            var targetCombo = CreateCombo(names);
            if (!string.IsNullOrWhiteSpace(targetDefault))
                targetCombo.Text = targetDefault;
            content.Children.Add(targetCombo);

            content.Children.Add(CreateLabel("Find Text"));
            var findInput = CreateTextBox();
            findInput.Text = findDefault ?? "";
            content.Children.Add(findInput);

            content.Children.Add(CreateLabel("Replace Text"));
            var replaceInput = CreateTextBox();
            replaceInput.Text = replaceDefault ?? "";
            content.Children.Add(replaceInput);

            content.Children.Add(CreateLabel("Prefix"));
            var prefixInput = CreateTextBox();
            prefixInput.Text = prefixDefault ?? "";
            content.Children.Add(prefixInput);

            content.Children.Add(CreateLabel("Suffix"));
            var suffixInput = CreateTextBox();
            suffixInput.Text = suffixDefault ?? "";
            content.Children.Add(suffixInput);

            var buttons = CreateOkCancelButtons("OK", "Cancel");
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            if (window.ShowDialog() != true)
                return null;

            return new ParameterCopyInputsResult
            {
                SourceParam = sourceCombo.Text ?? "",
                TargetParam = targetCombo.Text ?? "",
                FindText = findInput.Text ?? "",
                ReplaceText = replaceInput.Text ?? "",
                Prefix = prefixInput.Text ?? "",
                Suffix = suffixInput.Text ?? ""
            };
        }

        public static DuplicateViewOptionsResult DuplicateViewOptions(
            IEnumerable optionLabels,
            IEnumerable optionValues,
            int defaultIndex,
            string title,
            string description,
            string prefixDefault,
            string suffixDefault,
            string okText,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            if (!string.IsNullOrWhiteSpace(description))
                content.Children.Add(CreatePrompt(description));

            content.Children.Add(CreateLabel("Prefix"));
            var prefixInput = CreateTextBox();
            prefixInput.Text = prefixDefault ?? "";
            content.Children.Add(prefixInput);

            content.Children.Add(CreateLabel("Suffix"));
            var suffixInput = CreateTextBox();
            suffixInput.Text = suffixDefault ?? "";
            content.Children.Add(suffixInput);

            content.Children.Add(CreateLabel("Duplicate Option"));
            var labels = AsStringList(optionLabels);
            var values = AsStringList(optionValues);
            var optionsPanel = new StackPanel { Margin = new Thickness(0, 4, 0, 8) };
            var radios = new List<RadioButton>();
            for (var i = 0; i < labels.Count; i++)
            {
                var rb = new RadioButton
                {
                    Content = labels[i],
                    Margin = new Thickness(0, 2, 0, 2),
                    IsChecked = i == defaultIndex
                };
                radios.Add(rb);
                optionsPanel.Children.Add(rb);
            }
            content.Children.Add(optionsPanel);

            var buttons = CreateOkCancelButtons(okText, "Cancel");
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            if (window.ShowDialog() != true)
                return null;

            var selectedIndex = radios.FindIndex(r => r.IsChecked == true);
            if (selectedIndex < 0)
                selectedIndex = 0;
            var selectedValue = selectedIndex < values.Count ? values[selectedIndex] : "";

            return new DuplicateViewOptionsResult
            {
                Prefix = prefixInput.Text ?? "",
                Suffix = suffixInput.Text ?? "",
                OptionValue = selectedValue
            };
        }

        public static MarketingViewOptionsResult MarketingViewOptions(
            IEnumerable sheetParams,
            IEnumerable templateNames,
            IEnumerable titleblockNames,
            IEnumerable keyplanTemplateNames,
            IEnumerable fillTypeNames,
            string title,
            string areaLabel,
            string doorLabel,
            bool keyplanEnabled,
            bool overwriteExisting,
            int templateIndex,
            int titleblockIndex,
            int keyplanTemplateIndex,
            int fillTypeIndex,
            string sheetNumberParam,
            string sheetNameParam,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };

            content.Children.Add(CreatePrompt("Area: " + (areaLabel ?? "(not selected)")));
            if (!string.IsNullOrWhiteSpace(doorLabel))
                content.Children.Add(CreatePrompt("Door: " + doorLabel));

            content.Children.Add(CreateLabel("Sheet number parameter:"));
            var sheetParamsList = AsStringList(sheetParams);
            var sheetNumberCombo = CreateCombo(sheetParamsList);
            if (!string.IsNullOrWhiteSpace(sheetNumberParam))
                sheetNumberCombo.SelectedItem = sheetNumberParam;
            content.Children.Add(sheetNumberCombo);

            content.Children.Add(CreateLabel("Sheet name parameter:"));
            var sheetNameCombo = CreateCombo(sheetParamsList);
            if (!string.IsNullOrWhiteSpace(sheetNameParam))
                sheetNameCombo.SelectedItem = sheetNameParam;
            content.Children.Add(sheetNameCombo);

            content.Children.Add(CreateLabel("Marketing view template:"));
            var templateCombo = CreateCombo(AsStringList(templateNames));
            if (templateIndex >= 0 && templateIndex < templateCombo.Items.Count)
                templateCombo.SelectedIndex = templateIndex;
            content.Children.Add(templateCombo);

            content.Children.Add(CreateLabel("Titleblock:"));
            var titleblockCombo = CreateCombo(AsStringList(titleblockNames));
            if (titleblockIndex >= 0 && titleblockIndex < titleblockCombo.Items.Count)
                titleblockCombo.SelectedIndex = titleblockIndex;
            content.Children.Add(titleblockCombo);

            var keyplanCheckbox = new CheckBox
            {
                Content = "Create Keyplan",
                IsChecked = keyplanEnabled,
                Margin = new Thickness(0, 6, 0, 4)
            };
            content.Children.Add(keyplanCheckbox);

            content.Children.Add(CreateLabel("Keyplan view template:"));
            var keyplanTemplateCombo = CreateCombo(AsStringList(keyplanTemplateNames));
            if (keyplanTemplateIndex >= 0 && keyplanTemplateIndex < keyplanTemplateCombo.Items.Count)
                keyplanTemplateCombo.SelectedIndex = keyplanTemplateIndex;
            content.Children.Add(keyplanTemplateCombo);

            content.Children.Add(CreateLabel("Keyplan filled region type:"));
            var fillCombo = CreateCombo(AsStringList(fillTypeNames));
            if (fillTypeIndex >= 0 && fillTypeIndex < fillCombo.Items.Count)
                fillCombo.SelectedIndex = fillTypeIndex;
            content.Children.Add(fillCombo);

            var overwriteCheckbox = new CheckBox
            {
                Content = "Overwrite existing sheet if sheet number exists",
                IsChecked = overwriteExisting,
                Margin = new Thickness(0, 6, 0, 4)
            };
            content.Children.Add(overwriteCheckbox);

            var buttons = CreateOkCancelButtons("Create", "Cancel");
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            if (window.ShowDialog() != true)
                return null;

            return new MarketingViewOptionsResult
            {
                SheetNumberParam = sheetNumberCombo.SelectedItem != null ? sheetNumberCombo.SelectedItem.ToString() : "",
                SheetNameParam = sheetNameCombo.SelectedItem != null ? sheetNameCombo.SelectedItem.ToString() : "",
                TemplateIndex = templateCombo.SelectedIndex,
                TitleblockIndex = titleblockCombo.SelectedIndex,
                KeyplanEnabled = keyplanCheckbox.IsChecked == true,
                KeyplanTemplateIndex = keyplanTemplateCombo.SelectedIndex,
                FillTypeIndex = fillCombo.SelectedIndex,
                OverwriteExisting = overwriteCheckbox.IsChecked == true
            };
        }

        public static KeyplanOptionsResult KeyplanOptions(
            IEnumerable templateNames,
            IEnumerable fillTypeNames,
            string title,
            string areaLabel,
            int templateIndex,
            int fillTypeIndex,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var content = new StackPanel { Margin = new Thickness(12) };
            content.Children.Add(CreatePrompt("Areas: " + (areaLabel ?? "(not selected)")));

            content.Children.Add(CreateLabel("Keyplan view template:"));
            var templateCombo = CreateCombo(AsStringList(templateNames));
            if (templateIndex >= 0 && templateIndex < templateCombo.Items.Count)
                templateCombo.SelectedIndex = templateIndex;
            content.Children.Add(templateCombo);

            content.Children.Add(CreateLabel("Keyplan filled region type:"));
            var fillCombo = CreateCombo(AsStringList(fillTypeNames));
            if (fillTypeIndex >= 0 && fillTypeIndex < fillCombo.Items.Count)
                fillCombo.SelectedIndex = fillTypeIndex;
            content.Children.Add(fillCombo);

            var buttons = CreateOkCancelButtons("Create", "Cancel");
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            content.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(content);
            if (window.ShowDialog() != true)
                return null;

            return new KeyplanOptionsResult
            {
                TemplateIndex = templateCombo.SelectedIndex,
                FillTypeIndex = fillCombo.SelectedIndex
            };
        }

        private static Window CreateMessageWindow(string message, string title, bool showCancel)
        {
            var window = CreateWindow(title, 520, 260);
            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var messageBlock = new TextBlock
            {
                Text = message ?? "",
                TextWrapping = TextWrapping.Wrap
            };
            var scroll = new ScrollViewer
            {
                Content = messageBlock,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Margin = new Thickness(0, 0, 0, 8),
                MinHeight = 80
            };
            Grid.SetRow(scroll, 0);
            grid.Children.Add(scroll);

            var buttons = CreateOkCancelButtons("OK", showCancel ? "Cancel" : null);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            if (showCancel)
                buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            Grid.SetRow(buttons.Panel, 1);
            grid.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(grid);
            return window;
        }

        public static LevelSetupInputsResult LevelSetupInputs(
            string title,
            string levelCountLabel,
            string height12Label,
            string height23Label,
            string typicalHeightLabel,
            string undergroundCountLabel,
            string heightP1ToL1Label,
            string typicalDepthLabel,
            string defaultLevelCount,
            string defaultHeight12,
            string defaultHeight23,
            string defaultTypicalHeight,
            string defaultUndergroundCount,
            string defaultHeightP1ToL1,
            string defaultTypicalDepth,
            string okText,
            string cancelText,
            int width,
            int height)
        {
            var window = CreateWindow(title, width, height);
            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var content = new StackPanel();

            content.Children.Add(CreateLabel(levelCountLabel));
            var levelCountInput = CreateTextBox();
            levelCountInput.Text = defaultLevelCount ?? "";
            content.Children.Add(levelCountInput);

            content.Children.Add(CreateLabel(height12Label));
            var height12Input = CreateTextBox();
            height12Input.Text = defaultHeight12 ?? "";
            content.Children.Add(height12Input);

            content.Children.Add(CreateLabel(height23Label));
            var height23Input = CreateTextBox();
            height23Input.Text = defaultHeight23 ?? "";
            content.Children.Add(height23Input);

            content.Children.Add(CreateLabel(typicalHeightLabel));
            var typicalInput = CreateTextBox();
            typicalInput.Text = defaultTypicalHeight ?? "";
            content.Children.Add(typicalInput);

            content.Children.Add(CreateLabel(undergroundCountLabel));
            var undergroundCountInput = CreateTextBox();
            undergroundCountInput.Text = defaultUndergroundCount ?? "";
            content.Children.Add(undergroundCountInput);

            content.Children.Add(CreateLabel(heightP1ToL1Label));
            var heightP1ToL1Input = CreateTextBox();
            heightP1ToL1Input.Text = defaultHeightP1ToL1 ?? "";
            content.Children.Add(heightP1ToL1Input);

            content.Children.Add(CreateLabel(typicalDepthLabel));
            var typicalDepthInput = CreateTextBox();
            typicalDepthInput.Text = defaultTypicalDepth ?? "";
            content.Children.Add(typicalDepthInput);

            var scroll = new ScrollViewer
            {
                Content = content,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Margin = new Thickness(0, 0, 0, 8)
            };
            Grid.SetRow(scroll, 0);
            grid.Children.Add(scroll);

            var buttons = CreateOkCancelButtons(okText, cancelText);
            buttons.Ok.Click += (_, __) => window.DialogResult = true;
            if (buttons.Cancel != null)
                buttons.Cancel.Click += (_, __) => window.DialogResult = false;
            Grid.SetRow(buttons.Panel, 1);
            grid.Children.Add(buttons.Panel);

            window.Content = WrapWithLogo(grid);
            if (window.ShowDialog() != true)
                return null;

            return new LevelSetupInputsResult
            {
                LevelCount = levelCountInput.Text ?? "",
                Height12 = height12Input.Text ?? "",
                Height23 = height23Input.Text ?? "",
                TypicalHeight = typicalInput.Text ?? "",
                UndergroundCount = undergroundCountInput.Text ?? "",
                HeightP1ToL1 = heightP1ToL1Input.Text ?? "",
                TypicalDepth = typicalDepthInput.Text ?? ""
            };
        }

        private static Window CreateWindow(string title, int width, int height)
        {
            EnsureApplication();
            var window = new Window
            {
                Title = title ?? "Dialog",
                Width = width > 0 ? width : 520,
                Height = height > 0 ? height : 320,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.CanResize,
                FontFamily = new FontFamily(DefaultFont),
                FontSize = DefaultFontSize,
                Background = Brushes.White,
                ShowInTaskbar = false
            };
            var owner = GetOwnerHandle();
            if (owner != IntPtr.Zero)
                new WindowInteropHelper(window) { Owner = owner };
            return window;
        }

        private static TextBlock CreatePrompt(string prompt)
        {
            return new TextBlock
            {
                Text = prompt ?? "",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            };
        }

        private static TextBlock CreateLabel(string text)
        {
            return new TextBlock
            {
                Text = text ?? "",
                Margin = new Thickness(0, 8, 0, 2)
            };
        }

        private static TextBox CreateTextBox()
        {
            return new TextBox
            {
                Margin = new Thickness(0, 0, 0, 8),
                MinWidth = 200
            };
        }

        private static ComboBox CreateCombo(IEnumerable<string> items)
        {
            var combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 8),
                MinWidth = 220
            };
            foreach (var item in items ?? Array.Empty<string>())
            {
                combo.Items.Add(item);
            }
            if (combo.Items.Count > 0)
                combo.SelectedIndex = 0;
            return combo;
        }

        private static int[] GetSelectedIndices(ListBox listBox)
        {
            if (listBox == null)
                return Array.Empty<int>();
            var selected = new List<int>();
            foreach (var item in listBox.SelectedItems)
            {
                var index = listBox.Items.IndexOf(item);
                if (index >= 0)
                    selected.Add(index);
            }
            selected.Sort();
            return selected.ToArray();
        }

        private static (StackPanel Panel, Button Ok, Button Cancel) CreateOkCancelButtons(string okText, string cancelText)
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };

            var okButton = new Button
            {
                Content = string.IsNullOrWhiteSpace(okText) ? "OK" : okText,
                MinWidth = 80,
                Margin = new Thickness(0, 0, 8, 0)
            };
            panel.Children.Add(okButton);

            Button cancelButton = null;
            if (!string.IsNullOrWhiteSpace(cancelText))
            {
                cancelButton = new Button
                {
                    Content = cancelText,
                    MinWidth = 80
                };
                panel.Children.Add(cancelButton);
            }

            return (panel, okButton, cancelButton);
        }

        private static UIElement WrapWithLogo(UIElement content)
        {
            var grid = new Grid();
            var contentHost = new Border { Child = content, Margin = new Thickness(0, 0, 0, 12) };
            grid.Children.Add(contentHost);

            var logo = LoadLogoImage();
            if (logo != null)
            {
                logo.HorizontalAlignment = HorizontalAlignment.Left;
                logo.VerticalAlignment = VerticalAlignment.Bottom;
                logo.Margin = new Thickness(12, 0, 0, 12);
                logo.IsHitTestVisible = false;
                grid.Children.Add(logo);
            }

            return grid;
        }

        private static Image LoadLogoImage()
        {
            try
            {
                var assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (string.IsNullOrWhiteSpace(assemblyPath))
                    return null;
                var logoPath = Path.Combine(assemblyPath, "WWPtools-logo.png");
                if (!File.Exists(logoPath))
                    return null;

                var image = new Image
                {
                    Width = 56,
                    Height = 56,
                    VerticalAlignment = VerticalAlignment.Top
                };

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(logoPath);
                bitmap.EndInit();
                image.Source = bitmap;
                return image;
            }
            catch
            {
                return null;
            }
        }

        private static void EnsureApplication()
        {
            if (Application.Current != null)
                return;
            var app = new Application();
            app.ShutdownMode = ShutdownMode.OnExplicitShutdown;
        }

        private static string BrowseForFolder(string title, string initialFolder)
        {
            IFileDialog dialog = null;
            try
            {
                dialog = (IFileDialog)new FileOpenDialog();
                uint options;
                dialog.GetOptions(out options);
                options |= (uint)(FOS.FOS_PICKFOLDERS | FOS.FOS_FORCEFILESYSTEM | FOS.FOS_PATHMUSTEXIST);
                dialog.SetOptions(options);

                if (!string.IsNullOrWhiteSpace(title))
                    dialog.SetTitle(title);

                if (!string.IsNullOrWhiteSpace(initialFolder))
                {
                    IShellItem folderItem;
                    var hr = SHCreateItemFromParsingName(initialFolder, IntPtr.Zero, typeof(IShellItem).GUID, out folderItem);
                    if (hr == 0 && folderItem != null)
                        dialog.SetFolder(folderItem);
                }

                var owner = GetOwnerHandle();
                var result = dialog.Show(owner);
                if (result != 0)
                    return null;

                IShellItem item;
                dialog.GetResult(out item);
                if (item == null)
                    return null;

                IntPtr pathPtr;
                item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out pathPtr);
                var path = Marshal.PtrToStringUni(pathPtr);
                Marshal.FreeCoTaskMem(pathPtr);
                return path;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (dialog != null)
                    Marshal.ReleaseComObject(dialog);
            }
        }

        private static IntPtr GetOwnerHandle()
        {
            try
            {
                return Process.GetCurrentProcess().MainWindowHandle;
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        [ComImport]
        [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
        private class FileOpenDialog
        {
        }

        [ComImport]
        [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileDialog
        {
            [PreserveSig] int Show(IntPtr parent);
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
        }

        [ComImport]
        [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            out IShellItem ppv);

        [Flags]
        private enum FOS : uint
        {
            FOS_PICKFOLDERS = 0x00000020,
            FOS_FORCEFILESYSTEM = 0x00000040,
            FOS_PATHMUSTEXIST = 0x00000800
        }

        private enum SIGDN : uint
        {
            SIGDN_FILESYSPATH = 0x80058000
        }

        private static List<string> AsStringList(IEnumerable items)
        {
            var results = new List<string>();
            if (items == null)
                return results;
            foreach (var item in items)
            {
                results.Add(item != null ? item.ToString() ?? "" : "");
            }
            return results;
        }

        private static int[] AsIntArray(IEnumerable items)
        {
            if (items == null)
                return Array.Empty<int>();
            var results = new List<int>();
            foreach (var item in items)
            {
                if (item == null)
                    continue;
                try
                {
                    results.Add(Convert.ToInt32(item));
                }
                catch
                {
                    // ignore invalid entries
                }
            }
            return results.ToArray();
        }
    }

    public class SelectionWithModeResult
    {
        public int[] SelectedIndices { get; set; } = Array.Empty<int>();
        public int Mode { get; set; }
    }

    public class SheetRenumberInputsResult
    {
        public string Category { get; set; } = "";
        public string PrintSet { get; set; } = "";
        public string StartingNumber { get; set; } = "";
    }

    public class SheetRenumberInputsWithListResult
    {
        public int[] SelectedIndices { get; set; } = Array.Empty<int>();
        public string StartingNumber { get; set; } = "";
    }

    public class ViewnameReplaceInputsResult
    {
        public string Find { get; set; } = "";
        public string Replace { get; set; } = "";
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
    }

    public class DuplicateSheetInputsResult
    {
        public int[] SelectedIndices { get; set; } = Array.Empty<int>();
        public bool DuplicateWithViews { get; set; }
        public int DuplicateOption { get; set; }
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
    }

    public class ProjectUpgraderOptionsResult
    {
        public string Folder { get; set; } = "";
        public bool IncludeSubfolders { get; set; }
    }

    public class FindReplaceResult
    {
        public string FindText { get; set; } = "";
        public string ReplaceText { get; set; } = "";
    }

    public class RandomTreeSettingsResult
    {
        public bool RandomRotation { get; set; }
        public bool RandomSize { get; set; }
        public double Percent { get; set; }
    }

    public class ParameterCopyInputsResult
    {
        public string SourceParam { get; set; } = "";
        public string TargetParam { get; set; } = "";
        public string FindText { get; set; } = "";
        public string ReplaceText { get; set; } = "";
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
    }

    public class DuplicateViewOptionsResult
    {
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
        public string OptionValue { get; set; } = "";
    }

    public class MarketingViewOptionsResult
    {
        public string SheetNumberParam { get; set; } = "";
        public string SheetNameParam { get; set; } = "";
        public int TemplateIndex { get; set; }
        public int TitleblockIndex { get; set; }
        public bool KeyplanEnabled { get; set; }
        public int KeyplanTemplateIndex { get; set; }
        public int FillTypeIndex { get; set; }
        public bool OverwriteExisting { get; set; }
    }

    public class KeyplanOptionsResult
    {
        public int TemplateIndex { get; set; }
        public int FillTypeIndex { get; set; }
    }

    public class LevelSetupInputsResult
    {
        public string LevelCount { get; set; } = "";
        public string Height12 { get; set; } = "";
        public string Height23 { get; set; } = "";
        public string TypicalHeight { get; set; } = "";
        public string UndergroundCount { get; set; } = "";
        public string HeightP1ToL1 { get; set; } = "";
        public string TypicalDepth { get; set; } = "";
    }
}
