using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace WWPTools.WpfUI.Views
{
    public partial class SelectItemsWithModeWindow : Window
    {
        public SelectItemsWithModeWindow()
        {
            InitializeComponent();
            DialogStyles.ApplyPrimaryButtonStyle(OkButton);
            DialogStyles.ApplyPrimaryButtonStyle(CancelButton);
            OkButton.Click += (_, __) => DialogResult = true;
            CancelButton.Click += (_, __) => DialogResult = false;
            LoadLogo();
        }

        public void Initialize(
            IList<string> items,
            string prompt,
            string modeLabelA,
            string modeLabelB,
            int defaultMode,
            ISet<int> prechecked)
        {
            PromptText.Text = prompt ?? "";
            ModeA.Content = string.IsNullOrWhiteSpace(modeLabelA) ? "Option A" : modeLabelA;
            ModeB.Content = string.IsNullOrWhiteSpace(modeLabelB) ? "Option B" : modeLabelB;
            ModeA.IsChecked = defaultMode == 0;
            ModeB.IsChecked = defaultMode == 1;

            ItemsList.Items.Clear();
            if (items != null)
            {
                foreach (var item in items)
                {
                    ItemsList.Items.Add(item);
                }
            }

            if (prechecked != null)
            {
                foreach (var index in prechecked)
                {
                    if (index >= 0 && index < ItemsList.Items.Count)
                        ItemsList.SelectedItems.Add(ItemsList.Items[index]);
                }
            }
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
            return ModeA.IsChecked == true ? 0 : 1;
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
