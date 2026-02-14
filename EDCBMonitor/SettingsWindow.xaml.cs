using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using System.Windows;
using System.Windows.Media;
using System.Windows.Markup;
using WinForms = System.Windows.Forms; 

namespace EDCBMonitor
{
    public partial class SettingsWindow : Window
    {
        private bool _isLoaded = false;
        private string _backupConfigXml = "";

        public SettingsWindow()
        {
            InitializeComponent();
            ChkEnableTitleRemove.Click += (s, e) => UpdatePreview(true);
            ChkHideDisabled.Click += (s, e) => UpdatePreview(true);
            BackupConfig();
            LoadValues();
            _isLoaded = true;
        }

        private void BackupConfig()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(ConfigData));
                using var ms = new MemoryStream();
                serializer.Serialize(ms, Config.Data);
                _backupConfigXml = Encoding.UTF8.GetString(ms.ToArray());
            } catch (Exception ex) { Logger.Write("BackupConfig Error: " + ex.Message); }
        }

        private void RestoreConfig()
        {
            try
            {
                if (!string.IsNullOrEmpty(_backupConfigXml))
                {
                    var serializer = new XmlSerializer(typeof(ConfigData));
                    using var ms = new MemoryStream(Encoding.UTF8.GetBytes(_backupConfigXml));
                    if (serializer.Deserialize(ms) is ConfigData restored) Config.Data = restored;
                }
            } catch (Exception ex) { Logger.Write("RestoreConfig Error: " + ex.Message); }
        }

        private void LoadValues()
        {
            TxtPath.Text = Config.Data.EdcbInstallPath;
            ChkTopmost.IsChecked = Config.Data.Topmost;
            ChkHideDisabled.IsChecked = Config.Data.HideDisabled;
            SldOpacity.Value = Config.Data.Opacity;
            
            TxtBgColor.Text = Config.Data.BackgroundColor;
            TxtScrollBarColor.Text = Config.Data.ScrollBarColor;
            TxtFgColor.Text = Config.Data.ForegroundColor;
            TxtRecColor.Text = Config.Data.RecColor;
            TxtDisabledColor.Text = Config.Data.DisabledColor;
            ChkRecBold.IsChecked = Config.Data.RecBold;
            TxtColumnBorderColor.Text = Config.Data.ColumnBorderColor;
            TxtFooterColor.Text = Config.Data.FooterColor;
            TxtMainBorderColor.Text = Config.Data.MainBorderColor;
            TxtErrorColor.Text = Config.Data.ReserveErrorColor;
            TxtProgressBarColor.Text = Config.Data.ProgressBarColor;
            TxtProgressBarBackColor.Text = Config.Data.ProgressBarBackColor;
            TxtFontSize.Text = Config.Data.FontSize.ToString();
            TxtHeaderFontSize.Text = Config.Data.HeaderFontSize.ToString();
            TxtFooterFontSize.Text = Config.Data.FooterFontSize.ToString();
            SldItemPadding.Value = Config.Data.ItemPadding;
            
            SldMarginLeft.Value = Config.Data.ListMarginLeft;
            SldMarginTop.Value = Config.Data.ListMarginTop;
            SldMarginRight.Value = Config.Data.ListMarginRight;
            SldMarginBottom.Value = Config.Data.ListMarginBottom;

            TxtToolTipFontSize.Text = Config.Data.ToolTipFontSize.ToString();
            TxtToolTipBgColor.Text = Config.Data.ToolTipBackColor;
            TxtToolTipFgColor.Text = Config.Data.ToolTipForeColor;
            TxtToolTipBorderColor.Text = Config.Data.ToolTipBorderColor;
            ChkShowToolTip.IsChecked = Config.Data.ShowToolTip;
            TxtToolTipWidth.Text = Config.Data.ToolTipWidth.ToString();

            TxtScrollV.Text = Config.Data.ScrollAmountVertical.ToString();
            TxtScrollH.Text = Config.Data.ScrollAmountHorizontal.ToString();

            ChkEnableTitleRemove.IsChecked = Config.Data.EnableTitleRemove;
            TxtTitleRemovePattern.Text = Config.Data.TitleRemovePattern;

            CmbFont.Items.Clear();
            try
            {
                var langJa = XmlLanguage.GetLanguage("ja-jp");
                foreach (var font in Fonts.SystemFontFamilies)
                {
                    if (font.FamilyNames.ContainsKey(langJa)) CmbFont.Items.Add(font.FamilyNames[langJa]);
                    else CmbFont.Items.Add(font.Source);
                }
            }
            catch (Exception ex) { Logger.Write("Font Load Error: " + ex.Message); }
            
            CmbFont.Text = Config.Data.FontFamily;

            ChkShowHeader.IsChecked = Config.Data.ShowHeader;
            ChkShowListHeader.IsChecked = Config.Data.ShowListHeader;
            ChkShowFooter.IsChecked = Config.Data.ShowFooter;
            
            ChkOmitYear.IsChecked = Config.Data.OmitYear;
            ChkOmitMonth.IsChecked = Config.Data.OmitMonth;
            ChkOmitEndTime.IsChecked = Config.Data.OmitEndTime;
            ChkOmitProgress.IsChecked = Config.Data.OmitProgress;
            ChkShowRemainingTime.IsChecked = Config.Data.ShowRemainingTime;

            ChkColStatus.IsChecked = Config.Data.ShowColStatus;
            ChkColDateTime.IsChecked = Config.Data.ShowColDateTime;
            ChkColDuration.IsChecked = Config.Data.ShowColDuration;
            ChkColNetwork.IsChecked = Config.Data.ShowColNetwork;
            ChkColServiceName.IsChecked = Config.Data.ShowColServiceName;
            ChkColTitle.IsChecked = Config.Data.ShowColTitle;
            ChkColDesc.IsChecked = Config.Data.ShowColDesc;
            ChkColGenre.IsChecked = Config.Data.ShowColGenre;
            ChkColExtraInfo.IsChecked = Config.Data.ShowColExtraInfo;
            ChkColEnabled.IsChecked = Config.Data.ShowColEnabled;
            ChkColProgramType.IsChecked = Config.Data.ShowColProgramType;
            ChkColComment.IsChecked = Config.Data.ShowColComment;
            ChkColError.IsChecked = Config.Data.ShowColError;
            ChkColRecFileName.IsChecked = Config.Data.ShowColRecFileName;
            ChkColRecFileNameList.IsChecked = Config.Data.ShowColRecFileNameList;
            ChkColTuner.IsChecked = Config.Data.ShowColTuner;
            ChkColEstSize.IsChecked = Config.Data.ShowColEstSize;
            ChkColPreset.IsChecked = Config.Data.ShowColPreset;
            ChkColRecMode.IsChecked = Config.Data.ShowColRecMode;
            ChkColPriority.IsChecked = Config.Data.ShowColPriority;
            ChkColTuijyuu.IsChecked = Config.Data.ShowColTuijyuu;
            ChkColPittari.IsChecked = Config.Data.ShowColPittari;
            ChkColTunerForce.IsChecked = Config.Data.ShowColTunerForce;
            ChkColRecEndMode.IsChecked = Config.Data.ShowColRecEndMode;
            ChkColReboot.IsChecked = Config.Data.ShowColReboot;
            ChkColBat.IsChecked = Config.Data.ShowColBat;
            ChkColRecTag.IsChecked = Config.Data.ShowColRecTag;
            ChkColRecFolder.IsChecked = Config.Data.ShowColRecFolder;
            ChkColStartMargin.IsChecked = Config.Data.ShowColStartMargin;
            ChkColEndMargin.IsChecked = Config.Data.ShowColEndMargin;
            ChkColID.IsChecked = Config.Data.ShowColID;
            TxtBtnColor.Text = Config.Data.FooterBtnColor;
            TxtTvTestPath.Text = Config.Data.TvTestPath;
            TxtTvTestCmd.Text = Config.Data.TvTestCmd;
        }

        private void ApplyUiToConfig()
        {
            Config.Data.EdcbInstallPath = TxtPath.Text;
            Config.Data.Topmost = ChkTopmost.IsChecked == true;
            Config.Data.HideDisabled = ChkHideDisabled.IsChecked == true;
            Config.Data.Opacity = SldOpacity.Value;
            
            Config.Data.BackgroundColor = TxtBgColor.Text;
            Config.Data.ScrollBarColor = TxtScrollBarColor.Text;
            Config.Data.ForegroundColor = TxtFgColor.Text;
            Config.Data.RecColor = TxtRecColor.Text;
            Config.Data.DisabledColor = TxtDisabledColor.Text;
            Config.Data.RecBold = ChkRecBold.IsChecked == true;
            Config.Data.ColumnBorderColor = TxtColumnBorderColor.Text;
            Config.Data.FooterColor = TxtFooterColor.Text;
            Config.Data.MainBorderColor = TxtMainBorderColor.Text;
            Config.Data.ReserveErrorColor = TxtErrorColor.Text;
            Config.Data.ProgressBarColor = TxtProgressBarColor.Text;
            Config.Data.ProgressBarBackColor = TxtProgressBarBackColor.Text;
            Config.Data.EnableTitleRemove = ChkEnableTitleRemove.IsChecked == true;
            Config.Data.TitleRemovePattern = TxtTitleRemovePattern.Text;
            
            if (CmbFont.SelectedItem != null) Config.Data.FontFamily = CmbFont.SelectedItem.ToString() ?? "";
            else Config.Data.FontFamily = CmbFont.Text ?? "";

            if (double.TryParse(TxtFontSize.Text, out double fs)) Config.Data.FontSize = fs;
            if (double.TryParse(TxtHeaderFontSize.Text, out double hfs)) Config.Data.HeaderFontSize = hfs;
            if (double.TryParse(TxtFooterFontSize.Text, out double ffs)) Config.Data.FooterFontSize = ffs;
            Config.Data.ItemPadding = SldItemPadding.Value;
            
            Config.Data.ListMarginLeft = SldMarginLeft.Value;
            Config.Data.ListMarginTop = SldMarginTop.Value;
            Config.Data.ListMarginRight = SldMarginRight.Value;
            Config.Data.ListMarginBottom = SldMarginBottom.Value;

            if (double.TryParse(TxtToolTipFontSize.Text, out double tfs)) Config.Data.ToolTipFontSize = tfs;
            Config.Data.ToolTipBackColor = TxtToolTipBgColor.Text;
            Config.Data.ToolTipForeColor = TxtToolTipFgColor.Text;
            Config.Data.ToolTipBorderColor = TxtToolTipBorderColor.Text;
            Config.Data.ShowToolTip = ChkShowToolTip.IsChecked == true;
            if (double.TryParse(TxtToolTipWidth.Text, out double ttw)) Config.Data.ToolTipWidth = ttw;

            if (int.TryParse(TxtScrollV.Text, out int sv)) Config.Data.ScrollAmountVertical = sv;
            if (int.TryParse(TxtScrollH.Text, out int sh)) Config.Data.ScrollAmountHorizontal = sh;

            Config.Data.ShowHeader = ChkShowHeader.IsChecked == true;
            Config.Data.ShowListHeader = ChkShowListHeader.IsChecked == true;
            Config.Data.ShowFooter = ChkShowFooter.IsChecked == true;
            
            Config.Data.OmitYear = ChkOmitYear.IsChecked == true;
            Config.Data.OmitMonth = ChkOmitMonth.IsChecked == true;
            Config.Data.OmitEndTime = ChkOmitEndTime.IsChecked == true;
            Config.Data.OmitProgress = ChkOmitProgress.IsChecked == true;
            Config.Data.ShowRemainingTime = ChkShowRemainingTime.IsChecked == true;

            Config.Data.ShowColStatus = ChkColStatus.IsChecked == true;
            Config.Data.ShowColDateTime = ChkColDateTime.IsChecked == true;
            Config.Data.ShowColDuration = ChkColDuration.IsChecked == true;
            Config.Data.ShowColNetwork = ChkColNetwork.IsChecked == true;
            Config.Data.ShowColServiceName = ChkColServiceName.IsChecked == true;
            Config.Data.ShowColTitle = ChkColTitle.IsChecked == true;
            Config.Data.ShowColDesc = ChkColDesc.IsChecked == true;
            Config.Data.ShowColGenre = ChkColGenre.IsChecked == true;
            Config.Data.ShowColExtraInfo = ChkColExtraInfo.IsChecked == true;
            Config.Data.ShowColEnabled = ChkColEnabled.IsChecked == true;
            Config.Data.ShowColProgramType = ChkColProgramType.IsChecked == true;
            Config.Data.ShowColComment = ChkColComment.IsChecked == true;
            Config.Data.ShowColError = ChkColError.IsChecked == true;
            Config.Data.ShowColRecFileName = ChkColRecFileName.IsChecked == true;
            Config.Data.ShowColRecFileNameList = ChkColRecFileNameList.IsChecked == true;
            Config.Data.ShowColTuner = ChkColTuner.IsChecked == true;
            Config.Data.ShowColEstSize = ChkColEstSize.IsChecked == true;
            Config.Data.ShowColPreset = ChkColPreset.IsChecked == true;
            Config.Data.ShowColRecMode = ChkColRecMode.IsChecked == true;
            Config.Data.ShowColPriority = ChkColPriority.IsChecked == true;
            Config.Data.ShowColTuijyuu = ChkColTuijyuu.IsChecked == true;
            Config.Data.ShowColPittari = ChkColPittari.IsChecked == true;
            Config.Data.ShowColTunerForce = ChkColTunerForce.IsChecked == true;
            Config.Data.ShowColRecEndMode = ChkColRecEndMode.IsChecked == true;
            Config.Data.ShowColReboot = ChkColReboot.IsChecked == true;
            Config.Data.ShowColBat = ChkColBat.IsChecked == true;
            Config.Data.ShowColRecTag = ChkColRecTag.IsChecked == true;
            Config.Data.ShowColRecFolder = ChkColRecFolder.IsChecked == true;
            Config.Data.ShowColStartMargin = ChkColStartMargin.IsChecked == true;
            Config.Data.ShowColEndMargin = ChkColEndMargin.IsChecked == true;
            Config.Data.ShowColID = ChkColID.IsChecked == true;
            Config.Data.FooterBtnColor = TxtBtnColor.Text;
            Config.Data.TvTestPath = TxtTvTestPath.Text;
            Config.Data.TvTestCmd = TxtTvTestCmd.Text;
        }

        private void UpdatePreview(bool needReload)
        {
            if (!_isLoaded) return;
            ApplyUiToConfig();
            if (this.Owner is MainWindow mw)
            {
                mw.ApplySettings();
                
                if (needReload)
                {
                    _ = mw.RefreshDataAsync();
                }
            }
        }

        private void Interact_Changed(object sender, RoutedEventArgs e) => UpdatePreview(false);
        private void DataSetting_Changed(object sender, RoutedEventArgs e) => UpdatePreview(true);
        private void SldOpacity_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) => UpdatePreview(false);
        private void SldItemPadding_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) => UpdatePreview(false);
        private void SldMargin_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) => UpdatePreview(false);

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            ApplyUiToConfig();
            Config.Save();
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => this.Close();
        
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            if (this.DialogResult != true)
            {
                RestoreConfig();
                if (this.Owner is MainWindow mw)
                {
                    mw.ApplySettings();
                    _ = mw.RefreshDataAsync();
                }
            }
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed) this.DragMove();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Reserve.txt|Reserve.txt|すべてのファイル|*.*" };
            if (dlg.ShowDialog() == true)
            {
                TxtPath.Text = dlg.FileName;
                UpdatePreview(true);
            }
        }

        private void BtnThemeDark_Click(object sender, RoutedEventArgs e)
        {
            TxtBgColor.Text = "#1E1E1E";
            TxtScrollBarColor.Text = "#393939";
            TxtFgColor.Text = "#EEEEEE";
            TxtDisabledColor.Text = "#777777";
            TxtColumnBorderColor.Text = "#808080";
            TxtBtnColor.Text = "#555555"; 
            TxtMainBorderColor.Text = "#555555";

            TxtFooterColor.Text = "#888888";
            TxtToolTipBgColor.Text = "#F2F2F2";
            TxtToolTipFgColor.Text = "#000000";
            TxtToolTipBorderColor.Text = "#767676";
            TxtProgressBarBackColor.Text = "#A9A9A9";
            TxtRecColor.Text = "#FF5555";
            UpdatePreview(false);
        }

        private void BtnThemeLight_Click(object sender, RoutedEventArgs e)
        {
            TxtBgColor.Text = "#FFFFFF";
            TxtScrollBarColor.Text = "#CCCCCC";
            TxtFgColor.Text = "#2D2D2D";
            TxtDisabledColor.Text = "#C0C0C0";
            TxtColumnBorderColor.Text = "#8C8C8C"; 
            TxtBtnColor.Text = "#ABABAB";
            TxtMainBorderColor.Text = "#ABABAB";
            
            TxtFooterColor.Text = "#ACACAC";
            TxtToolTipBgColor.Text = "#FFFFE1"; 
            TxtToolTipFgColor.Text = "#000000";
            TxtToolTipBorderColor.Text = "#7A7A7A";
            TxtProgressBarBackColor.Text = "#E6E6E6";
            TxtRecColor.Text = "#FF0000";
            UpdatePreview(false);
        }

        private void PickColor(System.Windows.Controls.TextBox txt)
        {
            using var dlg = new WinForms.ColorDialog();
            if (dlg.ShowDialog() == WinForms.DialogResult.OK)
            {
                var c = dlg.Color;
                txt.Text = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                UpdatePreview(false);
            }
        }
        
        private void BtnBrowseTvTest_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*",
                Title = "TVTest.exe の選択"
            };
            if (dlg.ShowDialog() == true)
            {
                TxtTvTestPath.Text = dlg.FileName;
                UpdatePreview(false);
            }
        }

        private void BtnPickBg_Click(object sender, RoutedEventArgs e) => PickColor(TxtBgColor);
        private void BtnPickScrollBar_Click(object sender, RoutedEventArgs e) => PickColor(TxtScrollBarColor);
        private void BtnPickFg_Click(object sender, RoutedEventArgs e) => PickColor(TxtFgColor);
        private void BtnPickRec_Click(object sender, RoutedEventArgs e) => PickColor(TxtRecColor);
        private void BtnPickDisabled_Click(object sender, RoutedEventArgs e) => PickColor(TxtDisabledColor);
        private void BtnPickColumnBorder_Click(object sender, RoutedEventArgs e) => PickColor(TxtColumnBorderColor);
        private void BtnPickFooter_Click(object sender, RoutedEventArgs e) => PickColor(TxtFooterColor);
        private void BtnPickMainBorder_Click(object sender, RoutedEventArgs e) => PickColor(TxtMainBorderColor);
        private void BtnPickError_Click(object sender, RoutedEventArgs e) => PickColor(TxtErrorColor);
        private void BtnPickProgressBar_Click(object sender, RoutedEventArgs e) => PickColor(TxtProgressBarColor);
        private void BtnPickProgressBarBack_Click(object sender, RoutedEventArgs e) => PickColor(TxtProgressBarBackColor);

        private void BtnPickToolTipBg_Click(object sender, RoutedEventArgs e) => PickColor(TxtToolTipBgColor);
        private void BtnPickToolTipFg_Click(object sender, RoutedEventArgs e) => PickColor(TxtToolTipFgColor);
        private void BtnPickToolTipBorder_Click(object sender, RoutedEventArgs e) => PickColor(TxtToolTipBorderColor);
        private void BtnPickBtnColor_Click(object sender, RoutedEventArgs e) => PickColor(TxtBtnColor);
    }
}