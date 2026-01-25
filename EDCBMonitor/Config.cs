using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml.Serialization;
using System.Windows;

namespace EDCBMonitor
{
    public class ConfigData : INotifyPropertyChanged
    {
        private string _edcbInstallPath = "";
        private bool _topmost = false;
        private double _opacity = 1.0;
        private bool _hideDisabled = false;
        
        private string _backgroundColor = "#1E1E1E";
        private string _scrollBarColor = "#393939";
        private string _foregroundColor = "#EEEEEE";
        private string _recColor = "#FF5555";
        public string ReserveErrorColor { get; set; } = "#C85D5A";
        private string _disabledColor = "#777777";
        private string _columnBorderColor = "#808080";
        private string _footerColor = "#888888";
        private bool _recBold = true;

        private string _fontFamily = "Yu Gothic UI";
        private double _fontSize = 12.0;
        private double _itemPadding = 0.0;
        private double _headerFontSize = 12.0;
        private double _footerFontSize = 11.0;
        private double _listMarginLeft = 10;
        private double _listMarginTop = 10;
        private double _listMarginRight = 0;
        private double _listMarginBottom = 0;

        private bool _enableTitleRemove = true;
        private string _titleRemovePattern = @"[\[\(【](SS|無料|[字デ解二無多映])[\]\)】]";

        private bool _showHeader = true;
        private bool _showListHeader = true;
        private bool _showFooter = true;

        public double Top { get; set; } = -10000;
        public double Left { get; set; } = -10000;
        public double Width { get; set; } = 715;
        public double Height { get; set; } = 500;
        
        public List<string> ColumnHeaderOrder { get; set; } = new List<string>();

        // EDCB全31項目フラグ
        public bool ShowColID { get; set; } = false;
        public bool ShowColStatus { get; set; } = false;
        public bool ShowColDateTime { get; set; } = true;
        public bool ShowColDuration { get; set; } = true;
        public bool ShowColNetwork { get; set; } = false;
        public bool ShowColServiceName { get; set; } = true;
        public bool ShowColTitle { get; set; } = true;
        public bool ShowColDesc { get; set; } = false;
        public bool ShowColGenre { get; set; } = false;
        public bool ShowColExtraInfo { get; set; } = false;
        public bool ShowColEnabled { get; set; } = false;
        public bool ShowColProgramType { get; set; } = false;
        public bool ShowColComment { get; set; } = false;
        public bool ShowColError { get; set; } = false;
        public bool ShowColRecFileName { get; set; } = false;
        public bool ShowColRecFileNameList { get; set; } = false;
        public bool ShowColTuner { get; set; } = false;
        public bool ShowColEstSize { get; set; } = false;
        public bool ShowColPreset { get; set; } = false;
        public bool ShowColRecMode { get; set; } = false;
        public bool ShowColPriority { get; set; } = false;
        public bool ShowColTuijyuu { get; set; } = false;
        public bool ShowColPittari { get; set; } = false;
        public bool ShowColTunerForce { get; set; } = false;
        public bool ShowColRecEndMode { get; set; } = false;
        public bool ShowColReboot { get; set; } = false;
        public bool ShowColBat { get; set; } = false;
        public bool ShowColRecTag { get; set; } = false;
        public bool ShowColRecFolder { get; set; } = false;
        public bool ShowColStartMargin { get; set; } = false;
        public bool ShowColEndMargin { get; set; } = false;

        // EDCB全31項目幅
        public double WidthColID { get; set; } = 50;
        public double WidthColStatus { get; set; } = 60;
        public double WidthColDateTime { get; set; } = 130;
        public double WidthColDuration { get; set; } = 31;
        public double WidthColNetwork { get; set; } = 70;
        public double WidthColServiceName { get; set; } = 70;
        public double WidthColTitle { get; set; } = 450;
        public double WidthColDesc { get; set; } = 150;
        public double WidthColGenre { get; set; } = 80;
        public double WidthColExtraInfo { get; set; } = 100;
        public double WidthColEnabled { get; set; } = 60;
        public double WidthColProgramType { get; set; } = 80;
        public double WidthColComment { get; set; } = 150;
        public double WidthColError { get; set; } = 100;
        public double WidthColRecFileName { get; set; } = 150;
        public double WidthColRecFileNameList { get; set; } = 150;
        public double WidthColTuner { get; set; } = 100;
        public double WidthColEstSize { get; set; } = 70;
        public double WidthColPreset { get; set; } = 70;
        public double WidthColRecMode { get; set; } = 70;
        public double WidthColPriority { get; set; } = 50;
        public double WidthColTuijyuu { get; set; } = 50;
        public double WidthColPittari { get; set; } = 50;
        public double WidthColTunerForce { get; set; } = 80;
        public double WidthColRecEndMode { get; set; } = 80;
        public double WidthColReboot { get; set; } = 50;
        public double WidthColBat { get; set; } = 100;
        public double WidthColRecTag { get; set; } = 100;
        public double WidthColRecFolder { get; set; } = 100;
        public double WidthColStartMargin { get; set; } = 80;
        public double WidthColEndMargin { get; set; } = 80;

        public string EdcbInstallPath { get => _edcbInstallPath; set => SetProperty(ref _edcbInstallPath, value); }
        public bool Topmost { get => _topmost; set => SetProperty(ref _topmost, value); }
        public double Opacity { get => _opacity; set => SetProperty(ref _opacity, value); }
        public bool HideDisabled { get => _hideDisabled; set => SetProperty(ref _hideDisabled, value); }
        public string BackgroundColor { get => _backgroundColor; set => SetProperty(ref _backgroundColor, value); }
        public string ScrollBarColor { get => _scrollBarColor; set => SetProperty(ref _scrollBarColor, value); }
        public string ForegroundColor { get => _foregroundColor; set => SetProperty(ref _foregroundColor, value); }
        public string RecColor { get => _recColor; set => SetProperty(ref _recColor, value); }
        public string DisabledColor { get => _disabledColor; set => SetProperty(ref _disabledColor, value); }
        public bool RecBold { get => _recBold; set => SetProperty(ref _recBold, value); }
        public string FontFamily { get => _fontFamily; set => SetProperty(ref _fontFamily, value); }
        public double FontSize { get => _fontSize; set => SetProperty(ref _fontSize, value); }
        public double HeaderFontSize { get => _headerFontSize; set => SetProperty(ref _headerFontSize, value); }
        public double FooterFontSize { get => _footerFontSize; set => SetProperty(ref _footerFontSize, value); }
        public double ItemPadding { get => _itemPadding; set => SetProperty(ref _itemPadding, value); }
        
        public double ListMarginLeft { get => _listMarginLeft; set { if (SetProperty(ref _listMarginLeft, value)) OnPropertyChanged(nameof(ListMargin)); } }
        public double ListMarginTop { get => _listMarginTop; set { if (SetProperty(ref _listMarginTop, value)) OnPropertyChanged(nameof(ListMargin)); } }
        public double ListMarginRight { get => _listMarginRight; set { if (SetProperty(ref _listMarginRight, value)) OnPropertyChanged(nameof(ListMargin)); } }
        public double ListMarginBottom { get => _listMarginBottom; set { if (SetProperty(ref _listMarginBottom, value)) OnPropertyChanged(nameof(ListMargin)); } }
        
        [XmlIgnore] public Thickness ListMargin => new Thickness(_listMarginLeft, _listMarginTop, _listMarginRight, _listMarginBottom);

        public bool EnableTitleRemove { get => _enableTitleRemove; set => SetProperty(ref _enableTitleRemove, value); }
        public string TitleRemovePattern { get => _titleRemovePattern; set => SetProperty(ref _titleRemovePattern, value); }
        public bool ShowHeader { get => _showHeader; set => SetProperty(ref _showHeader, value); }
        public bool ShowListHeader { get => _showListHeader; set => SetProperty(ref _showListHeader, value); }
        public bool ShowFooter { get => _showFooter; set => SetProperty(ref _showFooter, value); }
        public string ColumnBorderColor { get => _columnBorderColor; set => SetProperty(ref _columnBorderColor, value); }
        public string FooterColor { get => _footerColor; set => SetProperty(ref _footerColor, value); }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        protected bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
        {
            if (object.Equals(storage, value)) return false;
            storage = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }

    public static class Config
    {
        public static ConfigData Data { get; set; } = new ConfigData();
        private static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.xml");

        public static void Load()
        {
            try { if (File.Exists(ConfigPath)) { var serializer = new XmlSerializer(typeof(ConfigData)); using var sr = new StreamReader(ConfigPath, new UTF8Encoding(false)); if (serializer.Deserialize(sr) is ConfigData loaded) Data = loaded; } } catch { }
        }

        public static void Save()
        {
            try
            {
                string path = ConfigPath;
                string tempPath = path + ".tmp";

                // 一時ファイル(.tmp)に書き込む
                var serializer = new XmlSerializer(typeof(ConfigData));
                
                // FileStreamを使って確実にディスクへ書き出す
                using (var fs = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var sw = new StreamWriter(fs, new UTF8Encoding(false)))
                {
                    serializer.Serialize(sw, Data);
                    sw.Flush();
                    fs.Flush(true); // ディスクバッファへフラッシュ
                }

                // 一時ファイルの書き込みに成功したら元のファイルを削除して差し替える
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                File.Move(tempPath, path);
            }
            catch (Exception ex)
            {
                // エラー時はログに残す（もしLoggerがなければ catch { } だけでも可）
                try { Logger.Write("設定保存エラー: " + ex.Message); } catch { }
            }
        }
    }
}