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
    // カラムのすべての状態を管理
    public class ColumnState
    {
        public string Header { get; set; } = "";
        public bool IsVisible { get; set; } = false;
        public double Width { get; set; } = 100;

        public ColumnState() { }
        public ColumnState(string header, bool isVisible, double width)
        {
            Header = header; IsVisible = isVisible; Width = width;
        }
    }

    public class ConfigData : INotifyPropertyChanged
    {
        private string _edcbInstallPath = "";
        private bool _topmost = false;
        private bool _showTrayIcon = true;
        public bool ShowTrayIcon { get => _showTrayIcon; set => SetProperty(ref _showTrayIcon, value); }
        private double _opacity = 1.0;
        private bool _hideDisabled = false;
        
        // --- ミニモード設定 ---
        private bool _enableMiniMode = false;
        public bool EnableMiniMode { get => _enableMiniMode; set => SetProperty(ref _enableMiniMode, value); }
        private bool _miniRequireClickToExpand = true;
        public bool MiniRequireClickToExpand { get => _miniRequireClickToExpand; set => SetProperty(ref _miniRequireClickToExpand, value); }
        private bool _miniShowHeader = false;
        public bool MiniShowHeader { get => _miniShowHeader; set => SetProperty(ref _miniShowHeader, value); }
        private bool _miniShowListHeader = false;
        public bool MiniShowListHeader { get => _miniShowListHeader; set => SetProperty(ref _miniShowListHeader, value); }
        private bool _miniShowFooter = false;
        public bool MiniShowFooter { get => _miniShowFooter; set => SetProperty(ref _miniShowFooter, value); }
        private double _miniModeScaleX = 50.0;
        private double _miniModeScaleY = 50.0;
        private int _miniModeDirection = 1; // 右上固定
        private int _miniModeDelay = 500;   // 遅延時間(ms)
        private int _miniModeExpandDelay = 500;
        public int MiniModeExpandDelay { get => _miniModeExpandDelay; set => SetProperty(ref _miniModeExpandDelay, value); }

        // ミニモード用のリスト余白設定
        private double _miniListMarginLeft = 3;
        public double MiniListMarginLeft { get => _miniListMarginLeft; set { if (SetProperty(ref _miniListMarginLeft, value)) OnPropertyChanged(nameof(MiniListMargin)); } }
        private double _miniListMarginTop = 6;
        public double MiniListMarginTop { get => _miniListMarginTop; set { if (SetProperty(ref _miniListMarginTop, value)) OnPropertyChanged(nameof(MiniListMargin)); } }
        private double _miniListMarginRight = 0;
        public double MiniListMarginRight { get => _miniListMarginRight; set { if (SetProperty(ref _miniListMarginRight, value)) OnPropertyChanged(nameof(MiniListMargin)); } }
        private double _miniListMarginBottom = 6;
        public double MiniListMarginBottom { get => _miniListMarginBottom; set { if (SetProperty(ref _miniListMarginBottom, value)) OnPropertyChanged(nameof(MiniListMargin)); } }
        [XmlIgnore] public Thickness MiniListMargin => new Thickness(_miniListMarginLeft, _miniListMarginTop, _miniListMarginRight, _miniListMarginBottom);
        private string _backgroundColor = "#1E1E1E";
        private string _scrollBarColor = "#393939";
        private string _foregroundColor = "#EEEEEE";
        private string _recColor = "#FF5555";
        
        private string _reserveErrorColor = "#C85D5A";
        public string ReserveErrorColor { get => _reserveErrorColor; set => SetProperty(ref _reserveErrorColor, value); }
        
        private string _disabledColor = "#777777";
        private string _progressBarColor = "#0064C8";
        private string _progressBarBackColor = "#A9A9A9";
        private string _columnBorderColor = "#808080";
        private string _footerColor = "#888888";
        private string _mainBorderColor = "#555555";
        private bool _recBold = true;

        private string _selectedColor = "#50FFFFFF";
        public string SelectedColor { get => _selectedColor; set => SetProperty(ref _selectedColor, value); }
        private string _hoverColor = "#32FFFFFF";
        public string HoverColor { get => _hoverColor; set => SetProperty(ref _hoverColor, value); }

        private string _fontFamily = "Yu Gothic UI";
        private double _fontSize = 12.0;
        private double _itemPadding = 0.0;
        private double _headerFontSize = 12.0;
        private double _footerFontSize = 11.0;
        private double _listMarginLeft = 10;
        private double _listMarginTop = 10;
        private double _listMarginRight = 0;
        private double _listMarginBottom = 0;

        private double _toolTipFontSize = 12.0;
        private string _toolTipBackColor = "#F2F2F2";
        private string _toolTipForeColor = "#000000";
        private string _toolTipBorderColor = "#767676";
        private bool _showToolTip = true;
        private double _toolTipWidth = 500.0;

        private bool _enableTitleRemove = true;
        private string _titleRemovePattern = @"[\[\(【](SS|無料|[字デ解二無多映])[\]\)】]";
        
        private string _tvTestPath = "";
        public string TvTestPath { get => _tvTestPath; set => SetProperty(ref _tvTestPath, value); }

        private string _tvTestCmd = ""; 
        public string TvTestCmd { get => _tvTestCmd; set => SetProperty(ref _tvTestCmd, value); }

        private int _doubleClickAction = 0; // 0: EpgTimer, 1: Material WebUI, 2: 動作させない
        public int DoubleClickAction { get => _doubleClickAction; set => SetProperty(ref _doubleClickAction, value); }

        private string _materialWebUiUrl = "http://localhost:5510/EMWUI/";
        public string MaterialWebUiUrl { get => _materialWebUiUrl; set => SetProperty(ref _materialWebUiUrl, value); }

        private bool _showHeader = true;
        private bool _showListHeader = true;
        private bool _showFooter = true;

        public double Top { get; set; } = -10000;
        public double Left { get; set; } = -10000;
        public double Width { get; set; } = 660;
        public double Height { get; set; } = 478;
        public bool IsVerticalMaximized { get; set; } = false;
        public double RestoreTop { get; set; } = -10000;
        public double RestoreHeight { get; set; } = 500;

        // =========================================================
        // カラム管理システム (単一のマスターリスト)
        // =========================================================
        public List<ColumnState> Columns { get; set; } = new List<ColumnState>();

        public ConfigData()
        {
        }
        public void InitDefaults()
        {
            var defaultDefs = new List<ColumnState>
            {
                new ColumnState("状態", false, 60), new ColumnState("日時", true, 132), new ColumnState("長さ", true, 31),
                new ColumnState("ネットワーク", false, 70), new ColumnState("サービス名", true, 58), new ColumnState("番組名", true, 460),
                new ColumnState("番組内容", false, 150), new ColumnState("ジャンル", false, 80), new ColumnState("付属情報", false, 100),
                new ColumnState("有効", false, 60), new ColumnState("プログラム予約", false, 80), new ColumnState("予約状況", false, 150),
                new ColumnState("エラー状況", false, 100), new ColumnState("予定ファイル名", false, 150), new ColumnState("予定ファイル名リスト", false, 150),
                new ColumnState("使用予定チューナー", false, 100), new ColumnState("予想サイズ", false, 70), new ColumnState("プリセット", false, 70),
                new ColumnState("録画モード", false, 70), new ColumnState("優先度", false, 50), new ColumnState("追従", false, 50),
                new ColumnState("ぴったり", false, 50), new ColumnState("チューナー強制", false, 80), new ColumnState("録画後動作", false, 80),
                new ColumnState("復帰後再起動", false, 50), new ColumnState("録画後実行bat", false, 100), new ColumnState("録画タグ", false, 100),
                new ColumnState("録画フォルダ", false, 100), new ColumnState("開始", false, 80), new ColumnState("終了", false, 80),
                new ColumnState("ID", false, 50)
            };

            foreach (var def in defaultDefs)
            {
                if (!Columns.Any(c => c.Header == def.Header))
                {
                    Columns.Add(def);
                }
            }
        }

        public ColumnState GetColumn(string header)
        {
            var col = Columns.FirstOrDefault(c => c.Header == header);
            if (col == null)
            {
                col = new ColumnState(header, false, 100);
                Columns.Add(col);
            }
            return col;
        }

        // --- EDCB全31項目フラグ (Columnsマスターへのプロキシ) ---
        [System.Xml.Serialization.XmlIgnore] public bool ShowColStatus { get => GetColumn("状態").IsVisible; set { GetColumn("状態").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColDateTime { get => GetColumn("日時").IsVisible; set { GetColumn("日時").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColDuration { get => GetColumn("長さ").IsVisible; set { GetColumn("長さ").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColNetwork { get => GetColumn("ネットワーク").IsVisible; set { GetColumn("ネットワーク").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColServiceName { get => GetColumn("サービス名").IsVisible; set { GetColumn("サービス名").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColTitle { get => GetColumn("番組名").IsVisible; set { GetColumn("番組名").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColDesc { get => GetColumn("番組内容").IsVisible; set { GetColumn("番組内容").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColGenre { get => GetColumn("ジャンル").IsVisible; set { GetColumn("ジャンル").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColExtraInfo { get => GetColumn("付属情報").IsVisible; set { GetColumn("付属情報").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColEnabled { get => GetColumn("有効").IsVisible; set { GetColumn("有効").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColProgramType { get => GetColumn("プログラム予約").IsVisible; set { GetColumn("プログラム予約").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColComment { get => GetColumn("予約状況").IsVisible; set { GetColumn("予約状況").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColError { get => GetColumn("エラー状況").IsVisible; set { GetColumn("エラー状況").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColRecFileName { get => GetColumn("予定ファイル名").IsVisible; set { GetColumn("予定ファイル名").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColRecFileNameList { get => GetColumn("予定ファイル名リスト").IsVisible; set { GetColumn("予定ファイル名リスト").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColTuner { get => GetColumn("使用予定チューナー").IsVisible; set { GetColumn("使用予定チューナー").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColEstSize { get => GetColumn("予想サイズ").IsVisible; set { GetColumn("予想サイズ").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColPreset { get => GetColumn("プリセット").IsVisible; set { GetColumn("プリセット").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColRecMode { get => GetColumn("録画モード").IsVisible; set { GetColumn("録画モード").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColPriority { get => GetColumn("優先度").IsVisible; set { GetColumn("優先度").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColTuijyuu { get => GetColumn("追従").IsVisible; set { GetColumn("追従").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColPittari { get => GetColumn("ぴったり").IsVisible; set { GetColumn("ぴったり").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColTunerForce { get => GetColumn("チューナー強制").IsVisible; set { GetColumn("チューナー強制").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColRecEndMode { get => GetColumn("録画後動作").IsVisible; set { GetColumn("録画後動作").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColReboot { get => GetColumn("復帰後再起動").IsVisible; set { GetColumn("復帰後再起動").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColBat { get => GetColumn("録画後実行bat").IsVisible; set { GetColumn("録画後実行bat").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColRecTag { get => GetColumn("録画タグ").IsVisible; set { GetColumn("録画タグ").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColRecFolder { get => GetColumn("録画フォルダ").IsVisible; set { GetColumn("録画フォルダ").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColStartMargin { get => GetColumn("開始").IsVisible; set { GetColumn("開始").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColEndMargin { get => GetColumn("終了").IsVisible; set { GetColumn("終了").IsVisible = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public bool ShowColID { get => GetColumn("ID").IsVisible; set { GetColumn("ID").IsVisible = value; OnPropertyChanged(); } }

        // --- EDCB全31項目幅 (Columnsマスターへのプロキシ) ---
        [System.Xml.Serialization.XmlIgnore] public double WidthColID { get => GetColumn("ID").Width; set { GetColumn("ID").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColStatus { get => GetColumn("状態").Width; set { GetColumn("状態").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColDateTime { get => GetColumn("日時").Width; set { GetColumn("日時").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColDuration { get => GetColumn("長さ").Width; set { GetColumn("長さ").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColNetwork { get => GetColumn("ネットワーク").Width; set { GetColumn("ネットワーク").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColServiceName { get => GetColumn("サービス名").Width; set { GetColumn("サービス名").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColTitle { get => GetColumn("番組名").Width; set { GetColumn("番組名").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColDesc { get => GetColumn("番組内容").Width; set { GetColumn("番組内容").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColGenre { get => GetColumn("ジャンル").Width; set { GetColumn("ジャンル").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColExtraInfo { get => GetColumn("付属情報").Width; set { GetColumn("付属情報").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColEnabled { get => GetColumn("有効").Width; set { GetColumn("有効").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColProgramType { get => GetColumn("プログラム予約").Width; set { GetColumn("プログラム予約").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColComment { get => GetColumn("予約状況").Width; set { GetColumn("予約状況").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColError { get => GetColumn("エラー状況").Width; set { GetColumn("エラー状況").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColRecFileName { get => GetColumn("予定ファイル名").Width; set { GetColumn("予定ファイル名").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColRecFileNameList { get => GetColumn("予定ファイル名リスト").Width; set { GetColumn("予定ファイル名リスト").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColTuner { get => GetColumn("使用予定チューナー").Width; set { GetColumn("使用予定チューナー").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColEstSize { get => GetColumn("予想サイズ").Width; set { GetColumn("予想サイズ").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColPreset { get => GetColumn("プリセット").Width; set { GetColumn("プリセット").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColRecMode { get => GetColumn("録画モード").Width; set { GetColumn("録画モード").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColPriority { get => GetColumn("優先度").Width; set { GetColumn("優先度").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColTuijyuu { get => GetColumn("追従").Width; set { GetColumn("追従").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColPittari { get => GetColumn("ぴったり").Width; set { GetColumn("ぴったり").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColTunerForce { get => GetColumn("チューナー強制").Width; set { GetColumn("チューナー強制").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColRecEndMode { get => GetColumn("録画後動作").Width; set { GetColumn("録画後動作").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColReboot { get => GetColumn("復帰後再起動").Width; set { GetColumn("復帰後再起動").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColBat { get => GetColumn("録画後実行bat").Width; set { GetColumn("録画後実行bat").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColRecTag { get => GetColumn("録画タグ").Width; set { GetColumn("録画タグ").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColRecFolder { get => GetColumn("録画フォルダ").Width; set { GetColumn("録画フォルダ").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColStartMargin { get => GetColumn("開始").Width; set { GetColumn("開始").Width = value; OnPropertyChanged(); } }
        [System.Xml.Serialization.XmlIgnore] public double WidthColEndMargin { get => GetColumn("終了").Width; set { GetColumn("終了").Width = value; OnPropertyChanged(); } }

        // ※ 以下の設定はColumnsとは無関係なので XmlIgnore を付けない（そのまま保存する）
        public bool OmitProgress { get; set; } = false;
        public bool ShowRemainingTime { get; set; } = false;
        public string FooterBtnColor { get; set; } = "#555555";
        public bool OmitYear { get; set; } = true;
        public bool OmitMonth { get; set; } = false;
        public bool OmitEndTime { get; set; } = false;

        public string EdcbInstallPath { get => _edcbInstallPath; set => SetProperty(ref _edcbInstallPath, value); }
        public bool Topmost { get => _topmost; set => SetProperty(ref _topmost, value); }
        public double Opacity { get => _opacity; set => SetProperty(ref _opacity, value); }
        public bool HideDisabled { get => _hideDisabled; set => SetProperty(ref _hideDisabled, value); }
        public double MiniModeScaleX { get => _miniModeScaleX; set => SetProperty(ref _miniModeScaleX, value); }
        public double MiniModeScaleY { get => _miniModeScaleY; set => SetProperty(ref _miniModeScaleY, value); }
        public int MiniModeDirection { get => _miniModeDirection; set => SetProperty(ref _miniModeDirection, value); }
        public int MiniModeDelay { get => _miniModeDelay; set => SetProperty(ref _miniModeDelay, value); }
        public string BackgroundColor { get => _backgroundColor; set => SetProperty(ref _backgroundColor, value); }
        public string ScrollBarColor { get => _scrollBarColor; set => SetProperty(ref _scrollBarColor, value); }
        public string ForegroundColor { get => _foregroundColor; set => SetProperty(ref _foregroundColor, value); }
        public string RecColor { get => _recColor; set => SetProperty(ref _recColor, value); }
        public string DisabledColor { get => _disabledColor; set => SetProperty(ref _disabledColor, value); }
        public string ProgressBarColor { get => _progressBarColor; set => SetProperty(ref _progressBarColor, value); }
        public string ProgressBarBackColor { get => _progressBarBackColor; set => SetProperty(ref _progressBarBackColor, value); }
        public bool RecBold { get => _recBold; set => SetProperty(ref _recBold, value); }
        public string FontFamily { get => _fontFamily; set => SetProperty(ref _fontFamily, value); }
        public double FontSize { get => _fontSize; set => SetProperty(ref _fontSize, value); }
        public double HeaderFontSize { get => _headerFontSize; set => SetProperty(ref _headerFontSize, value); }
        public double FooterFontSize { get => _footerFontSize; set => SetProperty(ref _footerFontSize, value); }
        
        private double _serviceLogoHeight = 14.0;
        public double ServiceLogoHeight { get => _serviceLogoHeight; set => SetProperty(ref _serviceLogoHeight, value); }
        
        public double ItemPadding { get => _itemPadding; set => SetProperty(ref _itemPadding, value); }
        
        public double ListMarginLeft { get => _listMarginLeft; set { if (SetProperty(ref _listMarginLeft, value)) OnPropertyChanged(nameof(ListMargin)); } }
        public double ListMarginTop { get => _listMarginTop; set { if (SetProperty(ref _listMarginTop, value)) OnPropertyChanged(nameof(ListMargin)); } }
        public double ListMarginRight { get => _listMarginRight; set { if (SetProperty(ref _listMarginRight, value)) OnPropertyChanged(nameof(ListMargin)); } }
        public double ListMarginBottom { get => _listMarginBottom; set { if (SetProperty(ref _listMarginBottom, value)) OnPropertyChanged(nameof(ListMargin)); } }
        
        [XmlIgnore] public Thickness ListMargin => new Thickness(_listMarginLeft, _listMarginTop, _listMarginRight, _listMarginBottom);

        public double ToolTipFontSize { get => _toolTipFontSize; set => SetProperty(ref _toolTipFontSize, value); }
        public string ToolTipBackColor { get => _toolTipBackColor; set => SetProperty(ref _toolTipBackColor, value); }
        public string ToolTipForeColor { get => _toolTipForeColor; set => SetProperty(ref _toolTipForeColor, value); }
        public string ToolTipBorderColor { get => _toolTipBorderColor; set => SetProperty(ref _toolTipBorderColor, value); }
        public bool ShowToolTip { get => _showToolTip; set => SetProperty(ref _showToolTip, value); }
        public double ToolTipWidth { get => _toolTipWidth; set => SetProperty(ref _toolTipWidth, value); }

        public int ScrollAmountVertical { get; set; } = 3;
        public int ScrollAmountHorizontal { get; set; } = 3;

        public bool EnableTitleRemove { get => _enableTitleRemove; set => SetProperty(ref _enableTitleRemove, value); }
        public string TitleRemovePattern { get => _titleRemovePattern; set => SetProperty(ref _titleRemovePattern, value); }
        public bool ShowHeader { get => _showHeader; set => SetProperty(ref _showHeader, value); }
        public bool ShowListHeader { get => _showListHeader; set => SetProperty(ref _showListHeader, value); }
        public bool ShowFooter { get => _showFooter; set => SetProperty(ref _showFooter, value); }

        private bool _showServiceLogo = false;
        public bool ShowServiceLogo { get => _showServiceLogo; set => SetProperty(ref _showServiceLogo, value); }

        public string ColumnBorderColor { get => _columnBorderColor; set => SetProperty(ref _columnBorderColor, value); }
        public string FooterColor { get => _footerColor; set => SetProperty(ref _footerColor, value); }
        public string MainBorderColor { get => _mainBorderColor; set => SetProperty(ref _mainBorderColor, value); }

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
        private static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EM_Config.xml");

        public static void Load()
        {
            string path = ConfigPath;
            string tempPath = path + ".tmp"; // EM_Config.xml.tmp

            // 1. 復旧処理 (try-catchの外で行い成否に関わらずログを出さない)
            if (!File.Exists(path) && File.Exists(tempPath))
            {
                try
                {
                    File.Move(tempPath, path);
                }
                catch
                {
                    // 移動失敗（ファイルロック中など）時は何もしない
                }
            }

            // 2. 読み込み処理
            try
            {
                if (File.Exists(path))
                {
                    var serializer = new XmlSerializer(typeof(ConfigData));
                    using (var sr = new StreamReader(path, new UTF8Encoding(false)))
                    {
                        if (serializer.Deserialize(sr) is ConfigData loaded)
                        {
                            Data = loaded;
                        }
                    }
                }
        
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
            catch (Exception ex)
            {
                try { Logger.Write("設定読み込みエラー: " + ex.Message); } catch { }
            }
        }

        public static void Save()
        {
            string path = ConfigPath;
            string tempPath = path + ".tmp";

            try
            {
                var serializer = new XmlSerializer(typeof(ConfigData));
        
                // 1. 一時ファイルへ確実に書き出す
                using (var fs = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var sw = new StreamWriter(fs, new UTF8Encoding(false)))
                {
                    serializer.Serialize(sw, Data);
                    sw.Flush();
                    fs.Flush(true); 
                }

                // 2. 削除の隙間を作らずにアトミックに近い形でファイルを更新する
                File.Move(tempPath, path, true);
            }
            catch (Exception ex)
            {
                try { Logger.Write("設定保存エラー: " + ex.Message); } catch { }
            }
        }
        
    }
}