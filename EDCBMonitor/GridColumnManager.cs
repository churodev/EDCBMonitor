using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;

// 型の衝突を避けるための明示的な別名定義
using Brush = System.Windows.Media.Brush;
using Color = System.Windows.Media.Color;
using ListView = System.Windows.Controls.ListView;
using Control = System.Windows.Controls.Control;
using Point = System.Windows.Point; // これで衝突を解消

namespace EDCBMonitor
{
    public class GridColumnManager
    {
        private class ColumnDef
        {
            public string Header { get; set; } = "";
            public string BindingPath { get; set; } = "";
            public Func<bool> GetShow { get; set; } = () => true;
            public Func<double> GetWidth { get; set; } = () => 100;
        }

        private readonly ListView _listView;
        private readonly RoutedEventHandler _checkBoxHandler;

        public GridColumnManager(ListView listView, RoutedEventHandler checkBoxHandler)
        {
            _listView = listView;
            _checkBoxHandler = checkBoxHandler;
        }

        public void UpdateColumns()
        {
            if (_listView.View is not GridView gv) return;
            gv.Columns.Clear();

            var defs = GetColumnDefinitions();
            var orderedHeaders = Config.Data.ColumnHeaderOrder ?? new List<string>();
            var usedHeaders = new HashSet<string>();

            foreach (var header in orderedHeaders)
            {
                var d = defs.FirstOrDefault(x => x.Header == header);
                if (d != null && d.GetShow())
                {
                    AddColumnByType(gv, d);
                    usedHeaders.Add(header);
                }
            }

            foreach (var d in defs)
            {
                if (d.GetShow() && !usedHeaders.Contains(d.Header))
                {
                    AddColumnByType(gv, d);
                }
            }
        }

        public void UpdateHeaderStyle(Brush bg, Brush fg, Brush border)
        {
            if (_listView.View is not GridView gv) return;

            try
            {
                var headerStyle = new Style(typeof(GridViewColumnHeader));
                headerStyle.Setters.Add(new Setter(Control.FontSizeProperty, Config.Data.HeaderFontSize));
                headerStyle.Setters.Add(new Setter(UIElement.VisibilityProperty, Config.Data.ShowListHeader ? Visibility.Visible : Visibility.Collapsed));
                headerStyle.Setters.Add(new Setter(Control.BackgroundProperty, bg));
                headerStyle.Setters.Add(new Setter(Control.ForegroundProperty, fg));
                headerStyle.Setters.Add(new Setter(Control.BorderBrushProperty, border));
                headerStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1, 1, 1, 1)));
                headerStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(-1, 0, 0, 1)));
                headerStyle.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, System.Windows.HorizontalAlignment.Left));
                headerStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(4, 2, 0, 2)));
                
                headerStyle.Setters.Add(new Setter(FrameworkElement.ContextMenuProperty, CreateHeaderContextMenu()));

                var paddingTrigger = new Trigger { Property = GridViewColumnHeader.RoleProperty, Value = GridViewColumnHeaderRole.Padding };
                paddingTrigger.Setters.Add(new Setter(UIElement.VisibilityProperty, Visibility.Collapsed));
                headerStyle.Triggers.Add(paddingTrigger);

                if (System.Windows.Application.Current.MainWindow?.FindResource("HeaderTemplate") is ControlTemplate template)
                {
                    headerStyle.Setters.Add(new Setter(Control.TemplateProperty, template));
                }
                gv.ColumnHeaderContainerStyle = headerStyle;
            }
            catch (Exception ex) { Logger.Write($"Header Style Error: {ex.Message}"); }
        }

        public void SaveColumnState()
        {
            if (_listView.View is not GridView gv) return;

            Config.Data.ColumnHeaderOrder.Clear();
            foreach (var col in gv.Columns)
            {
                if (col.Header is string headerText)
                {
                    Config.Data.ColumnHeaderOrder.Add(headerText);
                }
            }

            foreach (var col in gv.Columns)
            {
                if (col.Header is not string header) continue;
                double w = col.ActualWidth;
                SaveColumnWidth(header, w);
            }
        }

        private List<ColumnDef> GetColumnDefinitions()
        {
            return new List<ColumnDef>
            {
                new() { Header = "状態", BindingPath = "Status", GetShow = () => Config.Data.ShowColStatus, GetWidth = () => Config.Data.WidthColStatus },
                new() { Header = "日時", BindingPath = "DateTimeInfo", GetShow = () => Config.Data.ShowColDateTime, GetWidth = () => Config.Data.WidthColDateTime },
                new() { Header = "長さ", BindingPath = "Duration", GetShow = () => Config.Data.ShowColDuration, GetWidth = () => Config.Data.WidthColDuration },
                new() { Header = "ネットワーク", BindingPath = "NetworkName", GetShow = () => Config.Data.ShowColNetwork, GetWidth = () => Config.Data.WidthColNetwork },
                new() { Header = "サービス名", BindingPath = "ServiceName", GetShow = () => Config.Data.ShowColServiceName, GetWidth = () => Config.Data.WidthColServiceName },
                new() { Header = "番組名", BindingPath = "Title", GetShow = () => Config.Data.ShowColTitle, GetWidth = () => Config.Data.WidthColTitle },
                new() { Header = "番組内容", BindingPath = "Desc", GetShow = () => Config.Data.ShowColDesc, GetWidth = () => Config.Data.WidthColDesc },
                new() { Header = "ジャンル", BindingPath = "Genre", GetShow = () => Config.Data.ShowColGenre, GetWidth = () => Config.Data.WidthColGenre },
                new() { Header = "付属情報", BindingPath = "ExtraInfo", GetShow = () => Config.Data.ShowColExtraInfo, GetWidth = () => Config.Data.WidthColExtraInfo },
                new() { Header = "有効", BindingPath = "IsEnabled", GetShow = () => Config.Data.ShowColEnabled, GetWidth = () => Config.Data.WidthColEnabled },
                new() { Header = "プログラム予約", BindingPath = "ProgramType", GetShow = () => Config.Data.ShowColProgramType, GetWidth = () => Config.Data.WidthColProgramType },
                new() { Header = "予約状況", BindingPath = "Comment", GetShow = () => Config.Data.ShowColComment, GetWidth = () => Config.Data.WidthColComment },
                new() { Header = "エラー状況", BindingPath = "ErrorInfo", GetShow = () => Config.Data.ShowColError, GetWidth = () => Config.Data.WidthColError },
                new() { Header = "予定ファイル名", BindingPath = "RecFileName", GetShow = () => Config.Data.ShowColRecFileName, GetWidth = () => Config.Data.WidthColRecFileName },
                new() { Header = "予定ファイル名リスト", BindingPath = "RecFileNameList", GetShow = () => Config.Data.ShowColRecFileNameList, GetWidth = () => Config.Data.WidthColRecFileNameList },
                new() { Header = "使用予定チューナー", BindingPath = "Tuner", GetShow = () => Config.Data.ShowColTuner, GetWidth = () => Config.Data.WidthColTuner },
                new() { Header = "予想サイズ", BindingPath = "EstimatedSize", GetShow = () => Config.Data.ShowColEstSize, GetWidth = () => Config.Data.WidthColEstSize },
                new() { Header = "プリセット", BindingPath = "Preset", GetShow = () => Config.Data.ShowColPreset, GetWidth = () => Config.Data.WidthColPreset },
                new() { Header = "録画モード", BindingPath = "RecMode", GetShow = () => Config.Data.ShowColRecMode, GetWidth = () => Config.Data.WidthColRecMode },
                new() { Header = "優先度", BindingPath = "Priority", GetShow = () => Config.Data.ShowColPriority, GetWidth = () => Config.Data.WidthColPriority },
                new() { Header = "追従", BindingPath = "Tuijyuu", GetShow = () => Config.Data.ShowColTuijyuu, GetWidth = () => Config.Data.WidthColTuijyuu },
                new() { Header = "ぴったり", BindingPath = "Pittari", GetShow = () => Config.Data.ShowColPittari, GetWidth = () => Config.Data.WidthColPittari },
                new() { Header = "チューナー強制", BindingPath = "TunerForce", GetShow = () => Config.Data.ShowColTunerForce, GetWidth = () => Config.Data.WidthColTunerForce },
                new() { Header = "録画後動作", BindingPath = "RecEndMode", GetShow = () => Config.Data.ShowColRecEndMode, GetWidth = () => Config.Data.WidthColRecEndMode },
                new() { Header = "復帰後再起動", BindingPath = "Reboot", GetShow = () => Config.Data.ShowColReboot, GetWidth = () => Config.Data.WidthColReboot },
                new() { Header = "録画後実行bat", BindingPath = "Bat", GetShow = () => Config.Data.ShowColBat, GetWidth = () => Config.Data.WidthColBat },
                new() { Header = "録画タグ", BindingPath = "RecTag", GetShow = () => Config.Data.ShowColRecTag, GetWidth = () => Config.Data.WidthColRecTag },
                new() { Header = "録画フォルダ", BindingPath = "RecFolder", GetShow = () => Config.Data.ShowColRecFolder, GetWidth = () => Config.Data.WidthColRecFolder },
                new() { Header = "開始", BindingPath = "StartMargin", GetShow = () => Config.Data.ShowColStartMargin, GetWidth = () => Config.Data.WidthColStartMargin },
                new() { Header = "終了", BindingPath = "EndMargin", GetShow = () => Config.Data.ShowColEndMargin, GetWidth = () => Config.Data.WidthColEndMargin },
                new() { Header = "ID", BindingPath = "ID", GetShow = () => Config.Data.ShowColID, GetWidth = () => Config.Data.WidthColID }
            };
        }

        private void AddColumnByType(GridView gv, ColumnDef d)
        {
            switch (d.Header)
            {
                case "有効": AddCheckBoxColumn(gv, d); break;
                case "日時": AddDateTimeColumn(gv, d); break;
                case "長さ": AddDurationColumn(gv, d); break;
                default: AddColumn(gv, d); break;
            }
        }

        private void AddColumn(GridView gv, ColumnDef d)
        {
            var dataTemplate = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(d.BindingPath));
            factory.SetValue(TextBlock.MarginProperty, new Thickness(2, Config.Data.ItemPadding, -6, Config.Data.ItemPadding));
            
            factory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.NoWrap);
            factory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.None);
            
            dataTemplate.VisualTree = factory;

            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = d.GetWidth(), CellTemplate = dataTemplate });
        }

        private void AddCheckBoxColumn(GridView gv, ColumnDef d)
        {
            var dataTemplate = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(System.Windows.Controls.CheckBox));
            factory.SetBinding(ToggleButton.IsCheckedProperty, new System.Windows.Data.Binding(d.BindingPath) { Mode = BindingMode.OneWay });
            factory.SetValue(FrameworkElement.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
            factory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            factory.AddHandler(System.Windows.Controls.Primitives.ButtonBase.ClickEvent, _checkBoxHandler);
            dataTemplate.VisualTree = factory;

            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = d.GetWidth(), CellTemplate = dataTemplate });
        }

        private void AddDateTimeColumn(GridView gv, ColumnDef d)
        {
            var dataTemplate = new DataTemplate();
            var gridFactory = new FrameworkElementFactory(typeof(Grid));

            if (!Config.Data.OmitProgress)
            {
                var progressFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.ProgressBar));
                progressFactory.SetBinding(RangeBase.ValueProperty, new System.Windows.Data.Binding("ProgressValue"));
                progressFactory.SetValue(RangeBase.MinimumProperty, 0.0);
                progressFactory.SetValue(RangeBase.MaximumProperty, 100.0);
                progressFactory.SetValue(Control.BorderThicknessProperty, new Thickness(0));
                progressFactory.SetValue(Control.BackgroundProperty, System.Windows.Media.Brushes.Transparent);

                try
                {
                    if (System.Windows.Media.ColorConverter.ConvertFromString(Config.Data.ProgressBarBackColor) is Color backColor)
                        progressFactory.SetValue(Control.BackgroundProperty, new SolidColorBrush(backColor));
                }
                catch { }

                try
                {
                    if (System.Windows.Media.ColorConverter.ConvertFromString(Config.Data.ProgressBarColor) is Color color)
                        progressFactory.SetValue(Control.ForegroundProperty, new SolidColorBrush(color));
                    else
                        progressFactory.SetValue(Control.ForegroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 255, 0)));
                }
                catch { }

                progressFactory.SetValue(FrameworkElement.MarginProperty, new Thickness(2, 0, -6, 0));
                progressFactory.SetValue(FrameworkElement.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Stretch);
                progressFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Stretch);
                progressFactory.SetValue(UIElement.RenderTransformOriginProperty, new Point(0.5, 1.0));
                progressFactory.SetValue(UIElement.RenderTransformProperty, new ScaleTransform(1.0, 0.20));

                var style = new Style(typeof(System.Windows.Controls.ProgressBar));
                style.Setters.Add(new Setter(UIElement.VisibilityProperty, Visibility.Collapsed));
                var trigger = new DataTrigger { Binding = new System.Windows.Data.Binding("IsRecording"), Value = true };
                trigger.Setters.Add(new Setter(UIElement.VisibilityProperty, Visibility.Visible));
                style.Triggers.Add(trigger);

                progressFactory.SetValue(FrameworkElement.StyleProperty, style);
                gridFactory.AppendChild(progressFactory);
            }

            var textFactory = new FrameworkElementFactory(typeof(TextBlock));
            textFactory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(d.BindingPath));
            textFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            textFactory.SetValue(FrameworkElement.MarginProperty, new Thickness(2, Config.Data.ItemPadding, -6, Config.Data.ItemPadding));
            gridFactory.AppendChild(textFactory);

            dataTemplate.VisualTree = gridFactory;
            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = d.GetWidth(), CellTemplate = dataTemplate });
        }

        private void AddDurationColumn(GridView gv, ColumnDef d)
        {
            var dataTemplate = new DataTemplate();
            var stackFactory = new FrameworkElementFactory(typeof(StackPanel));
            stackFactory.SetValue(StackPanel.OrientationProperty, System.Windows.Controls.Orientation.Horizontal);
            stackFactory.SetValue(FrameworkElement.MarginProperty, new Thickness(2, Config.Data.ItemPadding, -6, Config.Data.ItemPadding));

            var txtH = new FrameworkElementFactory(typeof(TextBlock));
            txtH.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("DurationHour"));

            var txtCol = new FrameworkElementFactory(typeof(TextBlock));
            txtCol.SetValue(TextBlock.TextProperty, ":");
            txtCol.SetBinding(UIElement.OpacityProperty, new System.Windows.Data.Binding("ColonOpacity"));

            var txtM = new FrameworkElementFactory(typeof(TextBlock));
            txtM.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("DurationMinute"));

            stackFactory.AppendChild(txtH);
            stackFactory.AppendChild(txtCol);
            stackFactory.AppendChild(txtM);

            dataTemplate.VisualTree = stackFactory;
            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = d.GetWidth(), CellTemplate = dataTemplate });
        }

        private ContextMenu CreateHeaderContextMenu()
        {
            var menu = new ContextMenu();
            
            void AddItem(string header, bool current, Action<bool> setAction)
            {
                var item = new MenuItem { Header = header, IsCheckable = true, IsChecked = current };
                item.Click += (s, e) => {
                    setAction(item.IsChecked);
                    // カラムの状態が変わったので設定を保存
                    Config.Save(); 
                    
                    System.Windows.Application.Current.MainWindow.Dispatcher.Invoke(() => {
                        // サイズ変更(updateSize)を伴わずに設定を再適用する
                        (System.Windows.Application.Current.MainWindow as MainWindow)?.ApplySettings(false);
                    });
                };
                menu.Items.Add(item);
            }

            AddItem("状態", Config.Data.ShowColStatus, v => Config.Data.ShowColStatus = v);
            AddItem("日時", Config.Data.ShowColDateTime, v => Config.Data.ShowColDateTime = v);
            AddItem("長さ", Config.Data.ShowColDuration, v => Config.Data.ShowColDuration = v);
            AddItem("ネットワーク", Config.Data.ShowColNetwork, v => Config.Data.ShowColNetwork = v);
            AddItem("サービス名", Config.Data.ShowColServiceName, v => Config.Data.ShowColServiceName = v);
            AddItem("番組名", Config.Data.ShowColTitle, v => Config.Data.ShowColTitle = v);
            
            menu.Items.Add(new Separator());
            AddItem("番組内容", Config.Data.ShowColDesc, v => Config.Data.ShowColDesc = v);
            AddItem("ジャンル", Config.Data.ShowColGenre, v => Config.Data.ShowColGenre = v);
            AddItem("付属情報", Config.Data.ShowColExtraInfo, v => Config.Data.ShowColExtraInfo = v);
            AddItem("有効/無効", Config.Data.ShowColEnabled, v => Config.Data.ShowColEnabled = v);
            AddItem("プログラム予約", Config.Data.ShowColProgramType, v => Config.Data.ShowColProgramType = v);
            
            menu.Items.Add(new Separator());
            AddItem("予約状況", Config.Data.ShowColComment, v => Config.Data.ShowColComment = v);
            AddItem("エラー状況", Config.Data.ShowColError, v => Config.Data.ShowColError = v);
            AddItem("予定ファイル名", Config.Data.ShowColRecFileName, v => Config.Data.ShowColRecFileName = v);
            AddItem("予定ファイル名リスト", Config.Data.ShowColRecFileNameList, v => Config.Data.ShowColRecFileNameList = v);
            
            menu.Items.Add(new Separator());
            AddItem("使用予定チューナー", Config.Data.ShowColTuner, v => Config.Data.ShowColTuner = v);
            AddItem("予想サイズ", Config.Data.ShowColEstSize, v => Config.Data.ShowColEstSize = v);
            AddItem("プリセット", Config.Data.ShowColPreset, v => Config.Data.ShowColPreset = v);
            AddItem("録画モード", Config.Data.ShowColRecMode, v => Config.Data.ShowColRecMode = v);
            AddItem("優先度", Config.Data.ShowColPriority, v => Config.Data.ShowColPriority = v);
            AddItem("追従", Config.Data.ShowColTuijyuu, v => Config.Data.ShowColTuijyuu = v);
            AddItem("ぴったり", Config.Data.ShowColPittari, v => Config.Data.ShowColPittari = v);
            AddItem("チューナー強制", Config.Data.ShowColTunerForce, v => Config.Data.ShowColTunerForce = v);
            
            menu.Items.Add(new Separator());
            AddItem("録画後動作", Config.Data.ShowColRecEndMode, v => Config.Data.ShowColRecEndMode = v);
            AddItem("復帰後再起動", Config.Data.ShowColReboot, v => Config.Data.ShowColReboot = v);
            AddItem("録画後実行bat", Config.Data.ShowColBat, v => Config.Data.ShowColBat = v);
            AddItem("録画タグ", Config.Data.ShowColRecTag, v => Config.Data.ShowColRecTag = v);
            AddItem("録画フォルダ", Config.Data.ShowColRecFolder, v => Config.Data.ShowColRecFolder = v);
            AddItem("開始", Config.Data.ShowColStartMargin, v => Config.Data.ShowColStartMargin = v);
            AddItem("終了", Config.Data.ShowColEndMargin, v => Config.Data.ShowColEndMargin = v);
            AddItem("ID", Config.Data.ShowColID, v => Config.Data.ShowColID = v);

            return menu;
        }

        private void SaveColumnWidth(string header, double width)
        {
            switch (header)
            {
                case "ID": Config.Data.WidthColID = width; break;
                case "状態": Config.Data.WidthColStatus = width; break;
                case "日時": Config.Data.WidthColDateTime = width; break;
                case "長さ": Config.Data.WidthColDuration = width; break;
                case "ネットワーク": Config.Data.WidthColNetwork = width; break;
                case "サービス名": Config.Data.WidthColServiceName = width; break;
                case "番組名": Config.Data.WidthColTitle = width; break;
                case "番組内容": Config.Data.WidthColDesc = width; break;
                case "ジャンル": Config.Data.WidthColGenre = width; break;
                case "付属情報": Config.Data.WidthColExtraInfo = width; break;
                case "有効": Config.Data.WidthColEnabled = width; break;
                case "プログラム予約": Config.Data.WidthColProgramType = width; break;
                case "予約状況": Config.Data.WidthColComment = width; break;
                case "エラー状況": Config.Data.WidthColError = width; break;
                case "予定ファイル名": Config.Data.WidthColRecFileName = width; break;
                case "予定ファイル名リスト": Config.Data.WidthColRecFileNameList = width; break;
                case "使用予定チューナー": Config.Data.WidthColTuner = width; break;
                case "予想サイズ": Config.Data.WidthColEstSize = width; break;
                case "プリセット": Config.Data.WidthColPreset = width; break;
                case "録画モード": Config.Data.WidthColRecMode = width; break;
                case "優先度": Config.Data.WidthColPriority = width; break;
                case "追従": Config.Data.WidthColTuijyuu = width; break;
                case "ぴったり": Config.Data.WidthColPittari = width; break;
                case "チューナー強制": Config.Data.WidthColTunerForce = width; break;
                case "録画後動作": Config.Data.WidthColRecEndMode = width; break;
                case "復帰後再起動": Config.Data.WidthColReboot = width; break;
                case "録画後実行bat": Config.Data.WidthColBat = width; break;
                case "録画タグ": Config.Data.WidthColRecTag = width; break;
                case "録画フォルダ": Config.Data.WidthColRecFolder = width; break;
                case "開始": Config.Data.WidthColStartMargin = width; break;
                case "終了": Config.Data.WidthColEndMargin = width; break;
            }
        }
    }
}