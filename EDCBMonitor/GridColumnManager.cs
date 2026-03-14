using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using System.Collections.Specialized;
using System.Windows.Threading;

using Brush = System.Windows.Media.Brush;
using Color = System.Windows.Media.Color;
using ListView = System.Windows.Controls.ListView;
using Control = System.Windows.Controls.Control;
using Point = System.Windows.Point;

namespace EDCBMonitor
{
    public class GridColumnManager
    {
        private class ColumnDef
        {
            public string Header { get; set; } = "";
            public string BindingPath { get; set; } = "";
        }

        private readonly ListView _listView;
        private readonly RoutedEventHandler _checkBoxHandler;
        private bool _isUpdatingColumns = false;
        private bool _isSaveQueued = false;

        public GridColumnManager(ListView listView, RoutedEventHandler checkBoxHandler)
        {
            _listView = listView;
            _checkBoxHandler = checkBoxHandler;
        }

        public void UpdateColumns()
        {
            if (_listView.View is not GridView gv) return;

            // 1. データの不足があればデフォルト値で補完
            Config.Data.InitDefaults();

            _isUpdatingColumns = true;
            try
            {
                gv.Columns.Clear();
                var defs = GetColumnDefinitions();

                // 2. 設定(Config.Data.Columns)の順序通りに、表示設定がONのものだけを追加
                foreach (var state in Config.Data.Columns)
                {
                    if (state.IsVisible)
                    {
                        var d = defs.FirstOrDefault(x => x.Header == state.Header);
                        if (d != null)
                        {
                            AddColumnByType(gv, d, state.Width);
                        }
                    }
                }
            }
            finally
            {
                _isUpdatingColumns = false;
            }

            if (gv.Columns is INotifyCollectionChanged newCollection)
            {
                newCollection.CollectionChanged += Columns_CollectionChanged;
            }
        }

        private void Columns_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (_isUpdatingColumns) return;

            // ドラッグ移動(Move)や項目の増減を検知
            if (e.Action == NotifyCollectionChangedAction.Move ||
                e.Action == NotifyCollectionChangedAction.Add ||
                e.Action == NotifyCollectionChangedAction.Remove)
            {
                if (_isSaveQueued) return;
                _isSaveQueued = true;

                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    _isSaveQueued = false;
                    SaveColumnState();
                    Config.Save();
                }), DispatcherPriority.ApplicationIdle);
            }
        }

        public void SaveColumnState()
        {
            if (_listView.View is not GridView gv) return;

            // 1. 各カラムの幅を保存
            foreach (var col in gv.Columns)
            {
                if (col.Header is string headerText)
                {
                    Config.Data.GetColumn(headerText).Width = col.ActualWidth;
                }
            }

            // 2. 表示されているカラムの順序を取得
            var visibleHeaders = gv.Columns.Select(c => c.Header as string).Where(h => h != null).ToList();

            // 3. 全カラムリスト(Config.Data.Columns)を、表示列だけ現在のUIの順序に入れ替えて再構成
            var currentList = Config.Data.Columns.ToList();
            var visibleColumnsInMaster = currentList.Where(c => c.IsVisible).OrderBy(c => {
                int idx = visibleHeaders.IndexOf(c.Header);
                return idx == -1 ? int.MaxValue : idx;
            }).ToList();

            var newMaster = new List<ColumnState>();
            int visibleIdx = 0;
            foreach (var original in Config.Data.Columns)
            {
                if (original.IsVisible && visibleIdx < visibleColumnsInMaster.Count)
                {
                    newMaster.Add(visibleColumnsInMaster[visibleIdx++]);
                }
                else if (!original.IsVisible)
                {
                    newMaster.Add(original);
                }
            }

            Config.Data.Columns = newMaster;
        }

        public void UpdateHeaderStyle(Brush bg, Brush fg, Brush border, bool? showListHeader = null)
        {
            if (_listView.View is not GridView gv) return;

            try
            {
                var headerStyle = new Style(typeof(GridViewColumnHeader));
                headerStyle.Setters.Add(new Setter(Control.FontSizeProperty, Config.Data.HeaderFontSize));
                
                bool isListHeaderVisible = showListHeader ?? Config.Data.ShowListHeader;
                headerStyle.Setters.Add(new Setter(UIElement.VisibilityProperty, isListHeaderVisible ? Visibility.Visible : Visibility.Collapsed));
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

        private List<ColumnDef> GetColumnDefinitions()
        {
            return new List<ColumnDef>
            {
                new() { Header = "状態", BindingPath = "Status" },
                new() { Header = "日時", BindingPath = "DateTimeInfo" },
                new() { Header = "長さ", BindingPath = "Duration" },
                new() { Header = "ネットワーク", BindingPath = "NetworkName" },
                new() { Header = "サービス名", BindingPath = "ServiceName" },
                new() { Header = "番組名", BindingPath = "Title" },
                new() { Header = "番組内容", BindingPath = "Desc" },
                new() { Header = "ジャンル", BindingPath = "Genre" },
                new() { Header = "付属情報", BindingPath = "ExtraInfo" },
                new() { Header = "有効", BindingPath = "IsEnabled" },
                new() { Header = "プログラム予約", BindingPath = "ProgramType" },
                new() { Header = "予約状況", BindingPath = "Comment" },
                new() { Header = "エラー状況", BindingPath = "ErrorInfo" },
                new() { Header = "予定ファイル名", BindingPath = "RecFileName" },
                new() { Header = "予定ファイル名リスト", BindingPath = "RecFileNameList" },
                new() { Header = "使用予定チューナー", BindingPath = "Tuner" },
                new() { Header = "予想サイズ", BindingPath = "EstimatedSize" },
                new() { Header = "プリセット", BindingPath = "Preset" },
                new() { Header = "録画モード", BindingPath = "RecMode" },
                new() { Header = "優先度", BindingPath = "Priority" },
                new() { Header = "追従", BindingPath = "Tuijyuu" },
                new() { Header = "ぴったり", BindingPath = "Pittari" },
                new() { Header = "チューナー強制", BindingPath = "TunerForce" },
                new() { Header = "録画後動作", BindingPath = "RecEndMode" },
                new() { Header = "復帰後再起動", BindingPath = "Reboot" },
                new() { Header = "録画後実行bat", BindingPath = "Bat" },
                new() { Header = "録画タグ", BindingPath = "RecTag" },
                new() { Header = "録画フォルダ", BindingPath = "RecFolder" },
                new() { Header = "開始", BindingPath = "StartMargin" },
                new() { Header = "終了", BindingPath = "EndMargin" },
                new() { Header = "ID", BindingPath = "ID" }
            };
        }

        private void AddColumnByType(GridView gv, ColumnDef d, double width)
        {
            switch (d.Header)
            {
                case "有効": AddCheckBoxColumn(gv, d, width); break;
                case "日時": AddDateTimeColumn(gv, d, width); break;
                case "長さ": AddDurationColumn(gv, d, width); break;
                case "サービス名": AddServiceNameColumn(gv, d, width); break;
                default: AddColumn(gv, d, width); break;
            }
        }

        private void AddServiceNameColumn(GridView gv, ColumnDef d, double width)
        {
            var dataTemplate = new DataTemplate();
            var gridFactory = new FrameworkElementFactory(typeof(Grid));
            gridFactory.SetValue(FrameworkElement.MarginProperty, new Thickness(2, Config.Data.ItemPadding, -6, Config.Data.ItemPadding));

            var imgFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.Image));
            imgFactory.SetBinding(System.Windows.Controls.Image.SourceProperty, new System.Windows.Data.Binding("ServiceLogo"));
            imgFactory.SetBinding(UIElement.VisibilityProperty, new System.Windows.Data.Binding("ServiceLogoVisibility"));
            imgFactory.SetValue(FrameworkElement.HeightProperty, Config.Data.ServiceLogoHeight); 
            imgFactory.SetValue(System.Windows.Controls.Image.StretchProperty, Stretch.Uniform);
            imgFactory.SetValue(FrameworkElement.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Left);
            imgFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);

            var txtFactory = new FrameworkElementFactory(typeof(TextBlock));
            txtFactory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(d.BindingPath));
            txtFactory.SetBinding(UIElement.VisibilityProperty, new System.Windows.Data.Binding("ServiceNameVisibility"));
            txtFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            txtFactory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.NoWrap);
            txtFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.None);

            gridFactory.AppendChild(imgFactory);
            gridFactory.AppendChild(txtFactory);

            dataTemplate.VisualTree = gridFactory;
            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = width, CellTemplate = dataTemplate });
        }

        private void AddColumn(GridView gv, ColumnDef d, double width)
        {
            var dataTemplate = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(d.BindingPath));
            factory.SetValue(TextBlock.MarginProperty, new Thickness(2, Config.Data.ItemPadding, -6, Config.Data.ItemPadding));
            factory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.NoWrap);
            factory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.None);
            
            dataTemplate.VisualTree = factory;
            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = width, CellTemplate = dataTemplate });
        }

        private void AddCheckBoxColumn(GridView gv, ColumnDef d, double width)
        {
            var dataTemplate = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(System.Windows.Controls.CheckBox));
            factory.SetBinding(ToggleButton.IsCheckedProperty, new System.Windows.Data.Binding(d.BindingPath) { Mode = BindingMode.OneWay });
            factory.SetValue(FrameworkElement.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
            factory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            factory.AddHandler(System.Windows.Controls.Primitives.ButtonBase.ClickEvent, _checkBoxHandler);
            
            dataTemplate.VisualTree = factory;
            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = width, CellTemplate = dataTemplate });
        }

        private void AddDateTimeColumn(GridView gv, ColumnDef d, double width)
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
                } catch { }

                try
                {
                    if (System.Windows.Media.ColorConverter.ConvertFromString(Config.Data.ProgressBarColor) is Color color)
                        progressFactory.SetValue(Control.ForegroundProperty, new SolidColorBrush(color));
                    else
                        progressFactory.SetValue(Control.ForegroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 255, 0)));
                } catch { }

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
            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = width, CellTemplate = dataTemplate });
        }

        private void AddDurationColumn(GridView gv, ColumnDef d, double width)
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
            gv.Columns.Add(new GridViewColumn { Header = d.Header, Width = width, CellTemplate = dataTemplate });
        }

        private ContextMenu CreateHeaderContextMenu()
        {
            var menu = new ContextMenu();
            
            void AddItem(string header, bool current, Action<bool> setAction)
            {
                var item = new MenuItem { Header = header, IsCheckable = true, IsChecked = current };
                item.Click += (s, e) => {
                    SaveColumnState();
                    setAction(item.IsChecked);
                    Config.Save(); 
                    
                    System.Windows.Application.Current.MainWindow.Dispatcher.Invoke(() => {
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
    }
}