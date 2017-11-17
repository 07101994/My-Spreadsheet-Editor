#region Copyright Syncfusion Inc. 2001 - 2017
// Copyright Syncfusion Inc. 2001 - 2017. All rights reserved.
// Use of this code is subject to the terms of our license.
// A copy of the current license can be obtained at any time by e-mailing
// licensing@syncfusion.com. Any infringement will be prosecuted under
// applicable laws. 
#endregion
using Syncfusion.Windows.Tools.Controls;
using System.Windows;
using System.IO;
using System;
using Syncfusion.UI.Xaml.Grid.Utility;
using System.Windows.Media.Imaging;
using System.Linq;
using Syncfusion.UI.Xaml.SpreadsheetHelper;
using Syncfusion.UI.Xaml.Spreadsheet.GraphicCells;

namespace MySpreadsheetEditor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
		public string FilePath { get; set; }

        public MainWindow()
        {
            InitializeComponent();
			spreadsheetRibbon.Loaded += SpreadsheetRibbon_Loaded;
			//For importing charts,
			this.spreadsheetControl.AddGraphicChartCellRenderer(new GraphicChartCellRenderer());
		}

		private void SpreadsheetRibbon_Loaded(object sender, RoutedEventArgs e)
		{
			var ribbon1 = GridUtil.GetVisualChild<Ribbon>(sender as FrameworkElement);

			if (ribbon1 != null)
			{
				RibbonTab ribbonTab = new RibbonTab();
				ribbonTab.Caption = "OTHER";
				RibbonButton Button1 = new RibbonButton();
				Button1.Label = "PRINT";
				Button1.SmallIcon = new BitmapImage(new Uri("/../Icon/Icons_Print.png", UriKind.Relative));
				Button1.Click += PrintButton_Click;
				
				var customRibbonBar = new RibbonBar();
				customRibbonBar.Header = "Printing Options";
				customRibbonBar.Items.Add(Button1);
				customRibbonBar.IsLauncherButtonVisible = false;
				ribbonTab.Items.Add(customRibbonBar);
				ribbon1.Items.Add(ribbonTab);
			}
			//var ribbon1 = GridUtil.GetVisualChild<Ribbon>(sender as FrameworkElement);

			//// To add the ribbon button in View tab,

			//if (ribbon1 != null)
			//{
			//	var ribbonTab = ribbon1.Items[2] as RibbonTab;
			//	RibbonButton Button1 = new RibbonButton
			//	{
			//		Label = "PRINT",
			//		SmallIcon = new BitmapImage(new Uri("/../Icon/Icons_Print.png", UriKind.Relative))
			//	};
			//	Button1.Click += PrintButton_Click;
			//	ribbonTab.Items.Add(Button1);
			//}
		}

		private void PrintButton_Click(object sender, RoutedEventArgs e)
		{
			PrintPreviewWindow previewwindow = new PrintPreviewWindow() { Spreadsheet = spreadsheetControl };
			previewwindow.ShowDialog();
		}

		/// <summary>
		/// Provide support for Excel like closing operation when press the close button.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void RibbonWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
			spreadsheetControl.Dispose();
			spreadsheetRibbon.Dispose();
			System.Windows.Application.Current.Shutdown();
			//this.spreadsheetControl.Commands.FileClose.Execute(null);
   //         if (Application.Current.ShutdownMode != ShutdownMode.OnExplicitShutdown)
   //             e.Cancel = true;
        }
    }
}
