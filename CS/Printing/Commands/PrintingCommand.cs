using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Pdf;
using Syncfusion.UI.Xaml.Spreadsheet;
using Syncfusion.Windows.PdfViewer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MySpreadsheetEditor.Printing.Commands
{
	public static class PrintingCommand
	{
		static PrintingCommand()
		{
			CommandManager.RegisterClassCommandBinding(typeof(SfSpreadsheet), new CommandBinding(PrintPreview, OnExecutePrintPreview, OnCanExecutePrintPreview));
			CommandManager.RegisterClassCommandBinding(typeof(SfSpreadsheet), new CommandBinding(DirectPrint, OnExecuteDirectPrint, OnCanExecuteDirectPrint));
		}


		#region PrintPreview

		public static RoutedCommand PrintPreview = new RoutedCommand("PrintPreview", typeof(SfSpreadsheet));

		private static void OnExecutePrintPreview(object sender, ExecutedRoutedEventArgs args)
		{
			SfSpreadsheet spreadsheetControl = args.Source as SfSpreadsheet;
			PrintPreviewWindow previewwindow = new PrintPreviewWindow() { Spreadsheet = spreadsheetControl };
			previewwindow.ShowDialog();
		}

		private static void OnCanExecutePrintPreview(object sender, CanExecuteRoutedEventArgs args)
		{
			args.CanExecute = true;
		}

		#endregion

		#region DirectPrint

		public static RoutedCommand DirectPrint = new RoutedCommand("DirectPrint", typeof(SfSpreadsheet));

		private static void OnExecuteDirectPrint(object sender, ExecutedRoutedEventArgs args)
		{
			SfSpreadsheet spreadsheetControl = args.Source as SfSpreadsheet;
			PrintFromPdfViewer(spreadsheetControl);

		}

		private static void OnCanExecuteDirectPrint(object sender, CanExecuteRoutedEventArgs args)
		{
			args.CanExecute = true;
		}

		#endregion

		#region Direct print through PdfViewer

		private static void PrintFromPdfViewer(SfSpreadsheet spreadsheetControl)
		{
			//Create the pdfviewer for load the document.
			PdfViewerControl pdfviewer = new PdfViewerControl();

			// PdfDocumentViewer
			MemoryStream pdfstream = new MemoryStream();

			ExcelToPdfConverter converter = new ExcelToPdfConverter(spreadsheetControl.Workbook);
			//Intialize the PdfDocument
			PdfDocument pdfDoc = new PdfDocument();

			//Intialize the ExcelToPdfConverter Settings
			ExcelToPdfConverterSettings settings = new ExcelToPdfConverterSettings();
			settings.LayoutOptions = LayoutOptions.NoScaling;

			//Assign the PdfDocument to the templateDocument property of ExcelToPdfConverterSettings
			settings.TemplateDocument = pdfDoc;
			settings.DisplayGridLines = GridLinesDisplayStyle.Invisible;

			//Convert Excel Document into PDF document
			pdfDoc = converter.Convert(settings);

			//Save the PDF file
			pdfDoc.Save(pdfstream);

			//Load the document to pdfviewer
			pdfviewer.Load(pdfstream);

			//Show the print dialog.
			pdfviewer.Print(true);

		}
		#endregion


	}
}
