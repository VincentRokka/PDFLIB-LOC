using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ConsoleApp3
{
	/// <summary>
	/// エクセルからグラフを取得してWordに貼りPDFとして出力する
	/// </summary>
	class Program
	{
		[STAThread()]
		static void Main(string[] args)
		{
			var currentPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
			var xlFileName = args[0];
			var xlMacroName = args[1];
			string wdFileName = args[2];
			string pdfFileName = args[3];

			var xlFilePath = Path.Combine(currentPath, xlFileName);
			object wdFilePath = Path.Combine(currentPath, wdFileName);
			string pdfFilePath = Path.Combine(currentPath, pdfFileName);

			Microsoft.Office.Interop.Excel.Application xlApp = null;
			Workbooks xlWorkbooks = null;
			Workbook xlWorkbook = null;

			Microsoft.Office.Interop.Word.Application wdApp = null;
			Document document = null;

			try
			{
				xlApp = new Microsoft.Office.Interop.Excel.Application();
				xlApp.Visible = false;
				xlWorkbooks = xlApp.Workbooks;
				// 指定した.xlsmファイルを開く
				xlWorkbooks.Open(xlFilePath);
				//var fileName = Path.GetFileName(xlFilePath);
				// 指定したマクロを実行				
				xlApp.Run(xlFileName + "!" + xlMacroName);
				//xlApp.Run("1_140220.xlsm!Copy_Chart");


				wdApp = new Microsoft.Office.Interop.Word.Application();
				wdApp.Visible = false;
				document = wdApp.Documents.Add();
				// [互換モード]を解除する
				document.SetCompatibilityMode((int)WdCompatibilityMode.wdCurrent);

				if (Clipboard.ContainsImage())
				{
					var img = Clipboard.GetImage();
					if (img != null)
					{
						object start = 0;
						object end = 1;
						Microsoft.Office.Interop.Word.Range rng = document.Range(ref start, ref end);
						rng.Paste();
					}
				}

				// 文書を保存
				document.SaveAs2(ref wdFilePath);
				document.ExportAsFixedFormat(pdfFilePath, WdExportFormat.wdExportFormatPDF);

				// エクセルを閉じる
				xlWorkbooks.Close();
				xlWorkbooks = null;

				// 文書を閉じる
				document.Close();
				document = null;
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			finally
			{
				// 使用したことのあるCOMオブジェクトを最後に呼び出したものから開放する
				if (document != null)
				{
					Marshal.ReleaseComObject(document);
					document = null;
				}
				if (wdApp != null)
				{
					Marshal.ReleaseComObject(wdApp);
					wdApp = null;
				}
				if (xlWorkbook != null)
				{
					Marshal.ReleaseComObject(xlWorkbook);
					xlWorkbook = null;
				}
				if (xlWorkbooks != null)
				{
					Marshal.ReleaseComObject(xlWorkbooks);
					xlWorkbooks = null;
				}
				if (xlApp != null)
				{
					Marshal.ReleaseComObject(xlApp);
					xlApp = null;
				}
				// ここで一度ガベージコレクションを実行する
				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
				// EXCELアプリケーションのクローズはEXCELに関連するオブジェクトを開放してから実行してから
				// EXCELのCOMオブジェクトを開放する
				if (xlApp != null)
				{
					xlApp.Quit();
					Marshal.ReleaseComObject(xlApp);
					xlApp = null;
				}
				if (wdApp != null)
				{
					wdApp.Quit();
					Marshal.ReleaseComObject(wdApp);
					wdApp = null;
				}
				// アプリケーションを終了するためにもう一度ガベージコレクションを実行する
				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
			}
		}
	}
}
