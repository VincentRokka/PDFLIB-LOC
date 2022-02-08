using Microsoft.Office.Core;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace PdfLib.WindowsForms
{
    public partial class XlsxDocx : Form
    {
        private string sXlsxFilePath = string.Empty;
        private string sXlsmFilePath = string.Empty;
        private string sImageFilePath = string.Empty;
        private string sXlsmChartFilePath = string.Empty;
        private string sXlsmTableFilePath = string.Empty;

        public XlsxDocx()
        {
            InitializeComponent();

            sXlsxFilePath = $@"{Environment.CurrentDirectory}\Sample.xlsx";
            sXlsmFilePath = $@"{Environment.CurrentDirectory}\Sample.xlsm";
            sImageFilePath = $@"{Environment.CurrentDirectory}\1x1.png";
            sXlsmChartFilePath = $@"{Environment.CurrentDirectory}\140397_1_Template.xlsm";
            sXlsmTableFilePath = $@"{Environment.CurrentDirectory}\140397_2_Template.xlsm";
        }

        private void btnInsertImageToXlsxCell_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sXlsxFilePath);
            xlWorkSheet = xlWorkBook.Worksheets["Sheet1"];
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            Excel.Range oRange = (Excel.Range)xlWorkSheet.Cells[1, 1];
            float left = (float)((double)oRange.Left);
            float top = (float)((double)oRange.Top);
            float width = (float)((double)oRange.Width);
            float height = (float)((double)oRange.Height);
            xlWorkSheet.Shapes.AddPicture(sImageFilePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top, width, height);
        }

        /// <summary>
        /// Convert cell/range/chart MS EXCEL -> MS WORD -> PDF
        /// Input: Excel file (.xlsx)
        /// Temp: Word (.docx)
        /// Output: PDF (.pdf)
        ///   + Copy cell/range from.xlsx to clipboard
        ///   + Open new word app.docx
        ///   + Paste from clipboard to word
        ///   + Save as pdf
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnXlsxToPdf_Click(object sender, EventArgs e)
        {
            //Assumption: Have an object in Clipboard
            //TODO: WORD -> Paste -> Paste Special ...
            //  + Paste link -> Microsoft Excel Chart Object
            //TODO: Save as PDF

            var currentPath = Environment.CurrentDirectory;
            var xlFileName = "1_140220_2.xlsm";
            //var xlMacroName = "ExportPdf";
            string wdFileName = "1_140220_2.docx";
            string pdfFileName = "1_140220_2.pdf";

            var xlFilePath = Path.Combine(currentPath, xlFileName);
            object wdFilePath = Path.Combine(currentPath, wdFileName);
            string pdfFilePath = Path.Combine(currentPath, pdfFileName);

            Excel.Application xlApp = null;
            Excel.Workbooks xlWorkbooks = null;
            Excel.Workbook xlWorkbook = null;

            Word.Application wdApp = null;
            Word.Document document = null;

            try
            {
                wdApp = new Word.Application();
                wdApp.Visible = false;
                document = wdApp.Documents.Add();
                // [互換モード]を解除する
                document.SetCompatibilityMode((int)Word.WdCompatibilityMode.wdCurrent);

                if (Clipboard.ContainsImage())
                {
                    var img = Clipboard.GetImage();
                    if (img != null)
                    {
                        //object start = 0;
                        //object end = 1;
                        //Word.Range rng = document.Range(ref start, ref end);
                        //TODO: CONG + TUNG + LOC + DAT + DUC
                        //Paste => Paste Special ... => Paste link -> Microsoft Excel Chart Object
                        //rng.Paste();

                        //Paste Special
                        //wdApp.Selection.PasteSpecial(Link: true, DataType: Word.WdPasteDataType.wdPasteOLEObject);

                        //Paste Link
                        wdApp.Selection.PasteSpecial(Link: true, DataType: Word.WdPasteDataType.wdPasteMetafilePicture);
                        MessageBox.Show("Success");
                    }
                }

                // 文書を保存
                // Save to .DOCX (from Clipboard)
                document.SaveAs2(ref wdFilePath);

                // Save to .PDF (from WORD)
                document.ExportAsFixedFormat(pdfFilePath, Word.WdExportFormat.wdExportFormatPDF);

                // エクセルを閉じる
                if (xlWorkbooks != null)
                {
                    xlWorkbooks.Close();
                }
                xlWorkbooks = null;

                // 文書を閉じる
                if (document != null)
                {
                    document.Close();
                }
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

        private void XlsxDocx_Load(object sender, EventArgs e)
        {

        }
    }
}