using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CALab3
{
    class MyExcel
    {

        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range chartRange;

        public MyExcel(string[] data,string filename)
        {
            string dataLength = data.Length.ToString();
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(Type.Missing);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            WriteData(data);

            //Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(150, 100, 300, 250);
            //Excel.Chart chartPage = myChart.Chart;
            var chartPage = ChartPosition(200, 100, 300, 250);

            chartRange = xlWorkSheet.get_Range("A1", $"A{dataLength}");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

            var folderChoose = OpenDialog.FolderChooser();
            xlWorkBook.SaveAs($"{folderChoose}{filename}.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show($"Excel file created , you can find the file {folderChoose}{filename}.xls");

        }
        private void WriteData(string[] data)
        {
            for (int i = 0, j = 1; i < data.GetLength(0); i++, j++)
            {
                xlWorkSheet.Cells[j, 1] = data[i];
            }
        }
        private Excel.Chart ChartPosition(int leftLocation, int topLocation, int heigth, int width)
        {
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(150, 100, 300, 250);
            Excel.Chart chartPage = myChart.Chart;
            return chartPage;
        }
    }
}
