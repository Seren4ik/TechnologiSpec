using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace TechnologiSpec
{
    class ExcelHelper : IDisposable
    {
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        #region Открытие создание файла ексель
        internal bool Open(string filpath)
        {
            try
            {
                if (File.Exists(filpath))
                {
                    _workbook = _excel.Workbooks.Open(filpath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filpath;
                }
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;

        }
        #endregion

        #region Заполнение ячеек
        internal bool Set(object column, object row, object data)
        {
            try
            {
                Excel.Range range1 = ((Excel.Worksheet)_excel.ActiveSheet).Range[row, column].Value2 = data;
                range1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        #endregion

        #region Заполнение ячеек2
        internal bool Cells(int column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;                
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        #endregion

        #region Объединение ячеек
        internal bool Unite(object Cell1, object Cell2)
        {
            try
            {
                Excel.Range range2 = ((Excel.Worksheet)_excel.ActiveSheet).Range[Cell1, Cell2];
                range2.Select();
                range2.Merge();
                range2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        #endregion

        internal void save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
            }
        }
        internal void visible()
        {
            _excel.Visible = true;
        }

        public void Dispose()
        {
            try
            {
                _workbook.Close();     

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
