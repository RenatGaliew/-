using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Подбор_кандидатов
{
    public class AccessExcel
    {
        private ExcelPackage _appExcel;
        private ExcelWorksheet _xlsSheet;

        public int RowsCount => _xlsSheet.Cells.Rows;
        public int ColumnCount => _xlsSheet.Cells.Columns;
        private DataTable _dataTable = new DataTable();
        public void DoAccess(string path)
        {
            try
            {
                _appExcel = new ExcelPackage(new FileInfo(path));
                while (_appExcel.Workbook.Worksheets.Count == 0)
                    _appExcel.Workbook.Worksheets.Add("Лист 1");
                _xlsSheet = _appExcel.Workbook.Worksheets[1];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FinishAccess()
        {
            _appExcel.Save();
            _appExcel.Dispose();
        }

        public double GetValueCell(int columnIndex, int rowIndex)
        {
            return (double)_xlsSheet.Cells[columnIndex, rowIndex].Value;
        }

        public void WriteRow(int index, int[] row, int count)
        {
            _xlsSheet = _appExcel.Workbook.Worksheets[1];
            for (int i = 0; i < count; i++)
            {
                _xlsSheet.Cells[i + 1, index].Value = row[i];
            }
        }

        public void WriteCell(int iRow, int iColumn, double data)
        {
            _xlsSheet = _appExcel.Workbook.Worksheets[1];
            _xlsSheet.Cells[iRow, iColumn].Value = data;
        }

        public void WriteCell(int iRow, int iColumn, string data)
        {
            _xlsSheet = _appExcel.Workbook.Worksheets[1];
            _xlsSheet.Cells[iRow, iColumn].Value = data;
        }

        public T ReadCell<T>(int iRow, int iColumn)
        {
            _xlsSheet = _appExcel.Workbook.Worksheets[1];
            T result = default(T);
            try
            {
               result = (T)_xlsSheet.Cells[iRow, iColumn].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        public double ReadCellDouble(int iRow, int iColumn)
        {
            _xlsSheet = _appExcel.Workbook.Worksheets[1];
            double result = default(double);
            try
            {
                var tmp = _xlsSheet.Cells[iRow, iColumn].Value.ToString().Replace('.', ',');
                result = double.Parse(tmp);
                result = double.Parse(tmp);
                result = double.Parse(tmp);
                result = double.Parse(tmp);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        public DateTime ReadDate(int iRow, int iColumn)
        {
            _xlsSheet = _appExcel.Workbook.Worksheets[1];
            DateTime result = default(DateTime);
            try
            {
                result = DateTime.Parse(_xlsSheet.Cells[iRow, iColumn].Value.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        public void WriteCell<T>(int iRow, int iColumn, T Data)
        {
            _xlsSheet = _appExcel.Workbook.Worksheets[1];
            try
            {
                _xlsSheet.Cells[iRow, iColumn].Value = Data;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public int MaxRows()
        {
            int Count = 0;

            for (int i = 1; i < 1000; i++)
            {
                if (_xlsSheet.Cells[i, 1].Value == null)
                {
                    Count = i;
                    break;
                }
            }
            return Count;
        }
    }
}
