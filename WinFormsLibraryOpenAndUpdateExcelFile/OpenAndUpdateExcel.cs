using System.Data;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsLibraryOpenAndUpdateExcelFile
{
    public class OpenAndUpdateExcel
    {

        public void OpenFile(string path, DataTable dataTable)
        {
            // Відкриття файлу
            Excel.Application app = new();
            // Створення екземпляру об'єкта Workbook
            Excel.Workbook workbook = app.Workbooks.Open(path);

            // Доступ до першого листа в файлі Excel
            Excel._Worksheet worksheet = workbook.Sheets[1];

            // Отримання даних
            Excel.Range range = worksheet.UsedRange;

            // Макс. кількість рядків і колонок
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;

            try
            {
                // Заповнення DataTable даними з Excel
                for (int row = 1; row <= rows; row++)
                {
                    DataRow dr = dataTable.NewRow();
                    for (int column = 1; column <= columns; column++)
                    {
                        if (row == 1)
                        {
                            // Додавання стовпців для першого рядка Excel
                            dataTable.Columns.Add(range.Cells[row, column].Value2.ToString());
                        }
                        else
                        {
                            // Додавання даних у DataTable
                            dr[column - 1] = range.Cells[row, column].Value2;
                        }
                    }
                    if (row != 1)
                    {
                        // Додавання рядка до DataTable
                        dataTable.Rows.Add(dr);
                    }
                    dataTable.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Закриття об'єктів Excel
                workbook.Close();

                // Закриття додатку Excel
                app.Quit();
            }
        }

        public void UpdateFile(string path, DataTable dataTable)
        {
            // Відкриття файлу
            Excel.Application app = new();
            // Створення екземпляру об'єкта Workbook
            Excel.Workbook workbook = app.Workbooks.Open(path);

            // Доступ до першого листа в файлі Excel
            Excel._Worksheet worksheet = workbook.Sheets[1];

            try
            {
                int Nrows = dataTable.Rows.Count;
                int Ncols = dataTable.Columns.Count;


                // Заповнення заголовків Excel, даними з DataTable
                for (int column = 0; column < Ncols; column++)
                {
                    worksheet.Cells[1, column + 1] = dataTable.Columns[column].ColumnName;
                }

                // Заповнення полів Excel, даними з DataTable
                for (int row = 0; row < Nrows; row++)
                {
                    for (int column = 0; column < Ncols; column++)
                    {
                        worksheet.Cells[row + 2, column + 1] = dataTable.Rows[row][column].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Збереження книги Excel
                workbook.Save();
                // Закриття об'єктів Excel
                workbook.Close();
                // Закриття додатку Excel
                app.Quit();
            }
        }
        private void SearchColumn_Header(string searchString, out int rowIndex, out int columnIndex)
        {
            // Значення за замовчуванням
            rowIndex = columnIndex = 0;
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                if (column.HeaderText != null && column.HeaderText.ToString().Contains(searchString))
                {
                    // Знайшли заголовок, зберігаємо індекси рядка та стовпця і виходимо
                    columnIndex = column.Index;
                    return;
                }
            }
            MessageBox.Show("Заголовок не знайдено");
        }

        private void UpdatePayStatus(string in_String, string out_String, int rowIndex, int columnIndex, int linkIndexColumn)
        {
            if (dataTable.Rows[rowIndex][columnIndex].ToString().Contains(in_String))
            {
                dataTable.Rows[rowIndex][columnIndex] = out_String;
                // Додавання посилання
                dataTable.Rows[rowIndex][linkIndexColumn] = AddPaymentUrl(rowIndex, columnIndex);
            }
        }

        private string AddPaymentUrl(int rowIndex, int columnIndex)
        {
            string str = string.Empty;
            for (int i = 0; i < columnIndex; i++)
                str += dataTable.Rows[rowIndex][i] + (((i + 1) == columnIndex) ? "" : " ");
            return str;
        }
    }
}
