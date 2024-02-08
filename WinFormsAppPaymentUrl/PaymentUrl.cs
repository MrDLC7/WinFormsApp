using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Common;
using System.Data;
using System.Windows.Forms;
using System.IO.Compression;
using System.Xml;
using Microsoft.Office.Interop.Excel;

namespace WinFormsAppPaymentUrl
{
    public partial class PaymentUrl : Form
    {
        public PaymentUrl()
        {
            InitializeComponent();
        }

        private int rows = 0, columns = 0;
        private string path = string.Empty;
        // Створення екземпляру класу DataTable для збереження даних
        private System.Data.DataTable dataTable = new();

        private void btnOpenFile_Click(object sender, EventArgs e)
        {

            try
            {
                // Створення вікна вибору файлу Excel
                OpenFileDialog fileExcel = new();
                // Директорія
                fileExcel.InitialDirectory = "";
                // Фільтр по файлах
                fileExcel.DefaultExt = "*.xls;*.xlsx";
                fileExcel.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
                fileExcel.Title = "Select document Excel";
                // Запам'ятати директорію
                fileExcel.RestoreDirectory = true;

                // Відкриття і перевірка на існування файлу
                if (fileExcel.ShowDialog() == DialogResult.OK && fileExcel.FileName.Length > 0)
                {
                    // Шлях до файлу
                    path = fileExcel.FileName;
                    OpenFileExcel();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btnUpdateDataFile_Click(object sender, EventArgs e)
        {
            int statusIndexColumn = -1;
            int statusIndexRow = -1;
            int linkIndexColumn = -1;

            try
            {
                SearchString("Payment Link", out statusIndexRow, out linkIndexColumn);
                SearchString("Status", out statusIndexRow, out statusIndexColumn);


                while (statusIndexRow >= 0)
                {
                    statusIndexRow = SearchStringForStatus("open", "pending", statusIndexColumn);
                    AddPaymentUrl(statusIndexRow, linkIndexColumn);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            UpdateFileExcel();
        }

        private void OpenFileExcel()
        {

            // Відкриття файлу
            Excel.Application app = new();
            // Створення екземпляру об'єкта Workbook
            Workbook workbook = app.Workbooks.Open(path);

            // Доступ до першого листа в файлі Excel
            _Worksheet worksheet = workbook.Sheets[1];

            // Отримання даних
            Excel.Range range = worksheet.UsedRange;

            // Макс. кількість рядків і колонок
            rows = range.Rows.Count;
            columns = range.Columns.Count;

            try
            {
                for (int row = 1; row <= rows; row++)
                {
                    DataRow dr = dataTable.NewRow();
                    for (int column = 1; column <= columns; column++)
                    {
                        if (row == 1)
                        {
                            dataTable.Columns.Add(range.Cells[row, column].Value2.ToString());
                        }
                        else
                        {
                            dr[column - 1] = range.Cells[row, column].Value2;
                        }
                    }
                    if (row != 1)
                    {
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
                // Встановлення DataTable як джерела даних для DataGridView
                dataGridView.DataSource = dataTable;
                // Закриття об'єктів Excel
                workbook.Close();
                app.Quit();
            }
        }

        private void UpdateFileExcel()
        {
            try
            {

                /*
                // Створення екземпляру об'єкта Workbook
                Workbook workbook = new Workbook();

                // Доступ до першого в файлі Excel
                Worksheet worksheet = workbook.Worksheets[0];

                /*
                // Створення екземпляру класу DataTable для збереження даних
                DataTable dataTable = new DataTable();

                // Встановлення DataTable як джерела даних для DataGridView
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    dataTable.Columns.Add(column.HeaderText, typeof(string));
                }
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    DataRow dataRow = dataTable.NewRow();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        dataRow[cell.ColumnIndex] = cell.Value;
                    }
                    dataTable.Rows.Add(dataRow);
                }


                // Заповнення DataTable даними з Excel
                ImportTableOptions options = new ImportTableOptions();
                options.IsFieldNameShown = false;
                worksheet.Cells.ImportData(dataTable, 0, 0, options);

                workbook.Save(path);
                // Закриття файлового потоку
                workbook.Dispose();

                */
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void SearchString(string searchString, out int rowIndex, out int columnIndex)
        {
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                if (column.HeaderText != null && column.HeaderText.ToString().Contains(searchString))
                {
                    // Нашли подстроку, сохраняем индексы строки и столбца
                    rowIndex = 0;
                    columnIndex = column.Index;
                    return; // Если нужно найти только первое вхождение в DataGridView, раскомментируйте эту строку
                }
            }
            // Перебираем строки в DataGridView
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // Перебираем ячейки в текущей строке
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Проверяем, содержит ли значение ячейки искомую подстроку
                    if (cell.Value != null && cell.Value.ToString().Contains(searchString))
                    {
                        // Нашли подстроку, сохраняем индексы строки и столбца
                        rowIndex = row.Index;
                        columnIndex = cell.ColumnIndex;
                        return; // Если нужно найти только первое вхождение в DataGridView, раскомментируйте эту строку
                    }
                }
            }
            rowIndex = -1;
            columnIndex = -1;
        }

        private int SearchStringForStatus(string searchString, string renameString, int columnIndex)
        {
            // Перебираем строки в DataGridView
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // Перебираем ячейки в текущей строке
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Проверяем, содержит ли значение ячейки искомую подстроку
                    if (cell.Value != null && cell.ColumnIndex == columnIndex &&
                        cell.Value.ToString().Contains(searchString))
                    {
                        // Нашли подстроку, сохраняем индексы строки
                        cell.Value = renameString;
                        return row.Index;
                    }
                }
            }
            return -1;
        }

        private bool AddPaymentUrl(int rowIndex, int columnIndex)
        {
            // Перебираем строки в DataGridView
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.Index == rowIndex)
                {
                    // Перебираем ячейки в текущей строке
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        // Проверяем, содержит ли значение ячейки искомую подстроку
                        if (cell.ColumnIndex == columnIndex)
                        {
                            cell.Value = "yes";
                            return true;
                        }
                    }
                }
            }
            return false;
        }

    }
}
