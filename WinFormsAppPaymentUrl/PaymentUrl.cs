using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Common;
using System.Data;
using System.Windows.Forms;
using System.IO.Compression;
using System.Xml;
//using Microsoft.Office.Interop.Excel;

namespace WinFormsAppPaymentUrl
{
    public partial class PaymentUrl : Form
    {
        public PaymentUrl()
        {
            InitializeComponent();
        }

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

                // Встановлення DataTable як джерела даних для DataGridView
                dataGridView.DataSource = dataTable;
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
                // Пошук індексів потрібних заголовків
                SearchColumn_Header("Payment Link", out statusIndexRow, out linkIndexColumn);
                SearchColumn_Header("Status", out statusIndexRow, out statusIndexColumn);

                int N = dataTable.Rows.Count;   // Кількіть рядків без заголовка
                while (statusIndexRow < N)
                {
                    // Оновлення колонки з заголовком "Статус"
                    UpdatePayStatus("open", "pending", statusIndexRow, statusIndexColumn, linkIndexColumn);
                    statusIndexRow++;
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

        private void UpdateFileExcel()
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
                str += dataTable.Rows[rowIndex][i] + " ";
            return str;
        }
    }
}
