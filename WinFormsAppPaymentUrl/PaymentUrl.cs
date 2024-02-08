using Aspose.Cells;
using System.Data.Common;
using System.Data;
using System.Windows.Forms;
using System.IO.Compression;
using System.Xml;

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
        private DataTable dataTable;

        private void btnOpenFile_Click(object sender, EventArgs e)
        {

            try
            {
                // Створення вікна вибору файлу Excel
                OpenFileDialog fileExcel = new OpenFileDialog();
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
            try
            {
                // Створення файлового потоку, що містити Excel, який потрібно перевірити
                FileStream fstream = new FileStream(path, FileMode.Open);

                // Створення екземпляру об'єкта Workbook
                Workbook workbook = new Workbook(fstream);

                // Доступ до першого в файлі Excel
                Worksheet worksheet = workbook.Worksheets[0];

                // Макс. кількість рядків і колонок
                rows = worksheet.Cells.MaxDataRow + 1;
                columns = worksheet.Cells.MaxDataColumn + 1;


                // Заповнення DataTable даними з Excel
                dataTable = worksheet.Cells.ExportDataTable(0, 0, rows, columns, true);
                // Встановлення DataTable як джерела даних для DataGridView
                dataGridView.DataSource = dataTable;

                // Закриття файлового потоку
                workbook.Dispose();
                fstream.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void UpdateFileExcel()
        {
            try
            {

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
                */

                // Заповнення DataTable даними з Excel
                ImportTableOptions options = new ImportTableOptions();
                options.IsFieldNameShown = false;
                worksheet.Cells.ImportData(dataTable, 0, 0, options);

                workbook.Save(path);
                // Закриття файлового потоку
                workbook.Dispose();
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
            return true;
        }
    
    }
}
