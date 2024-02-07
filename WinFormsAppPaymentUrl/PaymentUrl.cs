using Aspose.Cells;
using System.Data.Common;
using System.Data;

namespace WinFormsAppPaymentUrl
{
    public partial class PaymentUrl : Form
    {
        public PaymentUrl()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            int row = 0, column = 0;
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
                    string path = fileExcel.FileName;

                    // Створення файлового потоку, що містити Excel, який потрібно перевірити
                    FileStream fstream = new FileStream(path, FileMode.Open);

                    // Створення екземпляру об'єкта Workbook
                    Workbook workbook = new Workbook(fstream);

                    // Доступ до першого в файлі Excel
                    Worksheet worksheet = workbook.Worksheets[0];
                    
                    // Макс. кількість рядків і колонок
                    row = worksheet.Cells.MaxDataRow + 1;
                    column = worksheet.Cells.MaxDataColumn + 1;

                    // Створення екземпляру класу DataTable для збереження даних
                    DataTable table = new DataTable();

                    // Заповнення DataTable даними з Excel
                    table = worksheet.Cells.ExportDataTable(0, 0, row, column, true);

                    // Встановлення DataTable як джерела даних для DataGridView
                    dataGridView.DataSource = table;

                    // Закриття файлового потоку
                    fstream.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }


        private void btnUpdateDataFile_Click(object sender, EventArgs e)
        {

        }
    }
}
