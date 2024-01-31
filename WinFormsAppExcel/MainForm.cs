using Aspose.Cells;
using System.Data;
using System.Reflection;


namespace WinFormsAppExcel
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            int row = 0;
            int column = 0;
            try
            {
                // Завантаження таблиці по рядок, по колонку
                row = Convert.ToInt32(textBoxRow.Text);
                column = ConvertColumnExcelToInt(textBoxColumn.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            LoadExcelData(row, column);
        }

        // Завантаження файлу Excel
        private void LoadExcelData(int row, int column)
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

                // Перевірка на існування файлу
                if (fileExcel.ShowDialog() == DialogResult.OK && fileExcel.FileName.Length > 0)
                {
                    // Шлях до файлу
                    string path = fileExcel.FileName;

                    // Створення файлового потоку, що містить Excel, який потрібно відкрити
                    FileStream fstream = new FileStream(path, FileMode.Open);

                    // Створення екземпляру об'єкта Workbook
                    Workbook workbook = new Workbook(fstream);

                    // Доступ до першого в файлі Excel
                    Worksheet worksheet = workbook.Worksheets[0];

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

        // Конвертація заголовку із символу в число
        private int ConvertColumnExcelToInt(string str)
        {
            int num = 0;
            // Якщо строка не символ, а число
            bool success = int.TryParse(str, out int result);
            if (success)
                return Convert.ToInt32(str);

            foreach (char c in str)
            {
                num = num * 26 + (c - '@');
            }
            return num;
        }

        /*
        // Нумерування рядків по кліку
        private void dataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int rowIndex = e.RowIndex + 1;
            dataGridView.Rows[e.RowIndex].HeaderCell.Value = rowIndex.ToString();
        }
        */

        // Авто-нумерування рядків
        private void dataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            // Отримуємо рядок, який потрібно пронумерувати
            DataGridViewRow row = dataGridView.Rows[e.RowIndex];

            // Відображення номера рядка у заголовку
            using (SolidBrush brush = new SolidBrush(dataGridView.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, brush,
                    e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4);
            }
        }
    }
}
