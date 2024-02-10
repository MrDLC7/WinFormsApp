using System.Data;

namespace WinFormsAppPaymentUrl
{
    public partial class PaymentUrl : Form
    {
        public PaymentUrl()
        {
            InitializeComponent();
        }
        // Для зберрігання шляху до файлу
        static protected string path = string.Empty;
        // Створення екземпляру класу DataTable для збереження даних
        static protected DataTable dataTable = new();


        // Відкриття файлу Excel
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
                    OpenAndUpdateExcel.OpenFile();
                }

                // Встановлення DataTable як джерела даних для DataGridView
                dataGridView.DataSource = dataTable;
                // Зміна розміру комірок взалежності від вмісту
                dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            Text = "PaymentUrl" + "      " + path;
        }
        
        // Оновлення файлу Excel
        private void btnUpdateDataFile_Click(object sender, EventArgs e)
        {
            int debtIndexColumn = -1;
            int statusIndexColumn = -1;
            int statusIndexRow = -1;
            int linkIndexColumn = -1;

            try
            {
                // Пошук індексів потрібних заголовків
                OpenAndUpdateExcel.SearchColumn_Header("Payment Link", 
                    out statusIndexRow, out linkIndexColumn);

                OpenAndUpdateExcel.SearchColumn_Header("Status", 
                    out statusIndexRow, out statusIndexColumn);

                OpenAndUpdateExcel.SearchColumn_Header("Debt", 
                    out statusIndexRow, out debtIndexColumn);

                int Nrow = dataTable.Rows.Count;   // Кількіть рядків без заголовка
                while (statusIndexRow < Nrow)
                {
                    // Оновлення колонки з заголовком "Статус"
                    OpenAndUpdateExcel.UpdatePayStatus("open", "pending", 
                        statusIndexRow, statusIndexColumn, linkIndexColumn, debtIndexColumn);

                    statusIndexRow++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

            // Оновлення документу Excel
            OpenAndUpdateExcel.UpdateFile();
            // Зміна розміру комірок взалежності від вмісту
            dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            MessageBox.Show("Посилання сформовані");
        }
    }
}
