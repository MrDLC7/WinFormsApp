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
                row = Convert.ToInt32(textBoxRow.Text);
                column = ConvertColumnExcelToInt(textBoxColumn.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            LoadExcelData(row, column);
        }

        private void LoadExcelData(int row, int column)
        {
            try
            {
                OpenFileDialog fileExcel = new OpenFileDialog();

                fileExcel.InitialDirectory = "";
                fileExcel.DefaultExt = "*.xls;*.xlsx";
                fileExcel.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
                fileExcel.Title = "Виберіть документ";
                fileExcel.RestoreDirectory = true;

                if (fileExcel.ShowDialog() == DialogResult.OK && fileExcel.FileName.Length > 0)
                {
                    string path = fileExcel.FileName;
                    FileStream fstream = new FileStream(path, FileMode.Open);
                    Workbook workbook = new Workbook(fstream);
                    Worksheet worksheet = workbook.Worksheets[0];
                    DataTable table = worksheet.Cells.ExportDataTable(0, 0, row, column, true);
                    dataGridView.DataSource = table;
                    fstream.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private int ConvertColumnExcelToInt(string str)
        {
            int num = 0;
            foreach (char c in str)
            {
                num = num * 26 + (c - '@');
            }
            return num;
        }
    }
}
