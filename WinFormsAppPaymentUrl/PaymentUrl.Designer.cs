namespace WinFormsAppPaymentUrl
{
    partial class PaymentUrl
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            panelMain = new Panel();
            dataGridView = new DataGridView();
            btnUpdateDataFile = new Button();
            btnOpenFile = new Button();
            openFileDialog = new OpenFileDialog();
            panelMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            SuspendLayout();
            // 
            // panelMain
            // 
            panelMain.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            panelMain.Controls.Add(dataGridView);
            panelMain.Controls.Add(btnUpdateDataFile);
            panelMain.Controls.Add(btnOpenFile);
            panelMain.Location = new Point(1, 1);
            panelMain.Name = "panelMain";
            panelMain.Size = new Size(582, 351);
            panelMain.TabIndex = 0;
            // 
            // dataGridView
            // 
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView.BackgroundColor = SystemColors.ButtonHighlight;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new Point(11, 86);
            dataGridView.Name = "dataGridView";
            dataGridView.ReadOnly = true;
            dataGridView.RowHeadersWidth = 51;
            dataGridView.Size = new Size(558, 254);
            dataGridView.TabIndex = 2;
            // 
            // btnUpdateDataFile
            // 
            btnUpdateDataFile.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnUpdateDataFile.BackColor = Color.YellowGreen;
            btnUpdateDataFile.Location = new Point(289, 11);
            btnUpdateDataFile.Name = "btnUpdateDataFile";
            btnUpdateDataFile.Size = new Size(275, 60);
            btnUpdateDataFile.TabIndex = 1;
            btnUpdateDataFile.Text = "Сформувати посилання";
            btnUpdateDataFile.UseVisualStyleBackColor = false;
            btnUpdateDataFile.Click += btnUpdateDataFile_Click;
            // 
            // btnOpenFile
            // 
            btnOpenFile.BackColor = Color.YellowGreen;
            btnOpenFile.Location = new Point(11, 11);
            btnOpenFile.Name = "btnOpenFile";
            btnOpenFile.Size = new Size(275, 60);
            btnOpenFile.TabIndex = 0;
            btnOpenFile.Text = "Відкрити файл Excel";
            btnOpenFile.UseVisualStyleBackColor = false;
            btnOpenFile.Click += btnOpenFile_Click;
            // 
            // openFileDialog
            // 
            openFileDialog.FileName = "openFileDialog";
            // 
            // PaymentUrl
            // 
            AutoScaleDimensions = new SizeF(9F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(582, 353);
            Controls.Add(panelMain);
            MinimumSize = new Size(600, 400);
            Name = "PaymentUrl";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "PaymentUrl";
            panelMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Panel panelMain;
        private DataGridView dataGridView;
        private Button btnUpdateDataFile;
        private Button btnOpenFile;
        private OpenFileDialog openFileDialog;
    }
}
