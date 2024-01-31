namespace WinFormsAppExcel
{
    partial class MainForm
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
            groupBox = new GroupBox();
            buttonOpenFile = new Button();
            textBoxColumn = new TextBox();
            textBoxRow = new TextBox();
            labelColumn = new Label();
            labelRow = new Label();
            groupBox1 = new GroupBox();
            dataGridView = new DataGridView();
            openFileDialog = new OpenFileDialog();
            groupBox.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            SuspendLayout();
            // 
            // groupBox
            // 
            groupBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBox.AutoSize = true;
            groupBox.BackColor = SystemColors.ControlLight;
            groupBox.Controls.Add(buttonOpenFile);
            groupBox.Controls.Add(textBoxColumn);
            groupBox.Controls.Add(textBoxRow);
            groupBox.Controls.Add(labelColumn);
            groupBox.Controls.Add(labelRow);
            groupBox.Controls.Add(groupBox1);
            groupBox.Location = new Point(0, 0);
            groupBox.Name = "groupBox";
            groupBox.Size = new Size(800, 450);
            groupBox.TabIndex = 0;
            groupBox.TabStop = false;
            // 
            // buttonOpenFile
            // 
            buttonOpenFile.Anchor = AnchorStyles.Top;
            buttonOpenFile.BackColor = SystemColors.ButtonHighlight;
            buttonOpenFile.Location = new Point(463, 24);
            buttonOpenFile.Name = "buttonOpenFile";
            buttonOpenFile.Size = new Size(320, 56);
            buttonOpenFile.TabIndex = 0;
            buttonOpenFile.Text = "Завантажити документ Excel";
            buttonOpenFile.UseVisualStyleBackColor = false;
            buttonOpenFile.Click += buttonOpenFile_Click;
            // 
            // textBoxColumn
            // 
            textBoxColumn.Anchor = AnchorStyles.Top;
            textBoxColumn.Location = new Point(155, 58);
            textBoxColumn.Name = "textBoxColumn";
            textBoxColumn.Size = new Size(118, 27);
            textBoxColumn.TabIndex = 5;
            textBoxColumn.Text = "I";
            textBoxColumn.TextAlign = HorizontalAlignment.Center;
            // 
            // textBoxRow
            // 
            textBoxRow.Anchor = AnchorStyles.Top;
            textBoxRow.Location = new Point(155, 22);
            textBoxRow.Name = "textBoxRow";
            textBoxRow.Size = new Size(118, 27);
            textBoxRow.TabIndex = 4;
            textBoxRow.Text = "5";
            textBoxRow.TextAlign = HorizontalAlignment.Center;
            // 
            // labelColumn
            // 
            labelColumn.Anchor = AnchorStyles.Top;
            labelColumn.AutoSize = true;
            labelColumn.Location = new Point(17, 60);
            labelColumn.Name = "labelColumn";
            labelColumn.Size = new Size(97, 20);
            labelColumn.TabIndex = 3;
            labelColumn.Text = "По колонку:";
            // 
            // labelRow
            // 
            labelRow.Anchor = AnchorStyles.Top;
            labelRow.AutoSize = true;
            labelRow.Location = new Point(17, 24);
            labelRow.Name = "labelRow";
            labelRow.Size = new Size(81, 20);
            labelRow.TabIndex = 2;
            labelRow.Text = "По рядок:";
            // 
            // groupBox1
            // 
            groupBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBox1.Controls.Add(dataGridView);
            groupBox1.Location = new Point(7, 84);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(787, 358);
            groupBox1.TabIndex = 6;
            groupBox1.TabStop = false;
            // 
            // dataGridView
            // 
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.AllowUserToResizeColumns = false;
            dataGridView.AllowUserToResizeRows = false;
            dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView.BackgroundColor = SystemColors.ButtonHighlight;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new Point(3, 23);
            dataGridView.Name = "dataGridView";
            dataGridView.ReadOnly = true;
            dataGridView.RowHeadersWidth = 51;
            dataGridView.Size = new Size(781, 332);
            dataGridView.TabIndex = 2;
            // 
            // openFileDialog
            // 
            openFileDialog.FileName = "openFileDialog";
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(9F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(groupBox);
            Name = "MainForm";
            Text = "Excel export to Table";
            groupBox.ResumeLayout(false);
            groupBox.PerformLayout();
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox groupBox;
        private Button buttonOpenFile;
        private OpenFileDialog openFileDialog;
        private TextBox textBoxColumn;
        private TextBox textBoxRow;
        private Label labelColumn;
        private Label labelRow;
        private GroupBox groupBox1;
        private DataGridView dataGridView;
    }
}
