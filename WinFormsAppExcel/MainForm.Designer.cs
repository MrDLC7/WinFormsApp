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
            groupBoxFirst = new GroupBox();
            buttonOpenFile = new Button();
            textBoxColumn = new TextBox();
            textBoxRow = new TextBox();
            labelColumn = new Label();
            labelRow = new Label();
            groupBoxSecond = new GroupBox();
            dataGridView = new DataGridView();
            openFileDialog = new OpenFileDialog();
            groupBoxFirst.SuspendLayout();
            groupBoxSecond.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            SuspendLayout();
            // 
            // groupBoxFirst
            // 
            groupBoxFirst.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxFirst.AutoSize = true;
            groupBoxFirst.BackColor = SystemColors.ControlLight;
            groupBoxFirst.Controls.Add(buttonOpenFile);
            groupBoxFirst.Controls.Add(textBoxColumn);
            groupBoxFirst.Controls.Add(textBoxRow);
            groupBoxFirst.Controls.Add(labelColumn);
            groupBoxFirst.Controls.Add(labelRow);
            groupBoxFirst.Controls.Add(groupBoxSecond);
            groupBoxFirst.Location = new Point(0, 0);
            groupBoxFirst.Name = "groupBoxFirst";
            groupBoxFirst.Size = new Size(800, 450);
            groupBoxFirst.TabIndex = 0;
            groupBoxFirst.TabStop = false;
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
            // groupBoxSecond
            // 
            groupBoxSecond.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxSecond.Controls.Add(dataGridView);
            groupBoxSecond.Location = new Point(7, 84);
            groupBoxSecond.Name = "groupBoxSecond";
            groupBoxSecond.Size = new Size(787, 358);
            groupBoxSecond.TabIndex = 6;
            groupBoxSecond.TabStop = false;
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
            dataGridView.Margin = new Padding(10);
            dataGridView.Name = "dataGridView";
            dataGridView.ReadOnly = true;
            dataGridView.RowHeadersWidth = 51;
            dataGridView.Size = new Size(781, 332);
            dataGridView.TabIndex = 2;
            dataGridView.RowPostPaint += dataGridView_RowPostPaint;
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
            Controls.Add(groupBoxFirst);
            MinimumSize = new Size(818, 497);
            Name = "MainForm";
            Text = "Excel export to Table";
            groupBoxFirst.ResumeLayout(false);
            groupBoxFirst.PerformLayout();
            groupBoxSecond.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox groupBoxFirst;
        private Button buttonOpenFile;
        private OpenFileDialog openFileDialog;
        private TextBox textBoxColumn;
        private TextBox textBoxRow;
        private Label labelColumn;
        private Label labelRow;
        private GroupBox groupBoxSecond;
        private DataGridView dataGridView;
    }
}
