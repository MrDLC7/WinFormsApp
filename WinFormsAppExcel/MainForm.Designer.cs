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
            dataGridView = new DataGridView();
            buttonOpenFile = new Button();
            openFileDialog = new OpenFileDialog();
            groupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            SuspendLayout();
            // 
            // groupBox
            // 
            groupBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBox.Controls.Add(dataGridView);
            groupBox.Controls.Add(buttonOpenFile);
            groupBox.Location = new Point(11, -6);
            groupBox.Name = "groupBox";
            groupBox.Size = new Size(778, 444);
            groupBox.TabIndex = 0;
            groupBox.TabStop = false;
            // 
            // dataGridView
            // 
            dataGridView.Anchor = AnchorStyles.None;
            dataGridView.BackgroundColor = SystemColors.ButtonHighlight;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new Point(6, 87);
            dataGridView.Name = "dataGridView";
            dataGridView.RowHeadersWidth = 51;
            dataGridView.Size = new Size(766, 351);
            dataGridView.TabIndex = 1;
            // 
            // buttonOpenFile
            // 
            buttonOpenFile.Anchor = AnchorStyles.Top;
            buttonOpenFile.BackColor = SystemColors.ButtonFace;
            buttonOpenFile.Location = new Point(231, 19);
            buttonOpenFile.Name = "buttonOpenFile";
            buttonOpenFile.Size = new Size(320, 56);
            buttonOpenFile.TabIndex = 0;
            buttonOpenFile.Text = "Відкрити документ Excel";
            buttonOpenFile.UseVisualStyleBackColor = false;
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
            Text = "Form1";
            groupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox;
        private Button buttonOpenFile;
        private DataGridView dataGridView;
        private OpenFileDialog openFileDialog;
    }
}
