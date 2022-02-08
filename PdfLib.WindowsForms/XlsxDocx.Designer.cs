namespace PdfLib.WindowsForms
{
    partial class XlsxDocx
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.btnXlsxToPdf = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(9, 10);
            this.button1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(226, 84);
            this.button1.TabIndex = 0;
            this.button1.Text = "Insert Image to Cell";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnInsertImageToXlsxCell_Click);
            // 
            // btnXlsxToPdf
            // 
            this.btnXlsxToPdf.Location = new System.Drawing.Point(250, 10);
            this.btnXlsxToPdf.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnXlsxToPdf.Name = "btnXlsxToPdf";
            this.btnXlsxToPdf.Size = new System.Drawing.Size(226, 84);
            this.btnXlsxToPdf.TabIndex = 1;
            this.btnXlsxToPdf.Text = "XLSX + DOCX -> PDF";
            this.btnXlsxToPdf.UseVisualStyleBackColor = true;
            this.btnXlsxToPdf.Click += new System.EventHandler(this.btnXlsxToPdf_Click);
            // 
            // XlsxDocx
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 106);
            this.Controls.Add(this.btnXlsxToPdf);
            this.Controls.Add(this.button1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "XlsxDocx";
            this.Text = "XLSX & DOCX -> PDF";
            this.Load += new System.EventHandler(this.XlsxDocx_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnXlsxToPdf;
    }
}

