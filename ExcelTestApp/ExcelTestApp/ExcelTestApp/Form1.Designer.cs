namespace ExcelTestApp
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxSheet = new System.Windows.Forms.TextBox();
            this.column2textBox = new System.Windows.Forms.TextBox();
            this.column3textBox = new System.Windows.Forms.TextBox();
            this.column1textBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(240, 88);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(32, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Номер сторінки";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Колонка 2";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 101);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Колонка 3";
            // 
            // textBoxSheet
            // 
            this.textBoxSheet.Location = new System.Drawing.Point(106, 12);
            this.textBoxSheet.Name = "textBoxSheet";
            this.textBoxSheet.Size = new System.Drawing.Size(100, 20);
            this.textBoxSheet.TabIndex = 4;
            // 
            // column2textBox
            // 
            this.column2textBox.Location = new System.Drawing.Point(106, 68);
            this.column2textBox.Name = "column2textBox";
            this.column2textBox.Size = new System.Drawing.Size(100, 20);
            this.column2textBox.TabIndex = 5;
            // 
            // column3textBox
            // 
            this.column3textBox.Location = new System.Drawing.Point(106, 98);
            this.column3textBox.Name = "column3textBox";
            this.column3textBox.Size = new System.Drawing.Size(100, 20);
            this.column3textBox.TabIndex = 6;
            // 
            // column1textBox
            // 
            this.column1textBox.Location = new System.Drawing.Point(106, 42);
            this.column1textBox.Name = "column1textBox";
            this.column1textBox.Size = new System.Drawing.Size(100, 20);
            this.column1textBox.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 45);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Колонка 1";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(240, 13);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(32, 23);
            this.button2.TabIndex = 9;
            this.button2.Text = "button";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 123);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.column1textBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.column3textBox);
            this.Controls.Add(this.column2textBox);
            this.Controls.Add(this.textBoxSheet);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxSheet;
        private System.Windows.Forms.TextBox column2textBox;
        private System.Windows.Forms.TextBox column3textBox;
        private System.Windows.Forms.TextBox column1textBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button2;
    }
}

