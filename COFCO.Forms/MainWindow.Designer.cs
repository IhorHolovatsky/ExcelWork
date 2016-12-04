namespace COFCO.Forms
{
    partial class MainWindow
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lblPort = new System.Windows.Forms.Label();
            this.lblSupplier = new System.Windows.Forms.Label();
            this.lblProduct = new System.Windows.Forms.Label();
            this.lblQuantity = new System.Windows.Forms.Label();
            this.lblDate = new System.Windows.Forms.Label();
            this.lblVehicleNumber = new System.Windows.Forms.Label();
            this.lblTTNNumber = new System.Windows.Forms.Label();
            this.lblContract = new System.Windows.Forms.Label();
            this.lblSheetNumber = new System.Windows.Forms.Label();
            this.lblStartRowNumber = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbContact = new System.Windows.Forms.TextBox();
            this.tbTTNNumber = new System.Windows.Forms.TextBox();
            this.tbVehicleNumber = new System.Windows.Forms.TextBox();
            this.tbDate = new System.Windows.Forms.TextBox();
            this.tbQuantity = new System.Windows.Forms.TextBox();
            this.tbProduct = new System.Windows.Forms.TextBox();
            this.tbPort = new System.Windows.Forms.TextBox();
            this.tbSupplier = new System.Windows.Forms.TextBox();
            this.tbSheetNumber = new System.Windows.Forms.TextBox();
            this.tbStartRowNumber = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnCreateTempExcel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnChooseOutputTempFolder = new System.Windows.Forms.Button();
            this.tbOutputFolderPath = new System.Windows.Forms.TextBox();
            this.lblOutputFolder = new System.Windows.Forms.Label();
            this.lblInputFile = new System.Windows.Forms.Label();
            this.tbInputFilePath = new System.Windows.Forms.TextBox();
            this.btnChooseInputFile = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnCreateTemplates = new System.Windows.Forms.Button();
            this.tbOutputTemplateFolder = new System.Windows.Forms.TextBox();
            this.lblOutputTemplateFolder = new System.Windows.Forms.Label();
            this.btnChooseTemplateFolderPath = new System.Windows.Forms.Button();
            this.btnChooseExcelFile = new System.Windows.Forms.Button();
            this.tbExcelFilePath = new System.Windows.Forms.TextBox();
            this.lblExcelFileWithContracts = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(805, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // lblPort
            // 
            this.lblPort.AutoSize = true;
            this.lblPort.Location = new System.Drawing.Point(20, 31);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(32, 13);
            this.lblPort.TabIndex = 1;
            this.lblPort.Text = "Порт";
            // 
            // lblSupplier
            // 
            this.lblSupplier.AutoSize = true;
            this.lblSupplier.Location = new System.Drawing.Point(20, 56);
            this.lblSupplier.Name = "lblSupplier";
            this.lblSupplier.Size = new System.Drawing.Size(79, 13);
            this.lblSupplier.TabIndex = 2;
            this.lblSupplier.Text = "Постачальник";
            // 
            // lblProduct
            // 
            this.lblProduct.AutoSize = true;
            this.lblProduct.Location = new System.Drawing.Point(20, 83);
            this.lblProduct.Name = "lblProduct";
            this.lblProduct.Size = new System.Drawing.Size(49, 13);
            this.lblProduct.TabIndex = 3;
            this.lblProduct.Text = "Продукт";
            // 
            // lblQuantity
            // 
            this.lblQuantity.AutoSize = true;
            this.lblQuantity.Location = new System.Drawing.Point(20, 109);
            this.lblQuantity.Name = "lblQuantity";
            this.lblQuantity.Size = new System.Drawing.Size(53, 13);
            this.lblQuantity.TabIndex = 4;
            this.lblQuantity.Text = "Кількість";
            // 
            // lblDate
            // 
            this.lblDate.AutoSize = true;
            this.lblDate.Location = new System.Drawing.Point(20, 134);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(33, 13);
            this.lblDate.TabIndex = 5;
            this.lblDate.Text = "Дата";
            // 
            // lblVehicleNumber
            // 
            this.lblVehicleNumber.AutoSize = true;
            this.lblVehicleNumber.Location = new System.Drawing.Point(20, 158);
            this.lblVehicleNumber.Name = "lblVehicleNumber";
            this.lblVehicleNumber.Size = new System.Drawing.Size(84, 13);
            this.lblVehicleNumber.TabIndex = 6;
            this.lblVehicleNumber.Text = "Номер машини";
            // 
            // lblTTNNumber
            // 
            this.lblTTNNumber.AutoSize = true;
            this.lblTTNNumber.Location = new System.Drawing.Point(20, 183);
            this.lblTTNNumber.Name = "lblTTNNumber";
            this.lblTTNNumber.Size = new System.Drawing.Size(66, 13);
            this.lblTTNNumber.TabIndex = 7;
            this.lblTTNNumber.Text = "Номер ТТН";
            // 
            // lblContract
            // 
            this.lblContract.AutoSize = true;
            this.lblContract.Location = new System.Drawing.Point(20, 208);
            this.lblContract.Name = "lblContract";
            this.lblContract.Size = new System.Drawing.Size(54, 13);
            this.lblContract.TabIndex = 8;
            this.lblContract.Text = "Контракт";
            // 
            // lblSheetNumber
            // 
            this.lblSheetNumber.AutoSize = true;
            this.lblSheetNumber.Location = new System.Drawing.Point(32, 322);
            this.lblSheetNumber.Name = "lblSheetNumber";
            this.lblSheetNumber.Size = new System.Drawing.Size(73, 13);
            this.lblSheetNumber.TabIndex = 9;
            this.lblSheetNumber.Text = "Номер листа";
            // 
            // lblStartRowNumber
            // 
            this.lblStartRowNumber.AutoSize = true;
            this.lblStartRowNumber.Location = new System.Drawing.Point(32, 347);
            this.lblStartRowNumber.Name = "lblStartRowNumber";
            this.lblStartRowNumber.Size = new System.Drawing.Size(135, 13);
            this.lblStartRowNumber.TabIndex = 10;
            this.lblStartRowNumber.Text = "Початковий номер рядка";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblContract);
            this.groupBox1.Controls.Add(this.lblPort);
            this.groupBox1.Controls.Add(this.lblSupplier);
            this.groupBox1.Controls.Add(this.tbContact);
            this.groupBox1.Controls.Add(this.lblProduct);
            this.groupBox1.Controls.Add(this.tbTTNNumber);
            this.groupBox1.Controls.Add(this.lblTTNNumber);
            this.groupBox1.Controls.Add(this.tbVehicleNumber);
            this.groupBox1.Controls.Add(this.lblQuantity);
            this.groupBox1.Controls.Add(this.tbDate);
            this.groupBox1.Controls.Add(this.lblVehicleNumber);
            this.groupBox1.Controls.Add(this.tbQuantity);
            this.groupBox1.Controls.Add(this.lblDate);
            this.groupBox1.Controls.Add(this.tbProduct);
            this.groupBox1.Controls.Add(this.tbPort);
            this.groupBox1.Controls.Add(this.tbSupplier);
            this.groupBox1.Location = new System.Drawing.Point(12, 54);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(303, 236);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Номера колонок";
            // 
            // tbContact
            // 
            this.tbContact.Location = new System.Drawing.Point(185, 201);
            this.tbContact.Name = "tbContact";
            this.tbContact.Size = new System.Drawing.Size(81, 20);
            this.tbContact.TabIndex = 19;
            // 
            // tbTTNNumber
            // 
            this.tbTTNNumber.Location = new System.Drawing.Point(185, 176);
            this.tbTTNNumber.Name = "tbTTNNumber";
            this.tbTTNNumber.Size = new System.Drawing.Size(81, 20);
            this.tbTTNNumber.TabIndex = 18;
            // 
            // tbVehicleNumber
            // 
            this.tbVehicleNumber.Location = new System.Drawing.Point(185, 151);
            this.tbVehicleNumber.Name = "tbVehicleNumber";
            this.tbVehicleNumber.Size = new System.Drawing.Size(81, 20);
            this.tbVehicleNumber.TabIndex = 17;
            // 
            // tbDate
            // 
            this.tbDate.Location = new System.Drawing.Point(185, 127);
            this.tbDate.Name = "tbDate";
            this.tbDate.Size = new System.Drawing.Size(81, 20);
            this.tbDate.TabIndex = 16;
            // 
            // tbQuantity
            // 
            this.tbQuantity.Location = new System.Drawing.Point(185, 102);
            this.tbQuantity.Name = "tbQuantity";
            this.tbQuantity.Size = new System.Drawing.Size(81, 20);
            this.tbQuantity.TabIndex = 15;
            // 
            // tbProduct
            // 
            this.tbProduct.Location = new System.Drawing.Point(185, 76);
            this.tbProduct.Name = "tbProduct";
            this.tbProduct.Size = new System.Drawing.Size(81, 20);
            this.tbProduct.TabIndex = 14;
            // 
            // tbPort
            // 
            this.tbPort.Location = new System.Drawing.Point(185, 24);
            this.tbPort.Name = "tbPort";
            this.tbPort.Size = new System.Drawing.Size(81, 20);
            this.tbPort.TabIndex = 12;
            // 
            // tbSupplier
            // 
            this.tbSupplier.Location = new System.Drawing.Point(185, 49);
            this.tbSupplier.Name = "tbSupplier";
            this.tbSupplier.Size = new System.Drawing.Size(81, 20);
            this.tbSupplier.TabIndex = 13;
            // 
            // tbSheetNumber
            // 
            this.tbSheetNumber.Location = new System.Drawing.Point(197, 315);
            this.tbSheetNumber.Name = "tbSheetNumber";
            this.tbSheetNumber.Size = new System.Drawing.Size(81, 20);
            this.tbSheetNumber.TabIndex = 20;
            // 
            // tbStartRowNumber
            // 
            this.tbStartRowNumber.Location = new System.Drawing.Point(197, 340);
            this.tbStartRowNumber.Name = "tbStartRowNumber";
            this.tbStartRowNumber.Size = new System.Drawing.Size(81, 20);
            this.tbStartRowNumber.TabIndex = 21;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnCreateTempExcel);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.btnChooseOutputTempFolder);
            this.groupBox2.Controls.Add(this.tbOutputFolderPath);
            this.groupBox2.Controls.Add(this.lblOutputFolder);
            this.groupBox2.Controls.Add(this.lblInputFile);
            this.groupBox2.Controls.Add(this.tbInputFilePath);
            this.groupBox2.Controls.Add(this.btnChooseInputFile);
            this.groupBox2.Location = new System.Drawing.Point(356, 54);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(437, 147);
            this.groupBox2.TabIndex = 22;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Формування Excel файлу (без контрактів)";
            // 
            // btnCreateTempExcel
            // 
            this.btnCreateTempExcel.Location = new System.Drawing.Point(334, 118);
            this.btnCreateTempExcel.Name = "btnCreateTempExcel";
            this.btnCreateTempExcel.Size = new System.Drawing.Size(97, 23);
            this.btnCreateTempExcel.TabIndex = 7;
            this.btnCreateTempExcel.Text = "Сформувати";
            this.btnCreateTempExcel.UseVisualStyleBackColor = true;
            this.btnCreateTempExcel.Click += new System.EventHandler(this.btnCreateTempExcel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(85, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "(куди зберегти)";
            // 
            // btnChooseOutputTempFolder
            // 
            this.btnChooseOutputTempFolder.Location = new System.Drawing.Point(334, 56);
            this.btnChooseOutputTempFolder.Name = "btnChooseOutputTempFolder";
            this.btnChooseOutputTempFolder.Size = new System.Drawing.Size(97, 23);
            this.btnChooseOutputTempFolder.TabIndex = 5;
            this.btnChooseOutputTempFolder.Text = "Вибрати...";
            this.btnChooseOutputTempFolder.UseVisualStyleBackColor = true;
            this.btnChooseOutputTempFolder.Click += new System.EventHandler(this.btnChooseOutputTempFolder_Click);
            // 
            // tbOutputFolderPath
            // 
            this.tbOutputFolderPath.Enabled = false;
            this.tbOutputFolderPath.Location = new System.Drawing.Point(95, 59);
            this.tbOutputFolderPath.Name = "tbOutputFolderPath";
            this.tbOutputFolderPath.ReadOnly = true;
            this.tbOutputFolderPath.Size = new System.Drawing.Size(233, 20);
            this.tbOutputFolderPath.TabIndex = 4;
            // 
            // lblOutputFolder
            // 
            this.lblOutputFolder.AutoSize = true;
            this.lblOutputFolder.Location = new System.Drawing.Point(11, 66);
            this.lblOutputFolder.Name = "lblOutputFolder";
            this.lblOutputFolder.Size = new System.Drawing.Size(78, 13);
            this.lblOutputFolder.TabIndex = 3;
            this.lblOutputFolder.Text = "Вихідна папка";
            // 
            // lblInputFile
            // 
            this.lblInputFile.AutoSize = true;
            this.lblInputFile.Location = new System.Drawing.Point(15, 27);
            this.lblInputFile.Name = "lblInputFile";
            this.lblInputFile.Size = new System.Drawing.Size(74, 13);
            this.lblInputFile.TabIndex = 2;
            this.lblInputFile.Text = "Вхідний файл";
            // 
            // tbInputFilePath
            // 
            this.tbInputFilePath.Enabled = false;
            this.tbInputFilePath.Location = new System.Drawing.Point(95, 21);
            this.tbInputFilePath.Name = "tbInputFilePath";
            this.tbInputFilePath.ReadOnly = true;
            this.tbInputFilePath.Size = new System.Drawing.Size(233, 20);
            this.tbInputFilePath.TabIndex = 1;
            // 
            // btnChooseInputFile
            // 
            this.btnChooseInputFile.Location = new System.Drawing.Point(334, 19);
            this.btnChooseInputFile.Name = "btnChooseInputFile";
            this.btnChooseInputFile.Size = new System.Drawing.Size(97, 23);
            this.btnChooseInputFile.TabIndex = 0;
            this.btnChooseInputFile.Text = "Вибрати...";
            this.btnChooseInputFile.UseVisualStyleBackColor = true;
            this.btnChooseInputFile.Click += new System.EventHandler(this.btnChooseInputFile_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnCreateTemplates);
            this.groupBox3.Controls.Add(this.tbOutputTemplateFolder);
            this.groupBox3.Controls.Add(this.lblOutputTemplateFolder);
            this.groupBox3.Controls.Add(this.btnChooseTemplateFolderPath);
            this.groupBox3.Controls.Add(this.btnChooseExcelFile);
            this.groupBox3.Controls.Add(this.tbExcelFilePath);
            this.groupBox3.Controls.Add(this.lblExcelFileWithContracts);
            this.groupBox3.Location = new System.Drawing.Point(356, 212);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(437, 160);
            this.groupBox3.TabIndex = 23;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Формування темплейтів";
            // 
            // btnCreateTemplates
            // 
            this.btnCreateTemplates.Location = new System.Drawing.Point(334, 125);
            this.btnCreateTemplates.Name = "btnCreateTemplates";
            this.btnCreateTemplates.Size = new System.Drawing.Size(97, 23);
            this.btnCreateTemplates.TabIndex = 8;
            this.btnCreateTemplates.Text = "Сформувати";
            this.btnCreateTemplates.UseVisualStyleBackColor = true;
            this.btnCreateTemplates.Click += new System.EventHandler(this.btnCreateTemplates_Click);
            // 
            // tbOutputTemplateFolder
            // 
            this.tbOutputTemplateFolder.Enabled = false;
            this.tbOutputTemplateFolder.Location = new System.Drawing.Point(135, 71);
            this.tbOutputTemplateFolder.Name = "tbOutputTemplateFolder";
            this.tbOutputTemplateFolder.ReadOnly = true;
            this.tbOutputTemplateFolder.Size = new System.Drawing.Size(193, 20);
            this.tbOutputTemplateFolder.TabIndex = 12;
            // 
            // lblOutputTemplateFolder
            // 
            this.lblOutputTemplateFolder.AutoSize = true;
            this.lblOutputTemplateFolder.Location = new System.Drawing.Point(15, 76);
            this.lblOutputTemplateFolder.Name = "lblOutputTemplateFolder";
            this.lblOutputTemplateFolder.Size = new System.Drawing.Size(78, 13);
            this.lblOutputTemplateFolder.TabIndex = 11;
            this.lblOutputTemplateFolder.Text = "Вихідна папка";
            this.lblOutputTemplateFolder.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnChooseTemplateFolderPath
            // 
            this.btnChooseTemplateFolderPath.Location = new System.Drawing.Point(334, 71);
            this.btnChooseTemplateFolderPath.Name = "btnChooseTemplateFolderPath";
            this.btnChooseTemplateFolderPath.Size = new System.Drawing.Size(97, 23);
            this.btnChooseTemplateFolderPath.TabIndex = 10;
            this.btnChooseTemplateFolderPath.Text = "Вибрати...";
            this.btnChooseTemplateFolderPath.UseVisualStyleBackColor = true;
            this.btnChooseTemplateFolderPath.Click += new System.EventHandler(this.btnChooseTemplateFolderPath_Click);
            // 
            // btnChooseExcelFile
            // 
            this.btnChooseExcelFile.Location = new System.Drawing.Point(334, 32);
            this.btnChooseExcelFile.Name = "btnChooseExcelFile";
            this.btnChooseExcelFile.Size = new System.Drawing.Size(97, 23);
            this.btnChooseExcelFile.TabIndex = 9;
            this.btnChooseExcelFile.Text = "Вибрати...";
            this.btnChooseExcelFile.UseVisualStyleBackColor = true;
            this.btnChooseExcelFile.Click += new System.EventHandler(this.btnChooseExcelFile_Click);
            // 
            // tbExcelFilePath
            // 
            this.tbExcelFilePath.Enabled = false;
            this.tbExcelFilePath.Location = new System.Drawing.Point(135, 35);
            this.tbExcelFilePath.Name = "tbExcelFilePath";
            this.tbExcelFilePath.ReadOnly = true;
            this.tbExcelFilePath.Size = new System.Drawing.Size(193, 20);
            this.tbExcelFilePath.TabIndex = 8;
            // 
            // lblExcelFileWithContracts
            // 
            this.lblExcelFileWithContracts.AutoSize = true;
            this.lblExcelFileWithContracts.Location = new System.Drawing.Point(15, 42);
            this.lblExcelFileWithContracts.Name = "lblExcelFileWithContracts";
            this.lblExcelFileWithContracts.Size = new System.Drawing.Size(114, 13);
            this.lblExcelFileWithContracts.TabIndex = 0;
            this.lblExcelFileWithContracts.Text = "Файл з контрактами";
            this.lblExcelFileWithContracts.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(805, 426);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.tbStartRowNumber);
            this.Controls.Add(this.tbSheetNumber);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblStartRowNumber);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.lblSheetNumber);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainWindow";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.Label lblPort;
        private System.Windows.Forms.Label lblSupplier;
        private System.Windows.Forms.Label lblProduct;
        private System.Windows.Forms.Label lblQuantity;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.Label lblVehicleNumber;
        private System.Windows.Forms.Label lblTTNNumber;
        private System.Windows.Forms.Label lblContract;
        private System.Windows.Forms.Label lblSheetNumber;
        private System.Windows.Forms.Label lblStartRowNumber;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tbContact;
        private System.Windows.Forms.TextBox tbTTNNumber;
        private System.Windows.Forms.TextBox tbVehicleNumber;
        private System.Windows.Forms.TextBox tbDate;
        private System.Windows.Forms.TextBox tbQuantity;
        private System.Windows.Forms.TextBox tbProduct;
        private System.Windows.Forms.TextBox tbPort;
        private System.Windows.Forms.TextBox tbSupplier;
        private System.Windows.Forms.TextBox tbSheetNumber;
        private System.Windows.Forms.TextBox tbStartRowNumber;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnCreateTempExcel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnChooseOutputTempFolder;
        private System.Windows.Forms.TextBox tbOutputFolderPath;
        private System.Windows.Forms.Label lblOutputFolder;
        private System.Windows.Forms.Label lblInputFile;
        private System.Windows.Forms.TextBox tbInputFilePath;
        private System.Windows.Forms.Button btnChooseInputFile;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox tbOutputTemplateFolder;
        private System.Windows.Forms.Label lblOutputTemplateFolder;
        private System.Windows.Forms.Button btnChooseTemplateFolderPath;
        private System.Windows.Forms.Button btnChooseExcelFile;
        private System.Windows.Forms.TextBox tbExcelFilePath;
        private System.Windows.Forms.Label lblExcelFileWithContracts;
        private System.Windows.Forms.Button btnCreateTemplates;
    }
}