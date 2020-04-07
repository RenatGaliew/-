namespace Подбор_кандидатов__выгрузка_данных
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
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabelName = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.spreadsLB = new System.Windows.Forms.ListBox();
            this.downloadBTN = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.sheetsLB = new System.Windows.Forms.ListBox();
            this.sendMessage = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.pdfList = new System.Windows.Forms.ListBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.emailsListView = new System.Windows.Forms.ListView();
            this.email = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.FIO = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.status = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnConnect = new System.Windows.Forms.ToolStripButton();
            this.updateBTN = new System.Windows.Forms.ToolStripButton();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.openPDFBTN = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.dwnloadFilesBTN = new System.Windows.Forms.Button();
            this.photosListView = new System.Windows.Forms.ListView();
            this.FIO2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.dwnloadBTN = new System.Windows.Forms.Button();
            this.statusLBL = new System.Windows.Forms.Label();
            this.statusStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabelName,
            this.statusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 527);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1269, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusLabelName
            // 
            this.statusLabelName.Name = "statusLabelName";
            this.statusLabelName.Size = new System.Drawing.Size(46, 17);
            this.statusLabelName.Text = "Статус:";
            // 
            // statusLabel
            // 
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(55, 17);
            this.statusLabel.Text = "Не готов";
            // 
            // spreadsLB
            // 
            this.spreadsLB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsLB.FormattingEnabled = true;
            this.spreadsLB.Location = new System.Drawing.Point(3, 16);
            this.spreadsLB.Name = "spreadsLB";
            this.spreadsLB.Size = new System.Drawing.Size(211, 55);
            this.spreadsLB.TabIndex = 6;
            // 
            // downloadBTN
            // 
            this.downloadBTN.Location = new System.Drawing.Point(235, 35);
            this.downloadBTN.Name = "downloadBTN";
            this.downloadBTN.Size = new System.Drawing.Size(85, 55);
            this.downloadBTN.TabIndex = 7;
            this.downloadBTN.Text = "Скачать xlsx";
            this.downloadBTN.UseVisualStyleBackColor = true;
            this.downloadBTN.Click += new System.EventHandler(this.dwnloadBTN_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.spreadsLB);
            this.groupBox1.Location = new System.Drawing.Point(12, 19);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(217, 74);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Первая форма";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.sheetsLB);
            this.groupBox2.Location = new System.Drawing.Point(6, 19);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(217, 74);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Вторая форма";
            // 
            // sheetsLB
            // 
            this.sheetsLB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sheetsLB.FormattingEnabled = true;
            this.sheetsLB.Location = new System.Drawing.Point(3, 16);
            this.sheetsLB.Name = "sheetsLB";
            this.sheetsLB.Size = new System.Drawing.Size(211, 55);
            this.sheetsLB.TabIndex = 6;
            // 
            // sendMessage
            // 
            this.sendMessage.Location = new System.Drawing.Point(450, 118);
            this.sendMessage.Name = "sendMessage";
            this.sendMessage.Size = new System.Drawing.Size(106, 47);
            this.sendMessage.TabIndex = 10;
            this.sendMessage.Text = "Отправить письма";
            this.sendMessage.UseVisualStyleBackColor = true;
            this.sendMessage.Click += new System.EventHandler(this.sendMessage_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button1);
            this.groupBox3.Controls.Add(this.pdfList);
            this.groupBox3.Location = new System.Drawing.Point(326, 19);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(236, 77);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Отправляемые файлы";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(149, 16);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(81, 55);
            this.button1.TabIndex = 7;
            this.button1.Text = "Обзор";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pdfList
            // 
            this.pdfList.FormattingEnabled = true;
            this.pdfList.Location = new System.Drawing.Point(6, 27);
            this.pdfList.Name = "pdfList";
            this.pdfList.Size = new System.Drawing.Size(137, 43);
            this.pdfList.TabIndex = 6;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.emailsListView);
            this.groupBox4.Location = new System.Drawing.Point(12, 102);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(435, 394);
            this.groupBox4.TabIndex = 11;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Emails";
            // 
            // emailsListView
            // 
            this.emailsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.email,
            this.FIO,
            this.status});
            this.emailsListView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.emailsListView.FullRowSelect = true;
            this.emailsListView.GridLines = true;
            this.emailsListView.Location = new System.Drawing.Point(3, 16);
            this.emailsListView.Name = "emailsListView";
            this.emailsListView.Size = new System.Drawing.Size(429, 375);
            this.emailsListView.TabIndex = 12;
            this.emailsListView.UseCompatibleStateImageBehavior = false;
            this.emailsListView.View = System.Windows.Forms.View.Details;
            // 
            // email
            // 
            this.email.Text = "Email";
            this.email.Width = 140;
            // 
            // FIO
            // 
            this.FIO.Text = "ФИО";
            this.FIO.Width = 150;
            // 
            // status
            // 
            this.status.Text = "СТАТУС";
            this.status.Width = 130;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripSeparator2,
            this.btnConnect,
            this.updateBTN});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1269, 25);
            this.toolStrip1.TabIndex = 16;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // btnConnect
            // 
            this.btnConnect.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(93, 22);
            this.btnConnect.Text = "Подключиться";
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // updateBTN
            // 
            this.updateBTN.Image = global::Подбор_кандидатов__выгрузка_данных.Properties.Resources.update;
            this.updateBTN.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.updateBTN.Name = "updateBTN";
            this.updateBTN.Size = new System.Drawing.Size(81, 22);
            this.updateBTN.Text = "Обновить";
            this.updateBTN.Click += new System.EventHandler(this.updateBTN_Click);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.statusLBL);
            this.groupBox7.Controls.Add(this.openPDFBTN);
            this.groupBox7.Controls.Add(this.button2);
            this.groupBox7.Controls.Add(this.sendMessage);
            this.groupBox7.Controls.Add(this.groupBox4);
            this.groupBox7.Controls.Add(this.groupBox3);
            this.groupBox7.Controls.Add(this.downloadBTN);
            this.groupBox7.Controls.Add(this.groupBox1);
            this.groupBox7.Dock = System.Windows.Forms.DockStyle.Left;
            this.groupBox7.Location = new System.Drawing.Point(0, 25);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(571, 502);
            this.groupBox7.TabIndex = 17;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "ЭТАП 1";
            // 
            // openPDFBTN
            // 
            this.openPDFBTN.Location = new System.Drawing.Point(453, 408);
            this.openPDFBTN.Name = "openPDFBTN";
            this.openPDFBTN.Size = new System.Drawing.Size(112, 40);
            this.openPDFBTN.TabIndex = 13;
            this.openPDFBTN.Text = "Открыть лист беседы";
            this.openPDFBTN.UseVisualStyleBackColor = true;
            this.openPDFBTN.Click += new System.EventHandler(this.openPDFBTN_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(453, 454);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 42);
            this.button2.TabIndex = 12;
            this.button2.Text = "Отправить повторно";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.dwnloadFilesBTN);
            this.groupBox5.Controls.Add(this.photosListView);
            this.groupBox5.Controls.Add(this.dwnloadBTN);
            this.groupBox5.Controls.Add(this.groupBox2);
            this.groupBox5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox5.Location = new System.Drawing.Point(571, 25);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(698, 502);
            this.groupBox5.TabIndex = 18;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "ЭТАП 2";
            // 
            // dwnloadFilesBTN
            // 
            this.dwnloadFilesBTN.Location = new System.Drawing.Point(168, 96);
            this.dwnloadFilesBTN.Name = "dwnloadFilesBTN";
            this.dwnloadFilesBTN.Size = new System.Drawing.Size(117, 53);
            this.dwnloadFilesBTN.TabIndex = 16;
            this.dwnloadFilesBTN.Text = "Скачать все файлы пользователей";
            this.dwnloadFilesBTN.UseVisualStyleBackColor = true;
            this.dwnloadFilesBTN.Click += new System.EventHandler(this.dwnloadFilesBTN_Click);
            // 
            // photosListView
            // 
            this.photosListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.FIO2});
            this.photosListView.Location = new System.Drawing.Point(6, 96);
            this.photosListView.MultiSelect = false;
            this.photosListView.Name = "photosListView";
            this.photosListView.Size = new System.Drawing.Size(156, 397);
            this.photosListView.TabIndex = 15;
            this.photosListView.UseCompatibleStateImageBehavior = false;
            this.photosListView.View = System.Windows.Forms.View.Details;
            // 
            // FIO2
            // 
            this.FIO2.Text = "ФИО";
            this.FIO2.Width = 150;
            // 
            // dwnloadBTN
            // 
            this.dwnloadBTN.Location = new System.Drawing.Point(226, 35);
            this.dwnloadBTN.Name = "dwnloadBTN";
            this.dwnloadBTN.Size = new System.Drawing.Size(85, 55);
            this.dwnloadBTN.TabIndex = 14;
            this.dwnloadBTN.Text = "Скачать xlsx";
            this.dwnloadBTN.UseVisualStyleBackColor = true;
            this.dwnloadBTN.Click += new System.EventHandler(this.downloadBTN_Click);
            // 
            // statusLBL
            // 
            this.statusLBL.AutoSize = true;
            this.statusLBL.Location = new System.Drawing.Point(453, 392);
            this.statusLBL.Name = "statusLBL";
            this.statusLBL.Size = new System.Drawing.Size(0, 13);
            this.statusLBL.TabIndex = 14;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1269, 549);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.statusStrip1);
            this.Name = "Form1";
            this.Text = "Выгрузка данных";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabelName;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
        private System.Windows.Forms.ListBox spreadsLB;
        private System.Windows.Forms.Button downloadBTN;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox sheetsLB;
        private System.Windows.Forms.Button sendMessage;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox pdfList;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnConnect;
        private System.Windows.Forms.ToolStripButton updateBTN;
        private System.Windows.Forms.ListView emailsListView;
        private System.Windows.Forms.ColumnHeader email;
        private System.Windows.Forms.ColumnHeader FIO;
        private System.Windows.Forms.ColumnHeader status;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button openPDFBTN;
        private System.Windows.Forms.Button dwnloadBTN;
        private System.Windows.Forms.ListView photosListView;
        private System.Windows.Forms.ColumnHeader FIO2;
        private System.Windows.Forms.Button dwnloadFilesBTN;
        private System.Windows.Forms.Label statusLBL;
    }
}

