namespace ExcelTimeCheck
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
            this.openBtn = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.sheetNameTB = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.firstCloumeTB = new System.Windows.Forms.TextBox();
            this.lastColumeTB = new System.Windows.Forms.TextBox();
            this.openPathLabel = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.saveBtn = new System.Windows.Forms.Button();
            this.outputLabel = new System.Windows.Forms.Label();
            this.outputPath = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openBtn
            // 
            this.openBtn.Font = new System.Drawing.Font("SimSun", 13F);
            this.openBtn.Location = new System.Drawing.Point(480, 25);
            this.openBtn.Name = "openBtn";
            this.openBtn.Size = new System.Drawing.Size(120, 27);
            this.openBtn.TabIndex = 0;
            this.openBtn.Text = "打开目录";
            this.openBtn.UseVisualStyleBackColor = true;
            this.openBtn.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("SimSun", 13F);
            this.label1.Location = new System.Drawing.Point(32, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "文件路径:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("SimSun", 13F);
            this.label2.Location = new System.Drawing.Point(32, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 18);
            this.label2.TabIndex = 3;
            this.label2.Text = "表名:";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // sheetNameTB
            // 
            this.sheetNameTB.Location = new System.Drawing.Point(100, 82);
            this.sheetNameTB.Name = "sheetNameTB";
            this.sheetNameTB.Size = new System.Drawing.Size(169, 21);
            this.sheetNameTB.TabIndex = 4;
            this.sheetNameTB.TextChanged += new System.EventHandler(this.sheetNameTB_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("SimSun", 13F);
            this.label3.Location = new System.Drawing.Point(32, 132);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 18);
            this.label3.TabIndex = 5;
            this.label3.Text = "首列:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("SimSun", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(32, 182);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 18);
            this.label4.TabIndex = 6;
            this.label4.Text = "尾列:";
            // 
            // firstCloumeTB
            // 
            this.firstCloumeTB.Location = new System.Drawing.Point(125, 132);
            this.firstCloumeTB.Name = "firstCloumeTB";
            this.firstCloumeTB.Size = new System.Drawing.Size(100, 21);
            this.firstCloumeTB.TabIndex = 7;
            this.firstCloumeTB.TextChanged += new System.EventHandler(this.firstCloumeTB_TextChanged);
            // 
            // lastColumeTB
            // 
            this.lastColumeTB.Location = new System.Drawing.Point(125, 182);
            this.lastColumeTB.Name = "lastColumeTB";
            this.lastColumeTB.Size = new System.Drawing.Size(100, 21);
            this.lastColumeTB.TabIndex = 8;
            this.lastColumeTB.TextChanged += new System.EventHandler(this.lastColumeTB_TextChanged);
            // 
            // openPathLabel
            // 
            this.openPathLabel.AutoSize = true;
            this.openPathLabel.Font = new System.Drawing.Font("SimSun", 9F);
            this.openPathLabel.Location = new System.Drawing.Point(130, 32);
            this.openPathLabel.Name = "openPathLabel";
            this.openPathLabel.Size = new System.Drawing.Size(0, 12);
            this.openPathLabel.TabIndex = 9;
            this.openPathLabel.Click += new System.EventHandler(this.label5_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "xlsx";
            // 
            // saveBtn
            // 
            this.saveBtn.Font = new System.Drawing.Font("SimSun", 13F);
            this.saveBtn.Location = new System.Drawing.Point(480, 232);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(120, 27);
            this.saveBtn.TabIndex = 10;
            this.saveBtn.Text = "保存目录";
            this.saveBtn.UseVisualStyleBackColor = true;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // outputLabel
            // 
            this.outputLabel.AutoSize = true;
            this.outputLabel.Font = new System.Drawing.Font("SimSun", 13F);
            this.outputLabel.Location = new System.Drawing.Point(32, 232);
            this.outputLabel.Name = "outputLabel";
            this.outputLabel.Size = new System.Drawing.Size(89, 18);
            this.outputLabel.TabIndex = 11;
            this.outputLabel.Text = "存放位置:";
            // 
            // outputPath
            // 
            this.outputPath.AutoSize = true;
            this.outputPath.Font = new System.Drawing.Font("SimSun", 9F);
            this.outputPath.Location = new System.Drawing.Point(146, 232);
            this.outputPath.Name = "outputPath";
            this.outputPath.Size = new System.Drawing.Size(0, 12);
            this.outputPath.TabIndex = 12;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("SimSun", 13F);
            this.button1.Location = new System.Drawing.Point(480, 285);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(120, 27);
            this.button1.TabIndex = 13;
            this.button1.Text = "执行";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(625, 335);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.outputPath);
            this.Controls.Add(this.outputLabel);
            this.Controls.Add(this.saveBtn);
            this.Controls.Add(this.openPathLabel);
            this.Controls.Add(this.lastColumeTB);
            this.Controls.Add(this.firstCloumeTB);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.sheetNameTB);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.openBtn);
            this.Name = "Form1";
            this.Text = "李Dog专用";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openBtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox sheetNameTB;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox firstCloumeTB;
        private System.Windows.Forms.TextBox lastColumeTB;
        private System.Windows.Forms.Label openPathLabel;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button saveBtn;
        private System.Windows.Forms.Label outputLabel;
        private System.Windows.Forms.Label outputPath;
        private System.Windows.Forms.Button button1;
    }
}

