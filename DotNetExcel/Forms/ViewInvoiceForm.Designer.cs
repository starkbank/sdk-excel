﻿namespace StarkBankExcel.Forms
{
    partial class ViewInvoiceForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ViewInvoiceForm));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.periodInput = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.beforeInput = new System.Windows.Forms.DateTimePicker();
            this.afterInput = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.statusInput = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.periodInput);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.beforeInput);
            this.groupBox1.Controls.Add(this.afterInput);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.statusInput);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(307, 155);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Consulta por Status e Data";
            // 
            // periodInput
            // 
            this.periodInput.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.periodInput.FormattingEnabled = true;
            this.periodInput.Items.AddRange(new object[] {
            "Data Inicial",
            "Data Final",
            "Intervalo",
            "Todos"});
            this.periodInput.Location = new System.Drawing.Point(104, 107);
            this.periodInput.Name = "periodInput";
            this.periodInput.Size = new System.Drawing.Size(193, 21);
            this.periodInput.TabIndex = 9;
            this.periodInput.SelectedIndexChanged += new System.EventHandler(this.periodInput_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 110);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Período:";
            // 
            // beforeInput
            // 
            this.beforeInput.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.beforeInput.Location = new System.Drawing.Point(217, 65);
            this.beforeInput.Name = "beforeInput";
            this.beforeInput.Size = new System.Drawing.Size(80, 20);
            this.beforeInput.TabIndex = 7;
            // 
            // afterInput
            // 
            this.afterInput.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.afterInput.Location = new System.Drawing.Point(104, 64);
            this.afterInput.Name = "afterInput";
            this.afterInput.Size = new System.Drawing.Size(79, 20);
            this.afterInput.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(189, 70);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(22, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "até";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Data de emissão:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Status:";
            // 
            // statusInput
            // 
            this.statusInput.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.statusInput.FormattingEnabled = true;
            this.statusInput.Location = new System.Drawing.Point(104, 24);
            this.statusInput.Name = "statusInput";
            this.statusInput.Size = new System.Drawing.Size(193, 21);
            this.statusInput.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(51, 173);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(222, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Consultar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // ViewInvoiceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(329, 202);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ViewInvoiceForm";
            this.Text = "Consulta de Faturas Pix emitidas";
            this.Load += new System.EventHandler(this.ViewInvoiceForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker beforeInput;
        private System.Windows.Forms.DateTimePicker afterInput;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox statusInput;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox periodInput;
        private System.Windows.Forms.Label label4;
    }
}