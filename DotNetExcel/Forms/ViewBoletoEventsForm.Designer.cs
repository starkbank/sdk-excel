﻿namespace StarkBankExcel.Forms
{
    partial class ViewBoletoEventsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ViewBoletoEventsForm));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.periodInput = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.beforeInput = new System.Windows.Forms.DateTimePicker();
            this.afterInput = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.OptionButtonEventCredited = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.OptionButtonEventOverdue = new System.Windows.Forms.RadioButton();
            this.OptionButtonEventCanceled = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
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
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(307, 97);
            this.groupBox1.TabIndex = 3;
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
            this.periodInput.Location = new System.Drawing.Point(102, 62);
            this.periodInput.Name = "periodInput";
            this.periodInput.Size = new System.Drawing.Size(193, 21);
            this.periodInput.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 65);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Período:";
            // 
            // beforeInput
            // 
            this.beforeInput.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.beforeInput.Location = new System.Drawing.Point(215, 20);
            this.beforeInput.Name = "beforeInput";
            this.beforeInput.Size = new System.Drawing.Size(80, 20);
            this.beforeInput.TabIndex = 7;
            // 
            // afterInput
            // 
            this.afterInput.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.afterInput.Location = new System.Drawing.Point(102, 19);
            this.afterInput.Name = "afterInput";
            this.afterInput.Size = new System.Drawing.Size(79, 20);
            this.afterInput.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(187, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(22, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "até";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Data de emissão:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(52, 177);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(222, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "Consultar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // OptionButtonEventCredited
            // 
            this.OptionButtonEventCredited.AutoSize = true;
            this.OptionButtonEventCredited.Checked = true;
            this.OptionButtonEventCredited.Location = new System.Drawing.Point(6, 19);
            this.OptionButtonEventCredited.Name = "OptionButtonEventCredited";
            this.OptionButtonEventCredited.Size = new System.Drawing.Size(117, 17);
            this.OptionButtonEventCredited.TabIndex = 6;
            this.OptionButtonEventCredited.TabStop = true;
            this.OptionButtonEventCredited.Text = "Creditado em conta";
            this.OptionButtonEventCredited.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.OptionButtonEventOverdue);
            this.groupBox2.Controls.Add(this.OptionButtonEventCanceled);
            this.groupBox2.Controls.Add(this.OptionButtonEventCredited);
            this.groupBox2.Location = new System.Drawing.Point(12, 115);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(307, 56);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Consulta por Evento";
            // 
            // OptionButtonEventOverdue
            // 
            this.OptionButtonEventOverdue.AutoSize = true;
            this.OptionButtonEventOverdue.Location = new System.Drawing.Point(138, 19);
            this.OptionButtonEventOverdue.Name = "OptionButtonEventOverdue";
            this.OptionButtonEventOverdue.Size = new System.Drawing.Size(64, 17);
            this.OptionButtonEventOverdue.TabIndex = 8;
            this.OptionButtonEventOverdue.Text = "Vencido";
            this.OptionButtonEventOverdue.UseVisualStyleBackColor = true;
            // 
            // OptionButtonEventCanceled
            // 
            this.OptionButtonEventCanceled.AutoSize = true;
            this.OptionButtonEventCanceled.Location = new System.Drawing.Point(225, 19);
            this.OptionButtonEventCanceled.Name = "OptionButtonEventCanceled";
            this.OptionButtonEventCanceled.Size = new System.Drawing.Size(76, 17);
            this.OptionButtonEventCanceled.TabIndex = 7;
            this.OptionButtonEventCanceled.Text = "Cancelado";
            this.OptionButtonEventCanceled.UseVisualStyleBackColor = true;
            // 
            // ViewChargeEventsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(329, 214);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ViewChargeEventsForm";
            this.Text = "ViewChargeEventsForm";
            this.Load += new System.EventHandler(this.ViewChargeEventsForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox periodInput;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker beforeInput;
        private System.Windows.Forms.DateTimePicker afterInput;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RadioButton OptionButtonEventCredited;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton OptionButtonEventOverdue;
        private System.Windows.Forms.RadioButton OptionButtonEventCanceled;
    }
}