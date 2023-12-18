namespace StarkBankExcel.Forms
{
    partial class ViewTransactions
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TransactionId = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.beforeInput = new System.Windows.Forms.DateTimePicker();
            this.afterInput = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.failed = new System.Windows.Forms.RadioButton();
            this.processing = new System.Windows.Forms.RadioButton();
            this.success = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.detail = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.TransactionId);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(8, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(307, 112);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Consulta por Transação";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(91, 13);
            this.label3.TabIndex = 23;
            this.label3.Text = "Id da Transação: ";
            // 
            // TransactionId
            // 
            this.TransactionId.Location = new System.Drawing.Point(103, 68);
            this.TransactionId.Name = "TransactionId";
            this.TransactionId.Size = new System.Drawing.Size(190, 20);
            this.TransactionId.TabIndex = 22;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "transação do extrato.";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(289, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Consultar todas as transferências realizadas em uma mesma";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.beforeInput);
            this.groupBox2.Controls.Add(this.afterInput);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Location = new System.Drawing.Point(8, 130);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(307, 58);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Consulta por Status e Data";
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
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(187, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(22, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "até";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(4, 25);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Data de emissão:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.failed);
            this.groupBox3.Controls.Add(this.processing);
            this.groupBox3.Controls.Add(this.success);
            this.groupBox3.Location = new System.Drawing.Point(8, 188);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(307, 58);
            this.groupBox3.TabIndex = 11;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Status";
            // 
            // failed
            // 
            this.failed.AutoSize = true;
            this.failed.Location = new System.Drawing.Point(219, 23);
            this.failed.Name = "failed";
            this.failed.Size = new System.Drawing.Size(51, 17);
            this.failed.TabIndex = 9;
            this.failed.Text = "Falha";
            this.failed.UseVisualStyleBackColor = true;
            // 
            // processing
            // 
            this.processing.AutoSize = true;
            this.processing.Location = new System.Drawing.Point(108, 23);
            this.processing.Name = "processing";
            this.processing.Size = new System.Drawing.Size(87, 17);
            this.processing.TabIndex = 8;
            this.processing.Text = "Processando";
            this.processing.UseVisualStyleBackColor = true;
            // 
            // success
            // 
            this.success.AutoSize = true;
            this.success.Location = new System.Drawing.Point(15, 23);
            this.success.Name = "success";
            this.success.Size = new System.Drawing.Size(66, 17);
            this.success.TabIndex = 7;
            this.success.Text = "Sucesso";
            this.success.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(15, 271);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(281, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "Consultar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // detail
            // 
            this.detail.AutoSize = true;
            this.detail.Location = new System.Drawing.Point(22, 248);
            this.detail.Name = "detail";
            this.detail.Size = new System.Drawing.Size(140, 17);
            this.detail.TabIndex = 13;
            this.detail.Text = "Transferência detalhada";
            this.detail.UseVisualStyleBackColor = true;
            // 
            // ViewTransactions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(321, 306);
            this.Controls.Add(this.detail);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "ViewTransactions";
            this.Text = "ViewTransactions";
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

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TransactionId;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DateTimePicker beforeInput;
        private System.Windows.Forms.DateTimePicker afterInput;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton failed;
        private System.Windows.Forms.RadioButton processing;
        private System.Windows.Forms.RadioButton success;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RadioButton detail;
    }
}