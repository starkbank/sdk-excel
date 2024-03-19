namespace StarkBankExcel.Forms
{
    partial class CardPurchaseForm
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.Voided = new System.Windows.Forms.CheckBox();
            this.Denied = new System.Windows.Forms.CheckBox();
            this.OptionButtonCanceled = new System.Windows.Forms.CheckBox();
            this.OptionButtonApproved = new System.Windows.Forms.CheckBox();
            this.OptionButtonConfirmed = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.beforeInput = new System.Windows.Forms.DateTimePicker();
            this.afterInput = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.Voided);
            this.groupBox2.Controls.Add(this.Denied);
            this.groupBox2.Controls.Add(this.OptionButtonCanceled);
            this.groupBox2.Controls.Add(this.OptionButtonApproved);
            this.groupBox2.Controls.Add(this.OptionButtonConfirmed);
            this.groupBox2.Location = new System.Drawing.Point(7, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(307, 82);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Consulta por Evento";
            // 
            // Voided
            // 
            this.Voided.AutoSize = true;
            this.Voided.Location = new System.Drawing.Point(114, 43);
            this.Voided.Name = "Voided";
            this.Voided.Size = new System.Drawing.Size(74, 17);
            this.Voided.TabIndex = 15;
            this.Voided.Text = "Devolvido";
            this.Voided.UseVisualStyleBackColor = true;
            // 
            // Denied
            // 
            this.Denied.AutoSize = true;
            this.Denied.Location = new System.Drawing.Point(7, 42);
            this.Denied.Name = "Denied";
            this.Denied.Size = new System.Drawing.Size(75, 17);
            this.Denied.TabIndex = 14;
            this.Denied.Text = "Recusado";
            this.Denied.UseVisualStyleBackColor = true;
            // 
            // OptionButtonCanceled
            // 
            this.OptionButtonCanceled.AutoSize = true;
            this.OptionButtonCanceled.Location = new System.Drawing.Point(216, 22);
            this.OptionButtonCanceled.Name = "OptionButtonCanceled";
            this.OptionButtonCanceled.Size = new System.Drawing.Size(77, 17);
            this.OptionButtonCanceled.TabIndex = 13;
            this.OptionButtonCanceled.Text = "Cancelado";
            this.OptionButtonCanceled.UseVisualStyleBackColor = true;
            // 
            // OptionButtonApproved
            // 
            this.OptionButtonApproved.AutoSize = true;
            this.OptionButtonApproved.Location = new System.Drawing.Point(114, 21);
            this.OptionButtonApproved.Name = "OptionButtonApproved";
            this.OptionButtonApproved.Size = new System.Drawing.Size(72, 17);
            this.OptionButtonApproved.TabIndex = 12;
            this.OptionButtonApproved.Text = "Aprovado";
            this.OptionButtonApproved.UseVisualStyleBackColor = true;
            // 
            // OptionButtonConfirmed
            // 
            this.OptionButtonConfirmed.AutoSize = true;
            this.OptionButtonConfirmed.Location = new System.Drawing.Point(6, 20);
            this.OptionButtonConfirmed.Name = "OptionButtonConfirmed";
            this.OptionButtonConfirmed.Size = new System.Drawing.Size(79, 17);
            this.OptionButtonConfirmed.TabIndex = 11;
            this.OptionButtonConfirmed.Text = "Confirmado";
            this.OptionButtonConfirmed.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.beforeInput);
            this.groupBox1.Controls.Add(this.afterInput);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(7, 9);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(307, 97);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Consulta por Status e Data";
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
            this.button1.Location = new System.Drawing.Point(47, 212);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(222, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "Consultar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // CardPurchaseForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(320, 247);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.Name = "CardPurchaseForm";
            this.Text = "CardPurchaseForm";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker beforeInput;
        private System.Windows.Forms.DateTimePicker afterInput;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox Voided;
        private System.Windows.Forms.CheckBox Denied;
        private System.Windows.Forms.CheckBox OptionButtonCanceled;
        private System.Windows.Forms.CheckBox OptionButtonApproved;
        private System.Windows.Forms.CheckBox OptionButtonConfirmed;
    }
}