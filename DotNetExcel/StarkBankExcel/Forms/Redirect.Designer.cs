namespace StarkBankExcel.Forms
{
    partial class Redirect
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
            this.ShopClick = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ShopClick
            // 
            this.ShopClick.Location = new System.Drawing.Point(5, 128);
            this.ShopClick.Name = "ShopClick";
            this.ShopClick.Size = new System.Drawing.Size(258, 23);
            this.ShopClick.TabIndex = 9;
            this.ShopClick.Text = "Acessar Carrinho";
            this.ShopClick.UseVisualStyleBackColor = true;
            this.ShopClick.Click += new System.EventHandler(this.ShopClick_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.label4.Location = new System.Drawing.Point(54, 69);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(167, 24);
            this.label4.TabIndex = 10;
            this.label4.Text = "Cartões Enviados !";
            // 
            // Redirect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(275, 163);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.ShopClick);
            this.Name = "Redirect";
            this.Text = "Redirect";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button ShopClick;
        private System.Windows.Forms.Label label4;
    }
}