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
            this.Login = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Login
            // 
            this.Login.Location = new System.Drawing.Point(112, 93);
            this.Login.Name = "Login";
            this.Login.Size = new System.Drawing.Size(258, 23);
            this.Login.TabIndex = 6;
            this.Login.Text = "Acessar Pedido";
            this.Login.UseVisualStyleBackColor = true;
            this.Login.Click += new System.EventHandler(this.Login_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(465, 17);
            this.label3.TabIndex = 7;
            this.label3.Text = "Clique no botão abaixo e acesse sua solicitação para confirmar o pedido";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Redirect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(488, 128);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Login);
            this.Name = "Redirect";
            this.Text = "Cartões Enviados";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Login;
        private System.Windows.Forms.Label label3;
    }
}