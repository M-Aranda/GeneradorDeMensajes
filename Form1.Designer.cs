namespace GeneradorDeMensajes
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
            this.slcExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // slcExcel
            // 
            this.slcExcel.Location = new System.Drawing.Point(339, 173);
            this.slcExcel.Name = "slcExcel";
            this.slcExcel.Size = new System.Drawing.Size(141, 53);
            this.slcExcel.TabIndex = 0;
            this.slcExcel.Text = "Seleccionar Excel";
            this.slcExcel.UseVisualStyleBackColor = true;
            this.slcExcel.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.slcExcel);
            this.Name = "Form1";
            this.Text = "Generador de mensajes";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button slcExcel;
    }
}

