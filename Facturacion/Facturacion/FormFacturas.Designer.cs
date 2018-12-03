namespace Facturacion
{
    partial class FormFacturas
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormFacturas));
            this.buttonSeleccion_Archivo = new System.Windows.Forms.Button();
            this.buttonLeer_Contenido = new System.Windows.Forms.Button();
            this.buttonGenera_Excel = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonSeleccion_Archivo
            // 
            this.buttonSeleccion_Archivo.BackColor = System.Drawing.SystemColors.Window;
            this.buttonSeleccion_Archivo.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.buttonSeleccion_Archivo.FlatAppearance.BorderSize = 0;
            this.buttonSeleccion_Archivo.FlatAppearance.MouseDownBackColor = System.Drawing.Color.NavajoWhite;
            this.buttonSeleccion_Archivo.FlatAppearance.MouseOverBackColor = System.Drawing.Color.AntiqueWhite;
            this.buttonSeleccion_Archivo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSeleccion_Archivo.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSeleccion_Archivo.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonSeleccion_Archivo.Image = ((System.Drawing.Image)(resources.GetObject("buttonSeleccion_Archivo.Image")));
            this.buttonSeleccion_Archivo.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.buttonSeleccion_Archivo.Location = new System.Drawing.Point(12, 75);
            this.buttonSeleccion_Archivo.Name = "buttonSeleccion_Archivo";
            this.buttonSeleccion_Archivo.Size = new System.Drawing.Size(343, 57);
            this.buttonSeleccion_Archivo.TabIndex = 0;
            this.buttonSeleccion_Archivo.Text = "Buscar archivo\r\n";
            this.buttonSeleccion_Archivo.UseVisualStyleBackColor = false;
            this.buttonSeleccion_Archivo.Click += new System.EventHandler(this.buttonSeleccion_Archivo_Click);
            // 
            // buttonLeer_Contenido
            // 
            this.buttonLeer_Contenido.BackColor = System.Drawing.SystemColors.Window;
            this.buttonLeer_Contenido.FlatAppearance.BorderSize = 0;
            this.buttonLeer_Contenido.FlatAppearance.MouseDownBackColor = System.Drawing.Color.NavajoWhite;
            this.buttonLeer_Contenido.FlatAppearance.MouseOverBackColor = System.Drawing.Color.AntiqueWhite;
            this.buttonLeer_Contenido.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonLeer_Contenido.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonLeer_Contenido.Image = ((System.Drawing.Image)(resources.GetObject("buttonLeer_Contenido.Image")));
            this.buttonLeer_Contenido.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonLeer_Contenido.Location = new System.Drawing.Point(10, 174);
            this.buttonLeer_Contenido.Name = "buttonLeer_Contenido";
            this.buttonLeer_Contenido.Size = new System.Drawing.Size(345, 57);
            this.buttonLeer_Contenido.TabIndex = 1;
            this.buttonLeer_Contenido.Text = "Marcar Registros";
            this.buttonLeer_Contenido.UseVisualStyleBackColor = false;
            this.buttonLeer_Contenido.Click += new System.EventHandler(this.buttonLeer_Contenido_Click);
            // 
            // buttonGenera_Excel
            // 
            this.buttonGenera_Excel.BackColor = System.Drawing.SystemColors.Window;
            this.buttonGenera_Excel.FlatAppearance.BorderSize = 0;
            this.buttonGenera_Excel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.NavajoWhite;
            this.buttonGenera_Excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.AntiqueWhite;
            this.buttonGenera_Excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonGenera_Excel.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonGenera_Excel.Image = ((System.Drawing.Image)(resources.GetObject("buttonGenera_Excel.Image")));
            this.buttonGenera_Excel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonGenera_Excel.Location = new System.Drawing.Point(13, 269);
            this.buttonGenera_Excel.Name = "buttonGenera_Excel";
            this.buttonGenera_Excel.Size = new System.Drawing.Size(342, 57);
            this.buttonGenera_Excel.TabIndex = 2;
            this.buttonGenera_Excel.Text = "Generación de Factura";
            this.buttonGenera_Excel.UseVisualStyleBackColor = false;
            this.buttonGenera_Excel.Click += new System.EventHandler(this.buttonGenera_Excel_Click);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("Franklin Gothic Book", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.Color.SteelBlue;
            this.textBox1.Location = new System.Drawing.Point(360, 391);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(390, 31);
            this.textBox1.TabIndex = 3;
            this.textBox1.Text = "Sistema de Facturación P&G";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(361, 64);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(418, 288);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // FormFacturas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(801, 467);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.buttonGenera_Excel);
            this.Controls.Add(this.buttonLeer_Contenido);
            this.Controls.Add(this.buttonSeleccion_Archivo);
            this.Name = "FormFacturas";
            this.Text = "Sistema de Facturación P&G";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonSeleccion_Archivo;
        private System.Windows.Forms.Button buttonLeer_Contenido;
        private System.Windows.Forms.Button buttonGenera_Excel;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

