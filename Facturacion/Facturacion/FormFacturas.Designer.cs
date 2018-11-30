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
            this.buttonSeleccion_Archivo = new System.Windows.Forms.Button();
            this.buttonLeer_Contenido = new System.Windows.Forms.Button();
            this.buttonGenera_Excel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonSeleccion_Archivo
            // 
            this.buttonSeleccion_Archivo.Location = new System.Drawing.Point(64, 30);
            this.buttonSeleccion_Archivo.Name = "buttonSeleccion_Archivo";
            this.buttonSeleccion_Archivo.Size = new System.Drawing.Size(188, 23);
            this.buttonSeleccion_Archivo.TabIndex = 0;
            this.buttonSeleccion_Archivo.Text = "Selecciona el archivo a leer";
            this.buttonSeleccion_Archivo.UseVisualStyleBackColor = true;
            this.buttonSeleccion_Archivo.Click += new System.EventHandler(this.buttonSeleccion_Archivo_Click);
            // 
            // buttonLeer_Contenido
            // 
            this.buttonLeer_Contenido.Location = new System.Drawing.Point(320, 30);
            this.buttonLeer_Contenido.Name = "buttonLeer_Contenido";
            this.buttonLeer_Contenido.Size = new System.Drawing.Size(220, 23);
            this.buttonLeer_Contenido.TabIndex = 1;
            this.buttonLeer_Contenido.Text = "Lee el contenido del archivo seleccionado";
            this.buttonLeer_Contenido.UseVisualStyleBackColor = true;
            this.buttonLeer_Contenido.Click += new System.EventHandler(this.buttonLeer_Contenido_Click);
            // 
            // buttonGenera_Excel
            // 
            this.buttonGenera_Excel.Location = new System.Drawing.Point(629, 30);
            this.buttonGenera_Excel.Name = "buttonGenera_Excel";
            this.buttonGenera_Excel.Size = new System.Drawing.Size(150, 23);
            this.buttonGenera_Excel.TabIndex = 2;
            this.buttonGenera_Excel.Text = "Genera el excel completo";
            this.buttonGenera_Excel.UseVisualStyleBackColor = true;
            this.buttonGenera_Excel.Click += new System.EventHandler(this.buttonGenera_Excel_Click);
            // 
            // FormFacturas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(866, 226);
            this.Controls.Add(this.buttonGenera_Excel);
            this.Controls.Add(this.buttonLeer_Contenido);
            this.Controls.Add(this.buttonSeleccion_Archivo);
            this.Name = "FormFacturas";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonSeleccion_Archivo;
        private System.Windows.Forms.Button buttonLeer_Contenido;
        private System.Windows.Forms.Button buttonGenera_Excel;
    }
}

