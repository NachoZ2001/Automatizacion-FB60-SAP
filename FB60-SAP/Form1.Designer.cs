namespace FB60_SAP
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            textBoxRutaExcel = new TextBox();
            textBoxTextoExcel = new TextBox();
            buttonSeleccionarExcel = new Button();
            buttonEjecutarScript = new Button();
            buttonTransformarExcel = new Button();
            SuspendLayout();
            // 
            // textBoxRutaExcel
            // 
            textBoxRutaExcel.Location = new Point(12, 44);
            textBoxRutaExcel.Name = "textBoxRutaExcel";
            textBoxRutaExcel.Size = new Size(294, 23);
            textBoxRutaExcel.TabIndex = 0;
            // 
            // textBoxTextoExcel
            // 
            textBoxTextoExcel.BackColor = Color.Purple;
            textBoxTextoExcel.BorderStyle = BorderStyle.None;
            textBoxTextoExcel.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            textBoxTextoExcel.ForeColor = Color.White;
            textBoxTextoExcel.Location = new Point(12, 20);
            textBoxTextoExcel.Name = "textBoxTextoExcel";
            textBoxTextoExcel.Size = new Size(100, 18);
            textBoxTextoExcel.TabIndex = 1;
            textBoxTextoExcel.Text = "Ruta Excel";
            // 
            // buttonSeleccionarExcel
            // 
            buttonSeleccionarExcel.BackColor = Color.BlueViolet;
            buttonSeleccionarExcel.FlatStyle = FlatStyle.Popup;
            buttonSeleccionarExcel.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            buttonSeleccionarExcel.ForeColor = Color.White;
            buttonSeleccionarExcel.Location = new Point(12, 83);
            buttonSeleccionarExcel.Name = "buttonSeleccionarExcel";
            buttonSeleccionarExcel.Size = new Size(294, 41);
            buttonSeleccionarExcel.TabIndex = 2;
            buttonSeleccionarExcel.Text = "Seleccionar Excel";
            buttonSeleccionarExcel.UseVisualStyleBackColor = false;
            buttonSeleccionarExcel.Click += buttonSeleccionarExcel_Click;
            // 
            // buttonEjecutarScript
            // 
            buttonEjecutarScript.BackColor = Color.BlueViolet;
            buttonEjecutarScript.FlatStyle = FlatStyle.Popup;
            buttonEjecutarScript.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            buttonEjecutarScript.ForeColor = Color.White;
            buttonEjecutarScript.Location = new Point(12, 209);
            buttonEjecutarScript.Name = "buttonEjecutarScript";
            buttonEjecutarScript.Size = new Size(294, 41);
            buttonEjecutarScript.TabIndex = 3;
            buttonEjecutarScript.Text = "Ejecutar Script";
            buttonEjecutarScript.UseVisualStyleBackColor = false;
            buttonEjecutarScript.Click += buttonEjecutarScript_Click;
            // 
            // buttonTransformarExcel
            // 
            buttonTransformarExcel.BackColor = Color.BlueViolet;
            buttonTransformarExcel.FlatStyle = FlatStyle.Popup;
            buttonTransformarExcel.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            buttonTransformarExcel.ForeColor = Color.White;
            buttonTransformarExcel.Location = new Point(12, 146);
            buttonTransformarExcel.Name = "buttonTransformarExcel";
            buttonTransformarExcel.Size = new Size(294, 41);
            buttonTransformarExcel.TabIndex = 4;
            buttonTransformarExcel.Text = "Transformar Excel";
            buttonTransformarExcel.UseVisualStyleBackColor = false;
            buttonTransformarExcel.Click += buttonTransformarExcel_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Purple;
            ClientSize = new Size(325, 262);
            Controls.Add(buttonTransformarExcel);
            Controls.Add(buttonEjecutarScript);
            Controls.Add(buttonSeleccionarExcel);
            Controls.Add(textBoxTextoExcel);
            Controls.Add(textBoxRutaExcel);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox textBoxRutaExcel;
        private TextBox textBoxTextoExcel;
        private Button buttonSeleccionarExcel;
        private Button buttonEjecutarScript;
        private Button buttonTransformarExcel;
    }
}
