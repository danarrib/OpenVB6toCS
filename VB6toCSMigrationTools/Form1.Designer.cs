namespace VB6toCSMigrationTools
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
            this.txtSource = new System.Windows.Forms.TextBox();
            this.txtDest = new System.Windows.Forms.TextBox();
            this.btnConvertVB6toCS = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtSource
            // 
            this.txtSource.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSource.Location = new System.Drawing.Point(12, 58);
            this.txtSource.Multiline = true;
            this.txtSource.Name = "txtSource";
            this.txtSource.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtSource.Size = new System.Drawing.Size(424, 317);
            this.txtSource.TabIndex = 0;
            // 
            // txtDest
            // 
            this.txtDest.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDest.Location = new System.Drawing.Point(574, 58);
            this.txtDest.Multiline = true;
            this.txtDest.Name = "txtDest";
            this.txtDest.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtDest.Size = new System.Drawing.Size(424, 317);
            this.txtDest.TabIndex = 1;
            // 
            // btnConvertVB6toCS
            // 
            this.btnConvertVB6toCS.Location = new System.Drawing.Point(442, 58);
            this.btnConvertVB6toCS.Name = "btnConvertVB6toCS";
            this.btnConvertVB6toCS.Size = new System.Drawing.Size(126, 23);
            this.btnConvertVB6toCS.TabIndex = 2;
            this.btnConvertVB6toCS.Text = "Convert VB6 to C#  >>>";
            this.btnConvertVB6toCS.UseVisualStyleBackColor = true;
            this.btnConvertVB6toCS.Click += new System.EventHandler(this.btnConvertVB6toCS_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1014, 441);
            this.Controls.Add(this.txtDest);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.btnConvertVB6toCS);
            this.Name = "Form1";
            this.Text = "VB6 to C# Converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.TextBox txtDest;
        private System.Windows.Forms.Button btnConvertVB6toCS;
    }
}

