namespace RevitAddin
{
    partial class Form_Carimbo
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
            this.btn_Mec = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtbox_Path = new System.Windows.Forms.TextBox();
            this.btn_Browser = new System.Windows.Forms.Button();
            this.btn_Plu = new System.Windows.Forms.Button();
            this.btn_Ele = new System.Windows.Forms.Button();
            this.btn_LVS = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_Mec
            // 
            this.btn_Mec.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.btn_Mec.Location = new System.Drawing.Point(12, 132);
            this.btn_Mec.Name = "btn_Mec";
            this.btn_Mec.Size = new System.Drawing.Size(96, 30);
            this.btn_Mec.TabIndex = 0;
            this.btn_Mec.Text = "MEC";
            this.btn_Mec.UseVisualStyleBackColor = true;
            this.btn_Mec.Click += new System.EventHandler(this.btn_Mec_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "Escolha o caminho do arquivo:";
            // 
            // txtbox_Path
            // 
            this.txtbox_Path.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtbox_Path.Location = new System.Drawing.Point(12, 32);
            this.txtbox_Path.Name = "txtbox_Path";
            this.txtbox_Path.ReadOnly = true;
            this.txtbox_Path.Size = new System.Drawing.Size(435, 26);
            this.txtbox_Path.TabIndex = 2;
            // 
            // btn_Browser
            // 
            this.btn_Browser.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.btn_Browser.Location = new System.Drawing.Point(453, 30);
            this.btn_Browser.Name = "btn_Browser";
            this.btn_Browser.Size = new System.Drawing.Size(96, 30);
            this.btn_Browser.TabIndex = 3;
            this.btn_Browser.Text = "Browser";
            this.btn_Browser.UseVisualStyleBackColor = true;
            this.btn_Browser.Click += new System.EventHandler(this.btn_Browser_Click);
            // 
            // btn_Plu
            // 
            this.btn_Plu.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.btn_Plu.Location = new System.Drawing.Point(114, 132);
            this.btn_Plu.Name = "btn_Plu";
            this.btn_Plu.Size = new System.Drawing.Size(96, 30);
            this.btn_Plu.TabIndex = 4;
            this.btn_Plu.Text = "PLU";
            this.btn_Plu.UseVisualStyleBackColor = true;
            this.btn_Plu.Click += new System.EventHandler(this.btn_Plu_Click);
            // 
            // btn_Ele
            // 
            this.btn_Ele.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.btn_Ele.Location = new System.Drawing.Point(216, 132);
            this.btn_Ele.Name = "btn_Ele";
            this.btn_Ele.Size = new System.Drawing.Size(96, 30);
            this.btn_Ele.TabIndex = 5;
            this.btn_Ele.Text = "ELE";
            this.btn_Ele.UseVisualStyleBackColor = true;
            this.btn_Ele.Click += new System.EventHandler(this.btn_Ele_Click);
            // 
            // btn_LVS
            // 
            this.btn_LVS.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.btn_LVS.Location = new System.Drawing.Point(318, 132);
            this.btn_LVS.Name = "btn_LVS";
            this.btn_LVS.Size = new System.Drawing.Size(96, 30);
            this.btn_LVS.TabIndex = 6;
            this.btn_LVS.Text = "LVS";
            this.btn_LVS.UseVisualStyleBackColor = true;
            this.btn_LVS.Click += new System.EventHandler(this.btn_LVS_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.btn_Cancel.Location = new System.Drawing.Point(453, 132);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(96, 30);
            this.btn_Cancel.TabIndex = 7;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // Form_Carimbo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(560, 175);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_LVS);
            this.Controls.Add(this.btn_Ele);
            this.Controls.Add(this.btn_Plu);
            this.Controls.Add(this.btn_Browser);
            this.Controls.Add(this.txtbox_Path);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_Mec);
            this.Name = "Form_Carimbo";
            this.Text = "Preenchimento Carimbo";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Mec;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtbox_Path;
        private System.Windows.Forms.Button btn_Browser;
        private System.Windows.Forms.Button btn_Plu;
        private System.Windows.Forms.Button btn_Ele;
        private System.Windows.Forms.Button btn_LVS;
        private System.Windows.Forms.Button btn_Cancel;
    }
}