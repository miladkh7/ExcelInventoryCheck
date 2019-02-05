namespace ExcelInventoryCheck
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.clmProductCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmProductNam = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmProductNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmEntranceTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnReciveDataBase = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clmProductCode,
            this.clmProductNam,
            this.clmProductNo,
            this.clmEntranceTime});
            this.dataGridView1.Location = new System.Drawing.Point(12, 75);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(605, 372);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEndEdit);
            this.dataGridView1.CellLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellLeave);
            // 
            // clmProductCode
            // 
            this.clmProductCode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmProductCode.Frozen = true;
            this.clmProductCode.HeaderText = "Code";
            this.clmProductCode.Name = "clmProductCode";
            this.clmProductCode.Width = 141;
            // 
            // clmProductNam
            // 
            this.clmProductNam.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmProductNam.HeaderText = "Procut Name";
            this.clmProductNam.Name = "clmProductNam";
            this.clmProductNam.ReadOnly = true;
            this.clmProductNam.Width = 141;
            // 
            // clmProductNo
            // 
            this.clmProductNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmProductNo.HeaderText = "NO";
            this.clmProductNo.Name = "clmProductNo";
            this.clmProductNo.Width = 140;
            // 
            // clmEntranceTime
            // 
            this.clmEntranceTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmEntranceTime.HeaderText = "EntranceTime";
            this.clmEntranceTime.Name = "clmEntranceTime";
            this.clmEntranceTime.Width = 140;
            // 
            // btnReciveDataBase
            // 
            this.btnReciveDataBase.Location = new System.Drawing.Point(12, 27);
            this.btnReciveDataBase.Name = "btnReciveDataBase";
            this.btnReciveDataBase.Size = new System.Drawing.Size(134, 23);
            this.btnReciveDataBase.TabIndex = 1;
            this.btnReciveDataBase.Text = "ورود دیتا بیس";
            this.btnReciveDataBase.UseVisualStyleBackColor = true;
            this.btnReciveDataBase.Click += new System.EventHandler(this.btnReciveDataBase_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1022, 501);
            this.Controls.Add(this.btnReciveDataBase);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmProductCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmProductNam;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmProductNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmEntranceTime;
        private System.Windows.Forms.Button btnReciveDataBase;
    }
}

