namespace ExcelInventoryCheck
{
    partial class frmMain
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnReciveDataBase = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.tslCurrentStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.clmCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmProductName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmExistInventory = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmDefrence = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnClearList = new System.Windows.Forms.Button();
            this.btnCollectData = new System.Windows.Forms.Button();
            this.btnSaveResult = new System.Windows.Forms.Button();
            this.btnDeleteItem = new System.Windows.Forms.Button();
            this.clmProductCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmProductNam = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmProductNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCurrentNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmEntranceTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("B Titr", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clmProductCode,
            this.clmProductNam,
            this.clmProductNo,
            this.clmCurrentNo,
            this.clmEntranceTime});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("B Nazanin", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Location = new System.Drawing.Point(14, 3);
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("B Titr", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.Size = new System.Drawing.Size(543, 484);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEndEdit);
            this.dataGridView1.CellLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellLeave);
            this.dataGridView1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
            // 
            // btnReciveDataBase
            // 
            this.btnReciveDataBase.Font = new System.Drawing.Font("B Nazanin", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.btnReciveDataBase.Location = new System.Drawing.Point(14, 6);
            this.btnReciveDataBase.Name = "btnReciveDataBase";
            this.btnReciveDataBase.Size = new System.Drawing.Size(99, 28);
            this.btnReciveDataBase.TabIndex = 1;
            this.btnReciveDataBase.Text = "ورود دیتا بیس";
            this.btnReciveDataBase.UseVisualStyleBackColor = true;
            this.btnReciveDataBase.Click += new System.EventHandler(this.btnReciveDataBase_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tslCurrentStatus});
            this.statusStrip1.Location = new System.Drawing.Point(0, 539);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1144, 22);
            this.statusStrip1.TabIndex = 5;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // tslCurrentStatus
            // 
            this.tslCurrentStatus.Name = "tslCurrentStatus";
            this.tslCurrentStatus.Size = new System.Drawing.Size(154, 17);
            this.tslCurrentStatus.Text = "برای شروع فایلی را انتخاب کنید";
            this.tslCurrentStatus.Click += new System.EventHandler(this.toolStripStatusLabel1_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 5;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 1F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 2F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 1F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.dataGridView2, 3, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 35);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 98F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 2F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1144, 501);
            this.tableLayoutPanel1.TabIndex = 6;
            this.tableLayoutPanel1.Paint += new System.Windows.Forms.PaintEventHandler(this.tableLayoutPanel1_Paint);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("B Titr", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clmCode,
            this.clmProductName,
            this.clmName,
            this.clmExistInventory,
            this.clmDefrence});
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("B Nazanin", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView2.DefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridView2.Location = new System.Drawing.Point(585, 3);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            this.dataGridView2.Size = new System.Drawing.Size(543, 484);
            this.dataGridView2.TabIndex = 2;
            // 
            // clmCode
            // 
            this.clmCode.HeaderText = "بارکد محصول";
            this.clmCode.Name = "clmCode";
            this.clmCode.ReadOnly = true;
            // 
            // clmProductName
            // 
            this.clmProductName.HeaderText = "نام محصول";
            this.clmProductName.Name = "clmProductName";
            this.clmProductName.ReadOnly = true;
            // 
            // clmName
            // 
            this.clmName.HeaderText = "موجودی اولیه";
            this.clmName.Name = "clmName";
            this.clmName.ReadOnly = true;
            // 
            // clmExistInventory
            // 
            this.clmExistInventory.HeaderText = "تعداد شمارش شده";
            this.clmExistInventory.Name = "clmExistInventory";
            this.clmExistInventory.ReadOnly = true;
            // 
            // clmDefrence
            // 
            this.clmDefrence.HeaderText = "کسری";
            this.clmDefrence.Name = "clmDefrence";
            this.clmDefrence.ReadOnly = true;
            // 
            // btnClearList
            // 
            this.btnClearList.Font = new System.Drawing.Font("B Nazanin", 11.25F, System.Drawing.FontStyle.Bold);
            this.btnClearList.Location = new System.Drawing.Point(119, 6);
            this.btnClearList.Name = "btnClearList";
            this.btnClearList.Size = new System.Drawing.Size(99, 28);
            this.btnClearList.TabIndex = 0;
            this.btnClearList.Text = "پاک کردن لیست";
            this.btnClearList.UseVisualStyleBackColor = true;
            this.btnClearList.Click += new System.EventHandler(this.btnClearList_Click);
            // 
            // btnCollectData
            // 
            this.btnCollectData.Font = new System.Drawing.Font("B Nazanin", 11.25F, System.Drawing.FontStyle.Bold);
            this.btnCollectData.Location = new System.Drawing.Point(585, 6);
            this.btnCollectData.Name = "btnCollectData";
            this.btnCollectData.Size = new System.Drawing.Size(99, 28);
            this.btnCollectData.TabIndex = 3;
            this.btnCollectData.Text = "اعمال مقایسه";
            this.btnCollectData.UseVisualStyleBackColor = true;
            this.btnCollectData.Click += new System.EventHandler(this.btnCollectData_Click);
            // 
            // btnSaveResult
            // 
            this.btnSaveResult.Font = new System.Drawing.Font("B Nazanin", 11.25F, System.Drawing.FontStyle.Bold);
            this.btnSaveResult.Location = new System.Drawing.Point(690, 6);
            this.btnSaveResult.Name = "btnSaveResult";
            this.btnSaveResult.Size = new System.Drawing.Size(99, 28);
            this.btnSaveResult.TabIndex = 4;
            this.btnSaveResult.Text = "ذخیره مقایسه";
            this.btnSaveResult.UseVisualStyleBackColor = true;
            this.btnSaveResult.Click += new System.EventHandler(this.btnSaveResult_Click);
            // 
            // btnDeleteItem
            // 
            this.btnDeleteItem.Font = new System.Drawing.Font("B Nazanin", 11.25F, System.Drawing.FontStyle.Bold);
            this.btnDeleteItem.Location = new System.Drawing.Point(224, 6);
            this.btnDeleteItem.Name = "btnDeleteItem";
            this.btnDeleteItem.Size = new System.Drawing.Size(99, 28);
            this.btnDeleteItem.TabIndex = 7;
            this.btnDeleteItem.Text = "حذف ایتم";
            this.btnDeleteItem.UseVisualStyleBackColor = true;
            this.btnDeleteItem.Click += new System.EventHandler(this.btnDeleteItem_Click);
            // 
            // clmProductCode
            // 
            this.clmProductCode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmProductCode.Frozen = true;
            this.clmProductCode.HeaderText = "بارکد محصول";
            this.clmProductCode.Name = "clmProductCode";
            this.clmProductCode.Width = 120;
            // 
            // clmProductNam
            // 
            this.clmProductNam.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmProductNam.HeaderText = "نام محصول";
            this.clmProductNam.Name = "clmProductNam";
            this.clmProductNam.ReadOnly = true;
            this.clmProductNam.Width = 141;
            // 
            // clmProductNo
            // 
            this.clmProductNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmProductNo.HeaderText = "تعداد اولیه";
            this.clmProductNo.Name = "clmProductNo";
            this.clmProductNo.Width = 50;
            // 
            // clmCurrentNo
            // 
            this.clmCurrentNo.HeaderText = "شمارش شده";
            this.clmCurrentNo.MinimumWidth = 50;
            this.clmCurrentNo.Name = "clmCurrentNo";
            this.clmCurrentNo.ReadOnly = true;
            this.clmCurrentNo.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // clmEntranceTime
            // 
            this.clmEntranceTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.clmEntranceTime.HeaderText = "زمان شمارش";
            this.clmEntranceTime.Name = "clmEntranceTime";
            this.clmEntranceTime.Width = 140;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1144, 561);
            this.Controls.Add(this.btnDeleteItem);
            this.Controls.Add(this.btnSaveResult);
            this.Controls.Add(this.btnCollectData);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.btnClearList);
            this.Controls.Add(this.btnReciveDataBase);
            this.Controls.Add(this.statusStrip1);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "گزارشگر انبار";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.SizeChanged += new System.EventHandler(this.frmMain_SizeChanged);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnReciveDataBase;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel tslCurrentStatus;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button btnClearList;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmProductName;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmName;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmExistInventory;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmDefrence;
        private System.Windows.Forms.Button btnCollectData;
        private System.Windows.Forms.Button btnSaveResult;
        private System.Windows.Forms.Button btnDeleteItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmProductCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmProductNam;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmProductNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCurrentNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmEntranceTime;
    }
}

