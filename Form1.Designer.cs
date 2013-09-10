namespace Excel
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnInput = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.btnExport = new System.Windows.Forms.Button();
            this.btncs = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmiexplain = new System.Windows.Forms.ToolStripMenuItem();
            this.tmsiHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiGuanyu = new System.Windows.Forms.ToolStripMenuItem();
            this.lblOpen = new System.Windows.Forms.Label();
            this.lblSave = new System.Windows.Forms.Label();
            this.lblState = new System.Windows.Forms.Label();
            this.lblexplain = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnInput
            // 
            this.btnInput.Location = new System.Drawing.Point(14, 37);
            this.btnInput.Name = "btnInput";
            this.btnInput.Size = new System.Drawing.Size(75, 23);
            this.btnInput.TabIndex = 0;
            this.btnInput.Text = "导入Excel";
            this.btnInput.UseVisualStyleBackColor = true;
            this.btnInput.Click += new System.EventHandler(this.btnInput_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(0, 122);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.ReadOnly = true;
            this.dataGridView.RowTemplate.Height = 23;
            this.dataGridView.Size = new System.Drawing.Size(700, 482);
            this.dataGridView.TabIndex = 1;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(125, 37);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "转换并导出Excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btncs
            // 
            this.btncs.Location = new System.Drawing.Point(236, 36);
            this.btncs.Name = "btncs";
            this.btncs.Size = new System.Drawing.Size(75, 23);
            this.btncs.TabIndex = 3;
            this.btncs.Text = "仅转换";
            this.btncs.UseVisualStyleBackColor = true;
            this.btncs.Click += new System.EventHandler(this.btncs_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiexplain,
            this.tmsiHelp,
            this.tsmiGuanyu});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(700, 24);
            this.menuStrip1.TabIndex = 4;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmiexplain
            // 
            this.tsmiexplain.Name = "tsmiexplain";
            this.tsmiexplain.Size = new System.Drawing.Size(41, 20);
            this.tsmiexplain.Text = "说明";
            this.tsmiexplain.Click += new System.EventHandler(this.tsmiexplain_Click);
            // 
            // tmsiHelp
            // 
            this.tmsiHelp.Name = "tmsiHelp";
            this.tmsiHelp.Size = new System.Drawing.Size(41, 20);
            this.tmsiHelp.Text = "帮助";
            this.tmsiHelp.Click += new System.EventHandler(this.tmsiHelp_Click);
            // 
            // tsmiGuanyu
            // 
            this.tsmiGuanyu.Name = "tsmiGuanyu";
            this.tsmiGuanyu.Size = new System.Drawing.Size(41, 20);
            this.tsmiGuanyu.Text = "关于";
            this.tsmiGuanyu.Click += new System.EventHandler(this.tsmiGuanyu_Click);
            // 
            // lblOpen
            // 
            this.lblOpen.AutoSize = true;
            this.lblOpen.Location = new System.Drawing.Point(22, 70);
            this.lblOpen.Name = "lblOpen";
            this.lblOpen.Size = new System.Drawing.Size(89, 12);
            this.lblOpen.TabIndex = 5;
            this.lblOpen.Text = "导入文件路径：";
            // 
            // lblSave
            // 
            this.lblSave.AutoSize = true;
            this.lblSave.Location = new System.Drawing.Point(22, 93);
            this.lblSave.Name = "lblSave";
            this.lblSave.Size = new System.Drawing.Size(89, 12);
            this.lblSave.TabIndex = 6;
            this.lblSave.Text = "导出文件路径：";
            // 
            // lblState
            // 
            this.lblState.AutoSize = true;
            this.lblState.ForeColor = System.Drawing.Color.Red;
            this.lblState.Location = new System.Drawing.Point(336, 55);
            this.lblState.Name = "lblState";
            this.lblState.Size = new System.Drawing.Size(41, 12);
            this.lblState.TabIndex = 7;
            this.lblState.Text = "状态：";
            // 
            // lblexplain
            // 
            this.lblexplain.AutoSize = true;
            this.lblexplain.ForeColor = System.Drawing.Color.Red;
            this.lblexplain.Location = new System.Drawing.Point(336, 34);
            this.lblexplain.Name = "lblexplain";
            this.lblexplain.Size = new System.Drawing.Size(41, 12);
            this.lblexplain.TabIndex = 8;
            this.lblexplain.Text = "说明：";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(700, 594);
            this.Controls.Add(this.lblexplain);
            this.Controls.Add(this.lblState);
            this.Controls.Add(this.lblSave);
            this.Controls.Add(this.lblOpen);
            this.Controls.Add(this.btncs);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.btnInput);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "流程环节配置转换工具";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnInput;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btncs;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmiexplain;
        private System.Windows.Forms.ToolStripMenuItem tmsiHelp;
        private System.Windows.Forms.Label lblOpen;
        private System.Windows.Forms.Label lblSave;
        private System.Windows.Forms.Label lblState;
        private System.Windows.Forms.Label lblexplain;
        private System.Windows.Forms.ToolStripMenuItem tsmiGuanyu;
    }
}

