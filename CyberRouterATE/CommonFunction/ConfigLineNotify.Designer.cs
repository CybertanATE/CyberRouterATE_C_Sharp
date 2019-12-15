namespace CyberRouterATE
{
    partial class ConfigLineNotify
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.clboxConfigLineNotifyGroup = new System.Windows.Forms.CheckedListBox();
            this.cbLineTestItem = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnExport);
            this.groupBox1.Controls.Add(this.btnImport);
            this.groupBox1.Controls.Add(this.clboxConfigLineNotifyGroup);
            this.groupBox1.Controls.Add(this.cbLineTestItem);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(339, 266);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Setup";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 12);
            this.label2.TabIndex = 31;
            this.label2.Text = "Groups:";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(248, 73);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 30;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(248, 44);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 29;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // clboxConfigLineNotifyGroup
            // 
            this.clboxConfigLineNotifyGroup.CheckOnClick = true;
            this.clboxConfigLineNotifyGroup.FormattingEnabled = true;
            this.clboxConfigLineNotifyGroup.Items.AddRange(new object[] {
            "ATE Notify 1",
            "ATE Notify 2",
            "ATE Notify 3",
            "ATE Notify 4",
            "ATE Notify 5",
            "ATE Notify 6"});
            this.clboxConfigLineNotifyGroup.Location = new System.Drawing.Point(14, 97);
            this.clboxConfigLineNotifyGroup.Name = "clboxConfigLineNotifyGroup";
            this.clboxConfigLineNotifyGroup.Size = new System.Drawing.Size(213, 106);
            this.clboxConfigLineNotifyGroup.TabIndex = 29;
            this.clboxConfigLineNotifyGroup.SelectedIndexChanged += new System.EventHandler(this.clboxConfigLineNotifyGroup_SelectedIndexChanged);
            // 
            // cbLineTestItem
            // 
            this.cbLineTestItem.BackColor = System.Drawing.SystemColors.Window;
            this.cbLineTestItem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLineTestItem.FormattingEnabled = true;
            this.cbLineTestItem.Items.AddRange(new object[] {
            "USB Storage",
            "GUI test"});
            this.cbLineTestItem.Location = new System.Drawing.Point(14, 44);
            this.cbLineTestItem.Name = "cbLineTestItem";
            this.cbLineTestItem.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cbLineTestItem.Size = new System.Drawing.Size(213, 20);
            this.cbLineTestItem.TabIndex = 28;
            this.cbLineTestItem.SelectedIndexChanged += new System.EventHandler(this.cbLineTestItem_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 12);
            this.label1.TabIndex = 14;
            this.label1.Text = "Test Item:";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(115, 285);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 27;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(277, 285);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 26;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(196, 285);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 28;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // ConfigLineNotify
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(364, 333);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox1);
            this.Name = "ConfigLineNotify";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ConfigLineNotify";
            this.Load += new System.EventHandler(this.ConfigLineNotify_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbLineTestItem;
        private System.Windows.Forms.CheckedListBox clboxConfigLineNotifyGroup;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;
    }
}