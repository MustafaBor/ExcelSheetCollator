namespace Excel_Sheet_Collator
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
            this.groupBoxFolder = new System.Windows.Forms.GroupBox();
            this.listBoxFiles = new System.Windows.Forms.ListBox();
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.textBoxSearch = new System.Windows.Forms.TextBox();
            this.buttonSearch = new System.Windows.Forms.Button();
            this.groupBoxUpdated = new System.Windows.Forms.GroupBox();
            this.listBoxNewFiles = new System.Windows.Forms.ListBox();
            this.buttonSelectFolder = new System.Windows.Forms.Button();
            this.groupBoxFolder.SuspendLayout();
            this.groupBoxUpdated.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxFolder
            // 
            this.groupBoxFolder.Controls.Add(this.listBoxFiles);
            this.groupBoxFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBoxFolder.Location = new System.Drawing.Point(20, 7);
            this.groupBoxFolder.Name = "groupBoxFolder";
            this.groupBoxFolder.Size = new System.Drawing.Size(399, 369);
            this.groupBoxFolder.TabIndex = 0;
            this.groupBoxFolder.TabStop = false;
            this.groupBoxFolder.Text = "Select Project";
            // 
            // listBoxFiles
            // 
            this.listBoxFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxFiles.FormattingEnabled = true;
            this.listBoxFiles.ItemHeight = 18;
            this.listBoxFiles.Location = new System.Drawing.Point(10, 27);
            this.listBoxFiles.Name = "listBoxFiles";
            this.listBoxFiles.Size = new System.Drawing.Size(378, 328);
            this.listBoxFiles.TabIndex = 0;
            this.listBoxFiles.SelectedIndexChanged += new System.EventHandler(this.listBoxFiles_SelectedIndexChanged);
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.buttonRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRefresh.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonRefresh.Location = new System.Drawing.Point(269, 383);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(84, 30);
            this.buttonRefresh.TabIndex = 1;
            this.buttonRefresh.Text = "Load";
            this.buttonRefresh.UseVisualStyleBackColor = false;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // textBoxSearch
            // 
            this.textBoxSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.textBoxSearch.Location = new System.Drawing.Point(515, 382);
            this.textBoxSearch.Name = "textBoxSearch";
            this.textBoxSearch.Size = new System.Drawing.Size(144, 26);
            this.textBoxSearch.TabIndex = 2;
            // 
            // buttonSearch
            // 
            this.buttonSearch.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.buttonSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSearch.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonSearch.Location = new System.Drawing.Point(683, 382);
            this.buttonSearch.Name = "buttonSearch";
            this.buttonSearch.Size = new System.Drawing.Size(84, 30);
            this.buttonSearch.TabIndex = 3;
            this.buttonSearch.Text = "Search";
            this.buttonSearch.UseVisualStyleBackColor = false;
            this.buttonSearch.Click += new System.EventHandler(this.buttonSearch_Click);
            // 
            // groupBoxUpdated
            // 
            this.groupBoxUpdated.Controls.Add(this.listBoxNewFiles);
            this.groupBoxUpdated.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBoxUpdated.Location = new System.Drawing.Point(438, 7);
            this.groupBoxUpdated.Name = "groupBoxUpdated";
            this.groupBoxUpdated.Size = new System.Drawing.Size(399, 369);
            this.groupBoxUpdated.TabIndex = 1;
            this.groupBoxUpdated.TabStop = false;
            this.groupBoxUpdated.Text = "Search Result";
            // 
            // listBoxNewFiles
            // 
            this.listBoxNewFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxNewFiles.FormattingEnabled = true;
            this.listBoxNewFiles.ItemHeight = 18;
            this.listBoxNewFiles.Location = new System.Drawing.Point(11, 27);
            this.listBoxNewFiles.Name = "listBoxNewFiles";
            this.listBoxNewFiles.Size = new System.Drawing.Size(378, 328);
            this.listBoxNewFiles.TabIndex = 0;
            this.listBoxNewFiles.SelectedIndexChanged += new System.EventHandler(this.listBoxNewFiles_SelectedIndexChanged);
            // 
            // buttonSelectFolder
            // 
            this.buttonSelectFolder.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.buttonSelectFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSelectFolder.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonSelectFolder.Location = new System.Drawing.Point(72, 382);
            this.buttonSelectFolder.Name = "buttonSelectFolder";
            this.buttonSelectFolder.Size = new System.Drawing.Size(131, 30);
            this.buttonSelectFolder.TabIndex = 5;
            this.buttonSelectFolder.Text = "Select Folder";
            this.buttonSelectFolder.UseVisualStyleBackColor = false;
            this.buttonSelectFolder.Click += new System.EventHandler(this.buttonSelectFolder_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(860, 425);
            this.Controls.Add(this.buttonSelectFolder);
            this.Controls.Add(this.groupBoxUpdated);
            this.Controls.Add(this.buttonSearch);
            this.Controls.Add(this.textBoxSearch);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.groupBoxFolder);
            this.Name = "Form1";
            this.Text = "Excel Sheet Collator";
            this.groupBoxFolder.ResumeLayout(false);
            this.groupBoxUpdated.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxFolder;
        private System.Windows.Forms.ListBox listBoxFiles;
        private System.Windows.Forms.Button buttonRefresh;
        private System.Windows.Forms.TextBox textBoxSearch;
        private System.Windows.Forms.Button buttonSearch;
        private System.Windows.Forms.GroupBox groupBoxUpdated;
        private System.Windows.Forms.ListBox listBoxNewFiles;
        private System.Windows.Forms.Button buttonSelectFolder;
    }
}

