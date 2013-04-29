namespace textFileReader
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
            this.listView1 = new System.Windows.Forms.ListView();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.outputText = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.AllowDrop = true;
            this.listView1.Dock = System.Windows.Forms.DockStyle.Top;
            this.listView1.Location = new System.Drawing.Point(3, 16);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(489, 192);
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.List;
            this.listView1.DragDrop += new System.Windows.Forms.DragEventHandler(this.onDragDrop);
            this.listView1.DragEnter += new System.Windows.Forms.DragEventHandler(this.onDragEnter);
            this.listView1.DragOver += new System.Windows.Forms.DragEventHandler(this.onDragOver);
            this.listView1.DragLeave += new System.EventHandler(this.onDragLeave);
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Right;
            this.button1.Location = new System.Drawing.Point(417, 208);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 25);
            this.button1.TabIndex = 1;
            this.button1.Text = "Browse Files";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Select Files to Merge";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.listView1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(495, 236);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select Input";
            // 
            // button3
            // 
            this.button3.Dock = System.Windows.Forms.DockStyle.Top;
            this.button3.Location = new System.Drawing.Point(0, 236);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(495, 23);
            this.button3.TabIndex = 6;
            this.button3.Text = "Process Files";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // outputText
            // 
            this.outputText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.outputText.Location = new System.Drawing.Point(0, 259);
            this.outputText.Multiline = true;
            this.outputText.Name = "outputText";
            this.outputText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.outputText.Size = new System.Drawing.Size(495, 341);
            this.outputText.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(495, 600);
            this.Controls.Add(this.outputText);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = ".msg file processor";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox outputText;
    }
}

