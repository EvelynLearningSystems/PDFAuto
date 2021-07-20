
namespace PDFAuto
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
            this.loadDocBtn = new System.Windows.Forms.Button();
            this.contentBox = new System.Windows.Forms.TextBox();
            this.headingsBox = new System.Windows.Forms.TextBox();
            this.allHeadingsBox = new System.Windows.Forms.TextBox();
            this.questionsBox = new System.Windows.Forms.TextBox();
            this.testButton = new System.Windows.Forms.Button();
            this.NewHeadingsBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.loadDocFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.fileBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // loadDocBtn
            // 
            this.loadDocBtn.Location = new System.Drawing.Point(12, 12);
            this.loadDocBtn.Name = "loadDocBtn";
            this.loadDocBtn.Size = new System.Drawing.Size(117, 23);
            this.loadDocBtn.TabIndex = 0;
            this.loadDocBtn.Text = "Load Document";
            this.loadDocBtn.UseVisualStyleBackColor = true;
            this.loadDocBtn.Click += new System.EventHandler(this.loadDocBtn_Click);
            // 
            // contentBox
            // 
            this.contentBox.Location = new System.Drawing.Point(239, 46);
            this.contentBox.Multiline = true;
            this.contentBox.Name = "contentBox";
            this.contentBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.contentBox.Size = new System.Drawing.Size(549, 122);
            this.contentBox.TabIndex = 1;
            this.contentBox.WordWrap = false;
            // 
            // headingsBox
            // 
            this.headingsBox.Location = new System.Drawing.Point(172, 365);
            this.headingsBox.Multiline = true;
            this.headingsBox.Name = "headingsBox";
            this.headingsBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.headingsBox.Size = new System.Drawing.Size(297, 140);
            this.headingsBox.TabIndex = 2;
            this.headingsBox.WordWrap = false;
            // 
            // allHeadingsBox
            // 
            this.allHeadingsBox.Location = new System.Drawing.Point(491, 365);
            this.allHeadingsBox.Multiline = true;
            this.allHeadingsBox.Name = "allHeadingsBox";
            this.allHeadingsBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.allHeadingsBox.Size = new System.Drawing.Size(297, 140);
            this.allHeadingsBox.TabIndex = 3;
            this.allHeadingsBox.WordWrap = false;
            // 
            // questionsBox
            // 
            this.questionsBox.Location = new System.Drawing.Point(239, 174);
            this.questionsBox.Multiline = true;
            this.questionsBox.Name = "questionsBox";
            this.questionsBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.questionsBox.Size = new System.Drawing.Size(549, 160);
            this.questionsBox.TabIndex = 4;
            this.questionsBox.WordWrap = false;
            // 
            // testButton
            // 
            this.testButton.Location = new System.Drawing.Point(13, 85);
            this.testButton.Name = "testButton";
            this.testButton.Size = new System.Drawing.Size(116, 23);
            this.testButton.TabIndex = 5;
            this.testButton.Text = "Test";
            this.testButton.UseVisualStyleBackColor = true;
            this.testButton.Click += new System.EventHandler(this.testButton_Click);
            // 
            // NewHeadingsBox
            // 
            this.NewHeadingsBox.Location = new System.Drawing.Point(12, 144);
            this.NewHeadingsBox.Multiline = true;
            this.NewHeadingsBox.Name = "NewHeadingsBox";
            this.NewHeadingsBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.NewHeadingsBox.Size = new System.Drawing.Size(203, 190);
            this.NewHeadingsBox.TabIndex = 6;
            this.NewHeadingsBox.WordWrap = false;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(10, 115);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(203, 26);
            this.label1.TabIndex = 7;
            this.label1.Text = "Headings: (Write each question heading in a new line)";
            // 
            // fileBox
            // 
            this.fileBox.Location = new System.Drawing.Point(239, 13);
            this.fileBox.Name = "fileBox";
            this.fileBox.Size = new System.Drawing.Size(549, 20);
            this.fileBox.TabIndex = 8;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(824, 347);
            this.Controls.Add(this.fileBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.NewHeadingsBox);
            this.Controls.Add(this.testButton);
            this.Controls.Add(this.questionsBox);
            this.Controls.Add(this.allHeadingsBox);
            this.Controls.Add(this.headingsBox);
            this.Controls.Add(this.contentBox);
            this.Controls.Add(this.loadDocBtn);
            this.MaximumSize = new System.Drawing.Size(840, 386);
            this.MinimumSize = new System.Drawing.Size(840, 386);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PDF Automation 0.1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button loadDocBtn;
        private System.Windows.Forms.TextBox contentBox;
        private System.Windows.Forms.TextBox headingsBox;
        private System.Windows.Forms.TextBox allHeadingsBox;
        private System.Windows.Forms.TextBox questionsBox;
        private System.Windows.Forms.Button testButton;
        private System.Windows.Forms.TextBox NewHeadingsBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog loadDocFileDialog;
        private System.Windows.Forms.TextBox fileBox;
    }
}

