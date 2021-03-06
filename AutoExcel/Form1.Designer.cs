﻿namespace AutoExcel
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnHelloWorld = new System.Windows.Forms.Button();
            this.btnInsertImage = new System.Windows.Forms.Button();
            this.btnFillColor = new System.Windows.Forms.Button();
            this.btnNewSheet = new System.Windows.Forms.Button();
            this.btnRenameSheet = new System.Windows.Forms.Button();
            this.btnChart = new System.Windows.Forms.Button();
            this.btnCopySheet = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbChartType = new System.Windows.Forms.ComboBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.cbShapeType = new System.Windows.Forms.ComboBox();
            this.btnChangeColorShape = new System.Windows.Forms.Button();
            this.btnAddTotalShapes = new System.Windows.Forms.Button();
            this.btnAddShape = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnHelloWorld
            // 
            this.btnHelloWorld.Location = new System.Drawing.Point(19, 33);
            this.btnHelloWorld.Name = "btnHelloWorld";
            this.btnHelloWorld.Size = new System.Drawing.Size(121, 23);
            this.btnHelloWorld.TabIndex = 0;
            this.btnHelloWorld.Text = "Hello World";
            this.btnHelloWorld.UseVisualStyleBackColor = true;
            this.btnHelloWorld.Click += new System.EventHandler(this.btnHelloWorld_Click);
            // 
            // btnInsertImage
            // 
            this.btnInsertImage.Location = new System.Drawing.Point(146, 33);
            this.btnInsertImage.Name = "btnInsertImage";
            this.btnInsertImage.Size = new System.Drawing.Size(121, 23);
            this.btnInsertImage.TabIndex = 0;
            this.btnInsertImage.Text = "Insert Image";
            this.btnInsertImage.UseVisualStyleBackColor = true;
            this.btnInsertImage.Click += new System.EventHandler(this.btnInsertImage_Click);
            // 
            // btnFillColor
            // 
            this.btnFillColor.Location = new System.Drawing.Point(273, 33);
            this.btnFillColor.Name = "btnFillColor";
            this.btnFillColor.Size = new System.Drawing.Size(121, 23);
            this.btnFillColor.TabIndex = 0;
            this.btnFillColor.Text = "Fill Color";
            this.btnFillColor.UseVisualStyleBackColor = true;
            this.btnFillColor.Click += new System.EventHandler(this.btnFillColor_Click);
            // 
            // btnNewSheet
            // 
            this.btnNewSheet.Location = new System.Drawing.Point(19, 35);
            this.btnNewSheet.Name = "btnNewSheet";
            this.btnNewSheet.Size = new System.Drawing.Size(121, 23);
            this.btnNewSheet.TabIndex = 0;
            this.btnNewSheet.Text = "Add new Sheet";
            this.btnNewSheet.UseVisualStyleBackColor = true;
            this.btnNewSheet.Click += new System.EventHandler(this.btnnewSheet_Click);
            // 
            // btnRenameSheet
            // 
            this.btnRenameSheet.Location = new System.Drawing.Point(146, 35);
            this.btnRenameSheet.Name = "btnRenameSheet";
            this.btnRenameSheet.Size = new System.Drawing.Size(121, 23);
            this.btnRenameSheet.TabIndex = 0;
            this.btnRenameSheet.Text = "Rename Sheet";
            this.btnRenameSheet.UseVisualStyleBackColor = true;
            this.btnRenameSheet.Click += new System.EventHandler(this.btnRenameSheet_Click);
            // 
            // btnChart
            // 
            this.btnChart.Location = new System.Drawing.Point(19, 49);
            this.btnChart.Name = "btnChart";
            this.btnChart.Size = new System.Drawing.Size(121, 23);
            this.btnChart.TabIndex = 0;
            this.btnChart.Text = "Create New Chart";
            this.btnChart.UseVisualStyleBackColor = true;
            this.btnChart.Click += new System.EventHandler(this.btnnewChart_Click);
            // 
            // btnCopySheet
            // 
            this.btnCopySheet.Location = new System.Drawing.Point(273, 35);
            this.btnCopySheet.Name = "btnCopySheet";
            this.btnCopySheet.Size = new System.Drawing.Size(121, 23);
            this.btnCopySheet.TabIndex = 0;
            this.btnCopySheet.Text = "Copy Sheet";
            this.btnCopySheet.UseVisualStyleBackColor = true;
            this.btnCopySheet.Click += new System.EventHandler(this.btnCopySheet_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnHelloWorld);
            this.groupBox1.Controls.Add(this.btnInsertImage);
            this.groupBox1.Controls.Add(this.btnFillColor);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(415, 81);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Basic Edit";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnNewSheet);
            this.groupBox3.Controls.Add(this.btnRenameSheet);
            this.groupBox3.Controls.Add(this.btnCopySheet);
            this.groupBox3.Location = new System.Drawing.Point(12, 110);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(415, 81);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Sheet Edit";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbChartType);
            this.groupBox2.Controls.Add(this.btnChart);
            this.groupBox2.Location = new System.Drawing.Point(12, 211);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(415, 100);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Chart";
            // 
            // cbChartType
            // 
            this.cbChartType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbChartType.FormattingEnabled = true;
            this.cbChartType.Items.AddRange(new object[] {
            "Bubble",
            "Cluster",
            "Pie",
            "ScatterLine",
            "3D Surface",
            "3D Cone Column",
            "3D Pyramid Column"});
            this.cbChartType.Location = new System.Drawing.Point(219, 49);
            this.cbChartType.Name = "cbChartType";
            this.cbChartType.Size = new System.Drawing.Size(175, 20);
            this.cbChartType.TabIndex = 1;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.cbShapeType);
            this.groupBox4.Controls.Add(this.btnChangeColorShape);
            this.groupBox4.Controls.Add(this.btnAddTotalShapes);
            this.groupBox4.Controls.Add(this.btnAddShape);
            this.groupBox4.Location = new System.Drawing.Point(14, 327);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(413, 101);
            this.groupBox4.TabIndex = 4;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Features";
            // 
            // cbShapeType
            // 
            this.cbShapeType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbShapeType.FormattingEnabled = true;
            this.cbShapeType.Items.AddRange(new object[] {
            "LineCallOut",
            "Rectangle",
            "Right Arrow",
            "Left Arrow",
            "Left Right Arror",
            "Up Arrow",
            "Right Triangle",
            "Octagon",
            "Oval"});
            this.cbShapeType.Location = new System.Drawing.Point(271, 34);
            this.cbShapeType.Name = "cbShapeType";
            this.cbShapeType.Size = new System.Drawing.Size(136, 20);
            this.cbShapeType.TabIndex = 1;
            // 
            // btnChangeColorShape
            // 
            this.btnChangeColorShape.Location = new System.Drawing.Point(17, 60);
            this.btnChangeColorShape.Name = "btnChangeColorShape";
            this.btnChangeColorShape.Size = new System.Drawing.Size(121, 22);
            this.btnChangeColorShape.TabIndex = 0;
            this.btnChangeColorShape.Text = "Change Color";
            this.btnChangeColorShape.UseVisualStyleBackColor = true;
            this.btnChangeColorShape.Click += new System.EventHandler(this.btnChangeColorShapes_Click);
            // 
            // btnAddTotalShapes
            // 
            this.btnAddTotalShapes.Location = new System.Drawing.Point(17, 32);
            this.btnAddTotalShapes.Name = "btnAddTotalShapes";
            this.btnAddTotalShapes.Size = new System.Drawing.Size(121, 22);
            this.btnAddTotalShapes.TabIndex = 0;
            this.btnAddTotalShapes.Text = "Add Total shapes";
            this.btnAddTotalShapes.UseVisualStyleBackColor = true;
            this.btnAddTotalShapes.Click += new System.EventHandler(this.btnAddTotalShapes_Click);
            // 
            // btnAddShape
            // 
            this.btnAddShape.Location = new System.Drawing.Point(164, 32);
            this.btnAddShape.Name = "btnAddShape";
            this.btnAddShape.Size = new System.Drawing.Size(101, 22);
            this.btnAddShape.TabIndex = 0;
            this.btnAddShape.Text = "Add shapes";
            this.btnAddShape.UseVisualStyleBackColor = true;
            this.btnAddShape.Click += new System.EventHandler(this.btnAddShape_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.button1);
            this.groupBox5.Location = new System.Drawing.Point(433, 12);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(424, 81);
            this.groupBox5.TabIndex = 5;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Images";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(27, 33);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Location = new System.Drawing.Point(433, 110);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(424, 81);
            this.groupBox6.TabIndex = 5;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Data Filter";
            // 
            // groupBox7
            // 
            this.groupBox7.Location = new System.Drawing.Point(433, 211);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(424, 100);
            this.groupBox7.TabIndex = 5;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Mathematics";
            // 
            // groupBox8
            // 
            this.groupBox8.Location = new System.Drawing.Point(433, 328);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(424, 100);
            this.groupBox8.TabIndex = 5;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Mic";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(878, 439);
            this.Controls.Add(this.groupBox8);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "AutoExcel Example";
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnHelloWorld;
        private System.Windows.Forms.Button btnInsertImage;
        private System.Windows.Forms.Button btnFillColor;
        private System.Windows.Forms.Button btnNewSheet;
        private System.Windows.Forms.Button btnRenameSheet;
        private System.Windows.Forms.Button btnChart;
        private System.Windows.Forms.Button btnCopySheet;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.ComboBox cbChartType;
        private System.Windows.Forms.Button btnAddShape;
        private System.Windows.Forms.ComboBox cbShapeType;
        private System.Windows.Forms.Button btnAddTotalShapes;
        private System.Windows.Forms.Button btnChangeColorShape;
        private System.Windows.Forms.Button button1;
    }
}

