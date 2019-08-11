using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Excel
{
    public partial class RailoadForm : Form
    {
        public bool[] railroaded;
        int onListElement = 0;

        public RailoadForm()
        {
            InitializeComponent();
            this.KeyPress += new KeyPressEventHandler(KeyPressed);
            this.Size = new Size(1500, 1000);
            CreateSaveButton();
        }

        private void CreateBox(string style, string fabric, bool railroad, string parts)
        {
            int horizontal = onListElement / 5;
            int vertical = onListElement % 5;
            int leftMostCorner = 100;
            GroupBox gb = new GroupBox();
            gb.Size = new Size(150, 125);
            gb.Location = new Point(leftMostCorner + horizontal * 150, 50 + vertical * 150);

            Label stl = new Label();
            stl.Location = new Point(20, 20);
            stl.Text = style;
            gb.Controls.Add(stl);

            Label fab = new Label();
            fab.Location = new Point(20, 45);
            fab.Text = fabric;
            gb.Controls.Add(fab);

            Label pts = new Label();
            pts.Location = new Point(20, 70);
            pts.Text = parts;
            gb.Controls.Add(pts);

            CheckBox cb = new CheckBox();
            cb.Checked = railroad;
            cb.Location = new Point(80, 90);
            gb.Controls.Add(cb);

            this.Controls.Add(gb);
        }

        public void PopulateItems(string style, string fabric, bool railroaded, List<string> parts = null)
        {
            string pts = "";
            if(parts != null) {
                foreach (string p in parts) {
                    pts += p + '/';
                }
                pts = pts.Remove(pts.Length - 1, 1);
            }
            CreateBox(style, fabric, railroaded, pts);
            onListElement++;
        }


        private void CreateSaveButton()
        {
            Button saveButton = new Button();

            saveButton.Location = new Point(750, 800);
            saveButton.Height = 30;
            saveButton.Width = 50;
            saveButton.Text = "Confirm";

            saveButton.Click += new EventHandler(SaveButton_Click);

            this.Controls.Add(saveButton);
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            AssignRailroad();
            this.Close();
        }

        private void AssignRailroad()
        {
            railroaded = new bool[onListElement];
            int i = 0;
            foreach (Control gb in this.Controls)
            {
                if (gb is GroupBox)
                {
                    foreach (Control c in gb.Controls)
                    {
                        if(c is CheckBox)
                        {
                            railroaded[i] = ((CheckBox)c).Checked;
                            i++;
                        }
                    }
                }
            }
        }

        private void RailoadForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveButton_Click(sender, e);
        }

        private void KeyPressed(Object o, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                SaveButton_Click(o, e);
            }
        }
    }
}
