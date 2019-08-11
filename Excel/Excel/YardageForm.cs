using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace Excel
{
    public partial class YardageForm : Form
    {
        private bool railroaded = false;
        public double yardage = 0;
        public string type = null;
        public YardageForm(string style, bool railroaded, string type)
        {
            this.type = type;
            this.railroaded = railroaded;

            InitializeComponent();

            yardageText.KeyPress += new KeyPressEventHandler(KeyPressed);

            MatchingSelect();
            LabelPopulate(style);         
        }

        private void LabelPopulate(string style)
        {
            string railRoadAddOn = "";
            if (railroaded) { railRoadAddOn = "(RailRoaded)"; }
            this.styleLabel.Text = $"Please confirm the yardage for {style}{railRoadAddOn}";
        }

        private void MatchingSelect()
        {
            if (this.type.ToUpper() == "2M") { twomButton.Checked = true; }
            else if (this.type.ToUpper() == "1M") { onemButton.Checked = true;  }
            else { nmButton.Checked = true; }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            bool isValid = false;
            bool repeat = false;
            double yar;

            if (nmButton.Checked) { this.type = "NM"; }
            else if (onemButton.Checked) { this.type = "1M"; }
            else { this.type = "2M"; }

            while (!isValid)
            {
                // The user will be continually prompted to enter an appropriate yardage number until it is entered
                if (repeat)
                {
                    string tempYar;
                    tempYar = (Interaction.InputBox("Please Enter an appropriate number for Yardage", "Yardage", "", -1, -1));
                    yardageText.Text = tempYar;
                }
                if (double.TryParse(yardageText.Text, out yar))
                {
                    if (yar >= 0)
                    {
                        this.yardage = yar;
                        isValid = true;
                    }
                    else { repeat = true; }
                }
                else { repeat = true; }
            }
            this.Close();
        }

        public void Reinitialize(string style)
        {
            if (this.type == "NM") { nmButton.Checked = true; }
            else if (this.type == "1M") { onemButton.Checked = true; }
            else { twomButton.Checked = true; }

            LabelPopulate(style);
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
