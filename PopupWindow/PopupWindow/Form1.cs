using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tulpep.NotificationWindow;

namespace PopupWindow
{
    public partial class Form1 : Form
    {
        private bool _widthChange;
        PopupNotifier popup;
        public Form1()
        {
            InitializeComponent();
        }

        #region "common"
        private void lnkFileName_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 1;
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                txtExcelName.Text = openFileDialog1.FileName;
                string sFileNameTarget = txtExcelName.Text;
                IWorkbook xlBookTarget = SpreadsheetGear.Factory.GetWorkbook(sFileNameTarget);
                int cSheet = xlBookTarget.Worksheets.Count;
                cbSheetName.Items.Clear();
                cbSheetName.ResetText();
                foreach (IWorksheet xlXheet in xlBookTarget.Worksheets)
                {
                    cbSheetName.Items.Add(xlXheet.Name);
                }
                _widthChange = true;
            }
        }

        private void cbSheetName_DropDown(object sender, EventArgs e)
        {
            if (!_widthChange)
            {
                return;
            }
            else
            {
                _widthChange = false;
            }
            ComboBox senderComboBox = (ComboBox)sender;
            int width = senderComboBox.DropDownWidth;
            Graphics g = senderComboBox.CreateGraphics();
            Font font = senderComboBox.Font;
            int vertScrollBarWidth =
                (senderComboBox.Items.Count > senderComboBox.MaxDropDownItems)
                ? SystemInformation.VerticalScrollBarWidth : 0;

            int newWidth;
            foreach (string s in ((ComboBox)sender).Items)
            {
                newWidth = (int)g.MeasureString(s, font).Width
                    + vertScrollBarWidth;
                if (width < newWidth)
                {
                    width = newWidth;
                }
            }
            senderComboBox.DropDownWidth = width;
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            popup = new MPopupNotifier();
            popup.TitleText = "BE HAPPY";
            popup.Delay = 500000;
            popup.OptionsMenu = contextMenuStrip1;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (li.Count > 0)
            {
                if (index < 0 || index >= li.Count) index = 0;
                popup.ContentText = li[index];
                index++;
            }
            else
            {
                popup.ContentText = "Thank you" + DateTime.Now.ToLongTimeString();
            }
            popup.Popup();
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            popup.Hide();
        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            timer1.Stop();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            timer1.Interval = (int)numericUpDown1.Value;
            timer1.Stop();
            timer1.Start();
        }


        List<string> li = new List<string>();
        int index = 0;
        private void cbSheetName_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sFileNameTarget = txtExcelName.Text;
            IWorkbook xlBookTarget = SpreadsheetGear.Factory.GetWorkbook(sFileNameTarget);
            string xlXheetNm = cbSheetName.SelectedItem.ToString();
            IWorksheet xlSheet = xlBookTarget.Worksheets[xlXheetNm];
            IRange rMax = xlSheet.UsedRange;
            int max = rMax.Rows.RowCount;
            int maxCo = rMax.Columns.ColumnCount;
            string r = "";
            li.Clear();
            for (int j = 0; j < max; j++)
            {
                r = "";
                for (int i = 0; i < maxCo; i++)
                {
                    r += rMax.Cells[j, i].Text + '\t';
                }
                li.Add(r);
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            index--;
            index--;
            timer1.Start();
        }

        private void btnLast_Click(object sender, EventArgs e)
        {
            index = li.Count - 1;
            timer1.Start();
        }

        private void btnFisrt_Click(object sender, EventArgs e)
        {
            index = 0;
            timer1.Start();
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnStop_Click(sender, null);
        }

        private void pauseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnPause_Click(sender, null);
        }

        private void fisrtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnFisrt_Click(sender, null);
        }

        private void backToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnBack_Click(sender, null);
        }

        private void lastToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnLast_Click(sender, null);
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnStart_Click(sender, null);
        }
    }
}
