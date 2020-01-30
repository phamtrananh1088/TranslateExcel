using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetGear;
using System.Configuration;
using System.IO;
using System.Threading;
using Newtonsoft.Json.Linq;
using System.Resources;
using System.Diagnostics;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;

namespace Anh.音声
{
	public partial class mazii : Form
	{
		private bool _widthChange;
        public mazii()
		{
			InitializeComponent();
		}

        private async void btntokenizer_Click(object sender, EventArgs e)
		{
			if (txtExcelName.Text.Length == 0)
			{
				linkFileName.Focus();
				return;
			}
			if (cbSheetName.SelectedItem == null)
			{
				cbSheetName.Focus();
				return;
			}
			try
			{

				string sFileNameTarget = txtExcelName.Text;
				IWorkbook xlBookTarget = SpreadsheetGear.Factory.GetWorkbook(sFileNameTarget);
				string xlXheetNm = cbSheetName.SelectedItem.ToString();
				bool res = await ConvertSheet(xlBookTarget.Worksheets[xlXheetNm]);
				if (res)
				{
                    xlBookTarget.SaveAs(sFileNameTarget, FileFormat.OpenXMLWorkbook);
					MessageBox.Show("Finished");
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}

		}

        #region "tokenizer"
        private async Task<bool> ConvertSheet(IWorksheet xlSheet)
		{
            IRange rMax = xlSheet.UsedRange;
            for (int j = 0; j < rMax.Rows.RowCount; j++)
            {
                string v1 = rMax.Cells[j, 0].Text.Trim();
                string v2 = rMax.Cells[j, 1].Text.Trim();
                if (v2.Length==0)
                {
                    rMax.Cells[j, 1].Value = await tokenizer(v1);
                }
            }
			return true;
		}

        /// <summary>
        /// Voice text
        /// </summary>
        /// <param name="q"></param>
        /// <param name="tl"></param>
        /// <returns></returns>
        public async Task<string> tokenizer(string q)
        {
            string res = null;
            HttpResponseMessage response = null;
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://mazii.net/");
                    string path = "/api/tokenizer";
                    JObject mixParam = new JObject();
                    mixParam.Add("text", q);
                    string send = Newtonsoft.Json.JsonConvert.SerializeObject(mixParam);
                    response = await client.PostAsJsonAsync(path, mixParam);
                }
            }
            catch (Exception ex)
            {
                Debug.Fail("エラー: {0}", ex.ToString());
            }
            finally
            {
            }
            if (response.IsSuccessStatusCode)
            {
                res = await response.Content.ReadAsStringAsync();
            }
            return res;
        }

        private void Authen(ref HttpRequestHeaders reqH)
        {
            reqH.Add("authority", "mazii.net");
            //reqH.Add("accept", "*/*");
            //reqH.Accept.Clear();
            //reqH.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //reqH.Add("scheme", "https");
            //reqH.Add("method", "GET");
            //reqH.Add("accept-encoding", "gzip, deflate, br");
            //reqH.Add("accept-language", "en-US,en;q=0.9,vi;q=0.8,fr-FR;q=0.7,fr;q=0.6");
            //reqH.Add("referer", "https://translate.google.com.vn/?hl=vi");
            string chrome = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36";
            //string edge = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.18362";
            reqH.Add("user-agent", chrome);
            //reqH.GetCookies().Add(new CookieHeaderValue("NID", "188=XFYB2Lnp4nax9yrsHueGO_c5X5mbBYLo5L95ZRPwDB-dvpyttWrVYyabPHPoHv_ItYhmdUOwlR1vE3lFvG7BgzEMxNgjjFX6Uv9gHmz6bvK-IdzbeNeq3oz8b60ZMSEahO97uW3ws2kIzvHQRgR9l_Dl9afmTtzRr3q2IIRWJlU"));
            //reqH.Add("Cache-Control", new string[] { "no-cache, no-store, must-revalidate" });
            //reqH.Add("Pragma","no-cache");
            //reqH.Add("x-client-data", "CIq2yQEIo7bJAQjBtskBCNG3yQEIqZ3KAQioo8oBCLGnygEI4qjKAQigqcoBCPGpygEIl63KAQjNrcoB");
            //reqH.Add("x-client-data", "aaaCIq2yQEIo7bJAQjBtskBCNG3yQEIqZ3KAQioo8oBCLGnygEI4qjKAQigqcoBCPGpygEIl63KAQjNrcoB");
        }

		#endregion

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
	}
}
