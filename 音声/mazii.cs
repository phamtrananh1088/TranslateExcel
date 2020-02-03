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
using Anh.Translate;

namespace Anh.音声
{
	public partial class mazii : Form
	{
		private bool _widthChange;
        private ActionF1 translate_a;
        public mazii()
		{
			InitializeComponent();
		}

        private async void btnsearchExample_Click(object sender, EventArgs e)
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
                IWorksheet xlSheet = xlBookTarget.Worksheets[xlXheetNm];
                IRange rMax = xlSheet.UsedRange;
                int max = rMax.Rows.RowCount;
                bool res = false;
                for (int j = 0; j < rMax.Rows.RowCount; j++)
                {
                    string v1 = rMax.Cells[j, 0].Text.Trim();
                    string v2 = rMax.Cells[j, 1].Text.Trim();
                    if (v2.Length == 0)
                    {
                        toolStripProgressBar1.Value = j*100 / max;
                        JObject dataEx = await searchExample(v1);
                        if (dataEx["results"]!=null)
                        {
                            for (int i = 0; i < 5; i++)
                            {
                                if ( dataEx["results"][i] !=null && dataEx["results"][i]["content"]!=null && dataEx["results"][i]["transcription"]!=null && dataEx["results"][i]["mean"]!=null)
                                {
                                    rMax.Cells[j, 1].Value = dataEx["results"][i]["content"].ToString();
                                    rMax.Cells[j, 2].Value = dataEx["results"][i]["transcription"].ToString();
                                    rMax.Cells[j, 3].Value = dataEx["results"][i]["mean"].ToString();
                                    break;
                                }
                            }
                        }
                        JObject dataExV = await searchExample(v1, "javi");
                        if (dataExV["results"]!=null)
                        {
                            for (int i = 0; i < 5; i++)
                            {
                                if (dataExV["results"][i] != null && dataExV["results"][i]["content"] != null && dataExV["results"][i]["transcription"] != null && dataExV["results"][i]["mean"] != null)
                                {
                                    rMax.Cells[j, 4].Value = dataExV["results"][i]["content"].ToString();
                                    rMax.Cells[j, 5].Value = dataExV["results"][i]["transcription"].ToString();
                                    rMax.Cells[j, 6].Value = dataExV["results"][i]["mean"].ToString();
                                    break;
                                }
                            }
                        }
                        res = true;
                    }
                }

                if (res)
                {
                    xlBookTarget.SaveAs(sFileNameTarget, FileFormat.Excel8);
                    MessageBox.Show("Updated");
                }
                else
                {
                    MessageBox.Show("No thing Updated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

		private async void btnkanji_Click(object sender, EventArgs e)
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
				IWorksheet xlSheet = xlBookTarget.Worksheets[xlXheetNm];
				IRange rMax = xlSheet.UsedRange;
                int max = rMax.Rows.RowCount;
				bool res = false;
				for (int j = 0; j < rMax.Rows.RowCount; j++)
				{
					string v1 = rMax.Cells[j, 0].Text.Trim();
					string v2 = rMax.Cells[j, 1].Text.Trim();
					if (v2.Length==0)
					{
                        toolStripProgressBar1.Value = j / max * 100;
						rMax.Cells[j, 1].Value = await kanji(v1);
						res = true;
					}
				}
			
				if (res)
				{
                    xlBookTarget.SaveAs(sFileNameTarget, FileFormat.OpenXMLWorkbook);
					MessageBox.Show("Updated");
				}
				else
				{
					MessageBox.Show("No thing Updated");
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
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
                IWorksheet xlSheet = xlBookTarget.Worksheets[xlXheetNm];
                IRange rMax = xlSheet.UsedRange;
                int max = rMax.Rows.RowCount;
                bool res = false;
                for (int j = 0; j < rMax.Rows.RowCount; j++)
                {
                    string v1 = rMax.Cells[j, 0].Text.Trim();
                    string v2 = rMax.Cells[j, 1].Text.Trim();
                    if (v2.Length == 0)
                    {
                        toolStripProgressBar1.Value = j / max * 100;
                        rMax.Cells[j, 1].Value = await tokenizer(v1);
                        res = true;
                    }
                }
				if (res)
				{
                    xlBookTarget.SaveAs(sFileNameTarget, FileFormat.OpenXMLWorkbook);
					MessageBox.Show("Finished");
                }
                else
                {
                    MessageBox.Show("No thing Updated");
                }
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}

		}

        #region "tokenizer"
        private async Task<string> tokenizer(string q)
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
		#endregion
		
		#region "kanji"
		private async Task<string> kanji(string q)
        {
            string res = null;
            HttpResponseMessage response = null;
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://data.mazii.net/");
                    string path = "/kanji/" + Functions.KanjiKodo(q) + ".svg";
                    JObject mixParam = new JObject();
                    response = await client.GetAsync(path);
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
		#endregion

        #region "search"
        private async Task<JObject> searchExample(string q, string dict = "jaen")
        {
            JObject res = null;
            HttpResponseMessage response = null;
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://mazii.net/");
                    string path = "/api/search";
                    JObject mixParam = new JObject();
                    mixParam.Add("dict", dict);
                    mixParam.Add("type", "example");
                    mixParam.Add("query", q);
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
                res = await response.Content.ReadAsAsync <JObject>();
            }
            return res;
        }
        private async Task<JObject> searchKanji(string q)
        {
            JObject res = null;
            HttpResponseMessage response = null;
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://mazii.net/");
                    string path = "/api/search";
                    JObject mixParam = new JObject();
                    mixParam.Add("dict", "jaen");
                    mixParam.Add("type", "kanji");
                    mixParam.Add("query", q);
                    mixParam.Add("page", 1);
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
                res = await response.Content.ReadAsAsync<JObject>();
            }
            return res;
        }
        #endregion

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

        private void mazii_Load(object sender, EventArgs e)
        {
            translate_a = new ActionF1();
        }
	}
	
	public static partial class Functions
	{
        public static string KanjiKodo(string q)
        {
            byte[] ba = UnicodeEncoding.Unicode.GetBytes(q);
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            for (int i = ba.Length - 1; i >= 0; i--)
            {
                hex.AppendFormat("{0:x2}", ba[i]);
            }
            return hex.ToString();
        }
	}
}
