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

namespace Anh.音声
{
	public partial class 音声 : Form
	{
		private bool _widthChange;
		public 音声()
		{
			InitializeComponent();
		}

		private async void btn聞く_Click(object sender, EventArgs e)
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
				byte[] arrayData = await ConvertSheet(xlBookTarget.Worksheets[xlXheetNm]);

				string sFilePath = Path.GetDirectoryName(txtExcelName.Text) + @"\" + Path.GetFileNameWithoutExtension(sFileNameTarget) + "_" + DateTime.Now.ToString("yyyyMMdd") + ".mp3";

				using (BinaryWriter w = new BinaryWriter(File.Open(sFilePath, FileMode.Create)))
				{
					w.Write(arrayData);
				}
				MessageBox.Show("Finished");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}

		}


		#region "音声"
		private async Task<byte[]> ConvertSheet(IWorksheet xlSheet)
		{
			Tuple<string, string> lang = GetHeaderSheet(xlSheet);
			List<Tuple<string, string>> listPair = ExtractDataSheet(xlSheet);
			var trans = new Anh.Translate.ActionF1();
			List<byte> l = new List<byte>();
			
			foreach (var item in listPair)
			{
				byte[] b1 = await trans.Translate_tts(item.Item1, lang.Item1);
				l.AddRange(b1);
				l.AddRange(Properties.Resources.ええと);
				byte[] b2 = await trans.Translate_tts(item.Item2, lang.Item2);
				l.AddRange(b2);
				l.AddRange(Properties.Resources.など);
			}
			return l.ToArray();
		}

		private Tuple<string, string> GetHeaderSheet(IWorksheet xlSheet)
		{
			IRange rMax = xlSheet.UsedRange;
			return new Tuple<string, string>(rMax.Cells[0, 0].Text.Trim(), rMax.Cells[0, 1].Text.Trim());
		}

		private List<Tuple<string, string>> ExtractDataSheet(IWorksheet xlSheet)
		{
			List<Tuple<string, string>> listPair = new List<Tuple<string, string>>();
			IRange rMax = xlSheet.UsedRange;
			for (int j = 1; j < Math.Min(rMax.Rows.RowCount,100); j++)
			{
				listPair.Add(new Tuple<string, string>(rMax.Cells[j, 0].Text.Trim(), rMax.Cells[j, 1].Text.Trim()));
				//break;
			}
			return listPair;
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
