﻿using System;
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

namespace Anh.音声
{
	public partial class 音声 : Form
	{
        public static readonly bool JoinWord = ConfigurationManager.AppSettings["JoinWord"] != null ? bool.Parse(ConfigurationManager.AppSettings["JoinWord"].ToString()) : false;
        public static readonly bool JoinSentence = ConfigurationManager.AppSettings["JoinSentence"] != null ? bool.Parse(ConfigurationManager.AppSettings["JoinSentence"].ToString()) : false;
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
				bool res = await ConvertSheet(xlBookTarget.Worksheets[xlXheetNm]);
				if (res)
				{
					MessageBox.Show("Finished");
					ExecuteCommand();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}

		}

		public void ExecuteCommand()
		{
			ProcessStartInfo ProcessInfo = new ProcessStartInfo(Application.StartupPath + @"\AudioOut\createAudio.bat");
			ProcessInfo.WorkingDirectory = Application.StartupPath + @"\AudioOut";
			Process.Start(ProcessInfo);
		}

		#region "音声"
		private async Task<bool> ConvertSheet(IWorksheet xlSheet)
		{
			Tuple<string, string> lang = GetHeaderSheet(xlSheet);
			List<Tuple<string, string>> listPair = ExtractDataSheet(xlSheet);
			var trans = new Anh.Translate.ActionF1();
			List<byte> l = new List<byte>();
			int i = 1;
			string folder = Application.StartupPath + @"\AudioOut\";
			if (Directory.Exists(folder))
			{
				foreach (var file in Directory.EnumerateFiles(folder))
				{
					FileInfo fi = new FileInfo(file);
					if (fi.Extension.Equals(".mp3"))
					{
						if (fi.Name == "the_next_word.mp3" || fi.Name == "wait_a_minute.mp3"
                             || fi.Name == "blanksound0.1s.mp3" || fi.Name == "blanksound0.2s.mp3" || fi.Name == "blanksound0.3s.mp3" || fi.Name == "blanksound0.5s.mp3")
							continue;
						fi.Delete();
					}
				}
			}
            StringBuilder sb = new StringBuilder();
            StringBuilder sbB = new StringBuilder();
            sbB.AppendLine("@ECHO OFF");
            string fileName = "";
            string sFilePath = "";
            int stt = 0;
            for (int im = 0; im < listPair.Count; im++)
            {
                var item = listPair[im];
                if (i + 1 % 200 == 0) 
                    Thread.Sleep((i + 1) * 10);
                if (!string.IsNullOrEmpty(lang.Item1))
                {
                    if (item.Item1 == "<<^Break>>" || item.Item1 == "<<Break>>" || item.Item1 == "<<$Break>>")
                    {
                        stt++;
                        switch (item.Item1)
                        {
                            case "<<^Break>>":
                                sb.Clear();
                                sbB.AppendLine(string.Format("ffmpeg -f concat -i audio{0}.txt -c copy {1}.mp3", stt, item.Item2));
                                break;
                            case "<<Break>>":
                                using (StreamWriter w = new StreamWriter(folder +string.Format(@"\audio{0}.txt",stt-1)))
                                {
                                    w.Write(sb.ToString());
                                }
                                sb.Clear();
                                sbB.AppendLine(string.Format("ffmpeg -f concat -i audio{0}.txt -c copy {1}.mp3", stt, item.Item2));
                                break;
                            case "<<$Break>>":
                                using (StreamWriter w = new StreamWriter(folder +string.Format(@"\audio{0}.txt",stt-1)))
                                {
                                    w.Write(sb.ToString());
                                }
                                sb.Clear();
                                break;
                            default:
                                break;
                        }
                        continue;
                    }
				    byte[] b1 = await trans.Translate_tts(item.Item1, lang.Item1);
				    l.AddRange(b1);
				    fileName = "A" + i + ".mp3";
				    sFilePath = folder + fileName;
				    sb.AppendLine(string.Format("file '{0}'", fileName));
                    if (JoinSentence)
                    {
                        if (im < listPair.Count - 1)
                        {
                            if (!string.IsNullOrEmpty(listPair[im+1].Item1))
                            {
                                string fc = listPair[im+1].Item1.Substring(0, 1);
                                char a = fc.ToCharArray()[0];
                                if (('A' <= a && a <= 'Z') || '0' <= a && a <= '9' || a =='Đ')
                                {
                                    sb.AppendLine("file 'blanksound0.2s.mp3'");
                                }
                                else
                                {
                                    //sb.AppendLine("file 'blanksound0.1s.mp3'");
                                }
                            }
                            else
                            {
                                sb.AppendLine("file 'blanksound0.2s.mp3'");
                            }
                        }
                    }
				    using (BinaryWriter w = new BinaryWriter(File.Open(sFilePath, FileMode.Create)))
				    {
					    w.Write(b1);
				    }
                }
                if (!string.IsNullOrEmpty(lang.Item2))
                {
				    fileName = "B" + i + ".mp3";
				    sFilePath = folder + fileName;
				    sb.AppendLine(string.Format("file '{0}'", fileName));
                    if (JoinWord) sb.AppendLine("file 'blanksound0.1s.mp3'");
				    byte[] b2 = await trans.Translate_tts(item.Item2, lang.Item2);
				    using (BinaryWriter w = new BinaryWriter(File.Open(sFilePath, FileMode.Create)))
				    {
					    w.Write(b2);
				    }
                }
				i++;
            }
			using (StreamWriter w = new StreamWriter(folder + @"\audio.txt"))
			{
				w.Write(sb.ToString());
			}
            using(StreamWriter w = new StreamWriter(folder + @"\createAudio.bat")){
                sbB.AppendLine("ECHO Congratulations! Your first batch file executed successfully.");
                sbB.AppendLine("PAUSE");
                w.Write(sbB.ToString());
            }
			return true;
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
			for (int j = 1; j < rMax.Rows.RowCount; j++)
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
