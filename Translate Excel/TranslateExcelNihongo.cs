using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.IO;
using Anh.Translate;
using System.Threading;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;

namespace Anh.TranslateExcel
{
	public partial class TranslateExcelNihongo : Form
	{
		const string startCellCode = "((";
		const string startCellCode2 = "( (";
		const string endCellCode = "))";
		const string endCellCode2 = ") )";
		public int _iMaxRequest = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("MaxRequest")) ? 200 : int.Parse(ConfigurationManager.AppSettings.Get("MaxRequest"));
		public string _fromLang = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("SL")) ? "ja" : ConfigurationManager.AppSettings.Get("SL");
		public string _toLang = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TL")) ? "en" : ConfigurationManager.AppSettings.Get("TL");
		public int _iMaxLenPerRequest = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("MaxLenPerRequest")) ? 1000 : int.Parse(ConfigurationManager.AppSettings.Get("MaxLenPerRequest"));
		public const string ecoEscapeBlank = "'";
		private Dictionary<string, string> _dicTableName;
		private string[] _arraySplitString = new string[] { "=", "＝", "||", "（+）", "(+)", "+", "-", "*", "/", " " };
		private bool _widthChange;
		public TranslateExcelNihongo()
		{
			_dicTableName = new Dictionary<string, string>();
			InitializeComponent();
			_widthChange = false;
		}

		#region Translate
		private async Task<int> CreateTranslateSheet2(Excel.Worksheet xlSheet, int numPrevRequest)
		{
			List<Excel.Range> arRange = ExtractTranslateDataSheet2(xlSheet);
			int i = await TranslateSheet2(xlSheet, arRange, numPrevRequest);
			return i;
		}

		private List<Excel.Range> ExtractTranslateDataSheet2(Excel.Worksheet xlSheet)
		{
			int step = 0;
			if (!int.TryParse(ConfigurationManager.AppSettings.Get("RowRange"), out step)) step = 10;
			List<Excel.Range> arRange = new List<Excel.Range>();
			Excel.Range usedRange = xlSheet.UsedRange;
			int maxColumnCount = usedRange.Columns.Count;
			int maxRowCount = usedRange.Rows.Count;
			//5000 max tring google can translate per request (i will send request has maxlen enough small)
			//excel is not zero based!!
			for (int columnOffset = 1; columnOffset <= maxColumnCount; columnOffset++)
			{
				Excel.Range item = null;
				int totalLen = 0;
				int rowOffset = 1, iRowStartInRange = 1;
				while (rowOffset < maxRowCount)
				{
					iRowStartInRange = rowOffset;
					while (totalLen < _iMaxLenPerRequest)
					{
						//end of used row
						if (rowOffset >= maxRowCount)
						{
							break;
						}
						//current cell
						item = usedRange.Cells[rowOffset++, columnOffset];
						totalLen = GetLength(totalLen, item);
					}
					//total length received equal or greater than [config:max length per request]
					if (totalLen >= _iMaxLenPerRequest)
					{
						if (rowOffset - 1 < iRowStartInRange)
						{
							item = usedRange.Cells[iRowStartInRange, columnOffset];
							rowOffset = iRowStartInRange + 1;
						}
						else
						{
							item = usedRange.Range[usedRange.Cells[iRowStartInRange, columnOffset], usedRange.Cells[rowOffset - 1, columnOffset]];
						}
						arRange.Add(item);
						totalLen = 0;
					}
					else if (totalLen > 0)
					{
						item = usedRange.Range[usedRange.Cells[iRowStartInRange, columnOffset], usedRange.Cells[rowOffset, columnOffset]];
						arRange.Add(item);
						totalLen = 0;
					}
				}
			}

			return arRange;
		}

		private int GetLength(int offLen, Excel.Range item)
		{
			object v = item.Value;
			if (v == null)
			{
				return offLen;
			}
			return offLen + v.ToString().Length;
		}

		private async Task<int> TranslateSheet2(Excel.Worksheet currentSheet, List<Excel.Range> arRange, int numPrevRequest)
		{
			ActionF1 ActionF1 = new ActionF1();
			var ienum = arRange.AsEnumerable();
			TaskScheduler tsc = TaskScheduler.Current;
			int i = 0;
			Excel.Range whatIR = null;
			int limit = arRange.Count;

			limit = Math.Min(limit, _iMaxRequest - numPrevRequest);
			for (i = 0; i < limit; i++)
			{
				whatIR = arRange[i];
				//DataTable orgTa = whatIR.GetDataTable(SpreadsheetGear.Data.GetDataFlags.NoColumnHeaders); //error convert type double (set columntype = type of first cell
				DataTable orgTa = GetTableFromIrange(whatIR);
				string originalText = GetTextFromTable2(orgTa);
				if (originalText.Length > 5000)
				{
					continue;
				}
				JArray jarr = await ActionF1.GetSingle(originalText, _fromLang, _toLang);
				toolStripProgressBar1.Value = (int)((i + 1) * 100 / limit);
				Thread.Sleep(480);
				List<string> transateText = ActionF1.ReadJArrayRes2(jarr);
				////fake start
				//JArray jarr = await FakeTrans(originalText);
				//List<string> transateText = null;
				////fake end
				DataTable traTa = MergeResultIntoTable(transateText, orgTa);
				//translate sucess
				if (traTa != null)
				{
					for (int id = 0; id < traTa.Rows.Count; id++)
					{
						var im = traTa.Rows[id][1].ToString().Split(',')[0];
                        object vv = traTa.Rows[id][2];
						//string[] speStart = new string[] {"\n", "「" , "\"","“",};
						if (vv != null && vv.ToString().Trim().Length > 1)
						//&& (System.Text.RegularExpressions.Regex.IsMatch(vv.ToString().Trim().Substring(0, 1), @"[a-zA-Z0-9]") || speStart.Contains(vv.ToString().Substring(0, 1))))
						{
							if (currentSheet.Range[whatIR.Address].Cells[im, 1].Comment != null)
							{
								currentSheet.Range[whatIR.Address].Cells[im, 1].ClearComments();
							}
							currentSheet.Range[whatIR.Address].Cells[im, 1].AddComment(vv.ToString());
							Excel.Comment ic = currentSheet.Range[whatIR.Address].Cells[im, 1].Comment;
							ic.Shape.TextFrame.AutoSize = true;
							//using (Graphics g = this.CreateGraphics())
							//{
							//	string item = ic.ToString();
							//	SizeF sizeF = g.MeasureString(item, Font);
							//	ic.Shape.Width = sizeF.Width;
							//}
						}
					}
				}
				else
				{
					//remain originalText
				}
			}
			return limit + numPrevRequest;
		}

		private async Task<JArray> FakeTrans(string t)
		{
			Task<JArray> tt = Task.Run(() =>
			{
				return new JArray();
			});
			var m = await tt;
			return m;
		}

		private string ReplaceSpecialText(string text)
		{
			if (text.EndsWith("..."))
			{
				text = text.Trim('.');
			}
            if (text.Contains("No"))
            {
                text = text.Replace("No", "NO");
            }
            if (text.Contains("..."))
            {
                text = text.Replace("...", "_");
            }
			return text;
        }
		private string GetTextFromTable2(DataTable orgTa)
		{
			if (orgTa == null)
			{
				return null;
			}
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < orgTa.Rows.Count; i++)
			{
				string res = "";
				DataRow r = orgTa.Rows[i];
				if (r.IsNull(0) || r[0].ToString().Trim().Length == 0)
				{
					//if (i == 0)
					//{
					//	res = "。。。" + "\n";
					//}
					//else
					//{
						res = "\n";
					//}
				}
				else
				{
					string[] artm = r[0].ToString().Split(new string[] { "\n", "。" }, StringSplitOptions.RemoveEmptyEntries);

					res = artm.Length == 1 ? ReplaceSpecialText(artm[0]) : artm.Aggregate((m, n) => ReplaceSpecialText(m) + "。" + ReplaceSpecialText(n));
					res = res + "。。。" + "\n";
				}
				sb.Append(res);
			}
			string t = sb.ToString();
			return t;
		}
		private DataTable GetTableFromIrange(Excel.Range range)
		{
			DataTable dt = new DataTable();
			dt.Columns.Add("Column1", typeof(string));
			dt.Columns.Add("Column2", typeof(string));
            object[,] arr = range.Value as object[,];
			if (arr != null)
			{
                for (int k = 1; k <= arr.GetLength(0); k++)
				{

                    for (int l = 1; l <= arr.GetLength(1); l++)
					{
                        if (arr[k, l] != null && arr[k, l].ToString().Trim().Length > 0)
						{
                            dt.Rows.Add(arr[k,l].ToString().Trim(), $"{k},{l}");
						}
                    }
				}
    //                    foreach (var item in arr)
				//{
				//	if (item != null && item.ToString().Trim().Length > 0)
				//	{
				//		dt.Rows.Add(item.ToString().Trim(), "");
				//	}
				//	//else
				//	//{
				//	//	dt.Rows.Add("");
				//	//}
				//}
			}
			dt.AcceptChanges();
			return dt;
		}

		private DataTable MergeResultIntoTable(List<string> orgtex, DataTable orgTa)
		{
			DataTable dt = orgTa.Copy();

            dt.Columns.Add("Column3", typeof(string));
			
			if (orgtex == null || orgtex.Count == 0)
			{
				return null;
			}
			int iLen = Math.Min(orgtex.Count, dt.Rows.Count);
			for (int i = 0; i < iLen; i++)
			{
				dt.Rows[i][2] = orgtex[i].Replace(" 。 。","");
			}
			dt.AcceptChanges();
			return dt;
		}
		#endregion

		private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls";
			openFileDialog1.FilterIndex = 1;
			DialogResult res = openFileDialog1.ShowDialog();
			if (res == DialogResult.OK)
			{
				txtExcelName.Text = openFileDialog1.FileName;
				string sFileNameTarget = txtExcelName.Text;
				Excel.Application excelApp = new Excel.Application();
				Excel.Workbook xlBookTarget = excelApp.Workbooks.Add(sFileNameTarget);
				try
				{
					string[] arraySheetName = ConfigurationManager.AppSettings.Get("ArraySheetConvert").Split(',');
					int cSheet = xlBookTarget.Worksheets.Count;
					cbSheetName.Items.Clear();
					cbSheetName.ResetText();
					numF.Value = 1;
					numT.Value = 1;
					numF.Maximum = cSheet;
					numT.Maximum = cSheet;
					foreach (Excel.Worksheet xlXheet in xlBookTarget.Worksheets)
					{
						cbSheetName.Items.Add(xlXheet.Name);
					}
					_widthChange = true;
				}
				finally
				{
					xlBookTarget.Close();
					Marshal.ReleaseComObject(xlBookTarget);
					excelApp.Quit();
					Marshal.ReleaseComObject(excelApp);
				}
			}
		}

		private async void btnPre_Click(object sender, EventArgs e)
		{
			if (txtExcelName.Text.Length == 0)
			{
				linkLabel1.Focus();
				return;
			}
			if (rbSelectSheet.Checked && cbSheetName.SelectedItem == null)
			{
				cbSheetName.Focus();
				return;
			}
			string sFileNameTarget = txtExcelName.Text;
			Excel.Application excelApp = new Excel.Application();
			Excel.Workbook xlBookTarget = excelApp.Workbooks.Add(sFileNameTarget);
			try
			{
				//Excelファイル作成

				int cc = 0;
				if (rbSelectSheet.Checked)
				{
					string xlXheetNm = cbSheetName.SelectedItem.ToString();
					if (xlBookTarget.Worksheets[xlXheetNm] != null)
					{
						toolStripProgressBar1.Value = 1;
						toolStripStatusLabel1.Text = "1/1 Sheet.";
						toolStripStatusLabel2.Text = string.Format("{0} requests.", cc);
						cc = await CreateTranslateSheet2(xlBookTarget.Worksheets[xlXheetNm], 0);
						toolStripStatusLabel2.Text = string.Format("{0} requests.", cc);
					}

				}

				if (rbRangeSheet.Checked)
				{
					int startSheet = (int)numF.Value;
					int endSheet = (int)numT.Value < startSheet ? startSheet : (int)numT.Value;
					int iNumSheet = endSheet - startSheet + 1;
					List<string> sheetNames = new List<string>();
					for (int im = 1; im <= xlBookTarget.Worksheets.Count; im++)
					{
						if (im >= startSheet && im <= endSheet)
						{
							sheetNames.Add(xlBookTarget.Worksheets[im].Name);
						}

					}
					int iS = 0;
					foreach (Excel.Worksheet xlXheetNm in xlBookTarget.Worksheets)
					{
						if (sheetNames.IndexOf(xlXheetNm.Name) < 0)
						{
							continue;
						}
						iS++;
						if (cc < 0 || cc >= _iMaxRequest)
						{
							break;
						}
						toolStripProgressBar1.Value = 1;
						toolStripStatusLabel1.Text = string.Format("{0}/{1} Sheet.", iS, iNumSheet);
						toolStripStatusLabel2.Text = string.Format("{0} requests.", cc);
						cc = await CreateTranslateSheet2(xlXheetNm, cc);
						Thread.Sleep(1000 + 100 * iS);
						toolStripStatusLabel2.Text = string.Format("{0} requests.", cc);
					}
				}

				if (rbAllSheet.Checked)
				{
					int iNumSheet = xlBookTarget.Worksheets.Count;
					int iS = 0;
					foreach (Excel.Worksheet xlXheetNm in xlBookTarget.Worksheets)
					{
						iS++;
						if (cc < 0 || cc >= _iMaxRequest)
						{
							break;
						}
						toolStripProgressBar1.Value = 1;
						toolStripStatusLabel1.Text = string.Format("{0}/{1} Sheet.", iS, iNumSheet);
						toolStripStatusLabel2.Text = string.Format("{0} requests.", cc);
						cc = await CreateTranslateSheet2(xlXheetNm, cc);
						Thread.Sleep(1000 + 50 * iS);
						toolStripStatusLabel2.Text = string.Format("{0} requests.", cc);
					}
				}
				string sFilePath = Path.GetDirectoryName(txtExcelName.Text) + @"\" + Path.GetFileNameWithoutExtension(sFileNameTarget) + "_" + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(sFileNameTarget);
				bool bW = Helper.CanReadFile(sFilePath);
				if (!bW)
				{
					bW = MessageBox.Show(sFilePath + " is open. Should you close before continue ?", "!ݲ", MessageBoxButtons.YesNo) == DialogResult.Yes;
					if (bW)
					{
						int pp = 10;
						while (!(bW = Helper.CanReadFile(sFilePath)) && pp > 0)
						{
							pp--;
							Thread.Sleep(5000);
						}
					}
				}
				if (bW)
				{
					//保存
					//excelApp.Visible = true;
					xlBookTarget.SaveAs(sFilePath, Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
										Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value,
										System.Reflection.Missing.Value, System.Reflection.Missing.Value);
					xlBookTarget.Save();
				}
				MessageBox.Show("Finished");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
			finally
			{
			xlBookTarget.Close();
			Marshal.ReleaseComObject(xlBookTarget);
			excelApp.Quit();
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

		private void rbAllSheet_CheckedChanged(object sender, EventArgs e)
		{
			cbSheetName.Enabled = false;
			numF.Enabled = false;
			numT.Enabled = false;
		}

		private void rbSelectSheet_CheckedChanged(object sender, EventArgs e)
		{
			cbSheetName.Enabled = true;
			numF.Enabled = false;
			numT.Enabled = false;
		}

		private void rbRangeSheet_CheckedChanged(object sender, EventArgs e)
		{
			cbSheetName.Enabled = false;
			numF.Enabled = true;
			numT.Enabled = true;
		}

		private void SampleCode()
		{
			//			.NET 4 + allows C# to read and manipulate Microsoft Excel files, for computers that have Excel installed (if you do not have Excel installed, see NPOI).

			//First, add the reference to Microsoft Excel XX.X Object Library, located in the COM tab of the Reference Manager.I have given this the using alias of Excel.

			//using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab
			//			Next, you'll need to create references for each COM object that is accessed. Each reference must be kept to effectively exit the application on completion.

			////Create COM Objects. Create a COM object for everything that is referenced
			//			Excel.Application xlApp = new Excel.Application();
			//			Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"sandbox_test.xlsx");
			//			Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
			//			Excel.Range xlRange = xlWorksheet.UsedRange;
			//			Then you can read from the sheet, keeping in mind that indexing in Excel is not 0 based.This just reads the cells and prints them back just as they were in the file.

			////iterate over the rows and columns and print to the console as it appears in the file
			////excel is not zero based!!
			//			for (int i = 1; i <= rowCount; i++)
			//			{
			//				for (int j = 1; j <= colCount; j++)
			//				{
			//					//new line
			//					if (j == 1)
			//						Console.Write("\r\n");

			//					//write the value to the console
			//					if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
			//						Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

			//					//add useful things here!   
			//				}
			//			}
			//			Lastly, the references to the unmanaged memory must be released.If this is not properly done, then there will be lingering processes that will hold the file access writes to your Excel workbook.

			////cleanup
			//			GC.Collect();
			//			GC.WaitForPendingFinalizers();

			//			//rule of thumb for releasing com objects:
			//			//  never use two dots, all COM objects must be referenced and released individually
			//			//  ex: [somthing].[something].[something] is bad

			//			//release com objects to fully kill excel process from running in the background
			//			Marshal.ReleaseComObject(xlRange);
			//			Marshal.ReleaseComObject(xlWorksheet);

			//			//close and release
			//			xlWorkbook.Close();
			//			Marshal.ReleaseComObject(xlWorkbook);

			//			//quit and release
			//			xlApp.Quit();
			//			Marshal.ReleaseComObject(xlApp);
		}
	}

	internal static class Helper
	{
		const int ERROR_SHARING_VIOLATION = 32;
		const int ERROR_LOCK_VIOLATION = 33;

		public static bool IsFileLocked(Exception exception)
		{
			int errorCode = System.Runtime.InteropServices.Marshal.GetHRForException(exception) & ((1 << 16) - 1);
			return errorCode == ERROR_SHARING_VIOLATION || errorCode == ERROR_LOCK_VIOLATION;
		}

		public static bool CanReadFile(string filePath)
		{
			//Try-Catch so we dont crash the program and can check the exception
			try
			{
				//The "using" is important because FileStream implements IDisposable and
				//"using" will avoid a heap exhaustion situation when too many handles  
				//are left undisposed.
				using (FileStream fileStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
				{
					if (fileStream != null) fileStream.Close();  //This line is me being overly cautious, fileStream will never be null unless an exception occurs... and I know the "using" does it but its helpful to be explicit - especially when we encounter errors - at least for me anyway!
				}
			}
			catch (IOException ex)
			{
				//THE FUNKY MAGIC - TO SEE IF THIS FILE REALLY IS LOCKED!!!
				if (IsFileLocked(ex))
				{
					// do something, eg File.Copy or present the user with a MsgBox - I do not recommend Killing the process that is locking the file
					return false;
				}
			}
			finally
			{ }
			return true;
		}
	}
}
