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

namespace Anh.DB定義図＿WRS
{
    public partial class DB定義図＿WRSMap : Form
    {
        public int _iMaxRequest = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("MaxRequest")) ? 200 : int.Parse(ConfigurationManager.AppSettings.Get("MaxRequest"));
        public string _fromLang = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("SL")) ? "ja" : ConfigurationManager.AppSettings.Get("SL");
        public string _toLang = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TL")) ? "en" : ConfigurationManager.AppSettings.Get("TL");
        public int _iMaxLenPerRequest = string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("MaxLenPerRequest")) ? 1000 : int.Parse(ConfigurationManager.AppSettings.Get("MaxLenPerRequest"));
        public const string ecoEscapeBlank = "'";
        private Dictionary<string, string> _dicTableName;
        private string[] _arraySplitString = new string[] { "=", "＝", "||", "（+）", "(+)", "+", "-", "*", "/", " " };
        private bool _widthChange;
        public DB定義図＿WRSMap()
        {
            _dicTableName = new Dictionary<string, string>();
            InitializeComponent();
            _widthChange = false;
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (txtExcelName.Text.Length == 0)
            {
                linkLabel1.Focus();
                return;
            }
            if (cbSheetName.SelectedItem == null)
            {
                cbSheetName.Focus();
                return;
            }
            try
            {
                string sFileNameDBMapping = @".\DB定義図＿WRS.xlsx";
                //Excelファイル作成
                IWorkbook xlBookDB = SpreadsheetGear.Factory.GetWorkbook(sFileNameDBMapping);
                CreateDicTableName(xlBookDB);

                string sFileNameTarget = txtExcelName.Text;
                IWorkbook xlBookTarget = SpreadsheetGear.Factory.GetWorkbook(sFileNameTarget);
                string xlXheetNm = cbSheetName.SelectedItem.ToString();
                ConvertSheet(xlBookTarget.Worksheets[xlXheetNm], xlBookDB);

                string sFilePath = Path.GetDirectoryName(txtExcelName.Text) + @"\" + Path.GetFileNameWithoutExtension(sFileNameTarget) + "_" + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(sFileNameTarget);

                //保存
                xlBookTarget.SaveAs(sFilePath, FileFormat.OpenXMLWorkbook);
                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void CreateDicTableName(IWorkbook xlBookDB)
        {
            _dicTableName.Clear();
            string[] excludeSheet = new string[] { "テーブル一覧", "M044 仕入先別職種マスタ2","W081 見積テンプレート大工種ワーク", "W082 見積テンプレート小工種ワーク", "W084 見積テンプレート原価内訳ワーク"
                ,"W305 ポータル申請・承認顧客物件","W308B ポータル当月成績（資金繰り係数）"
            };
            for (int i = 0; i < xlBookDB.Worksheets.Count; i++)
            {
                IWorksheet xlSheet = xlBookDB.Worksheets[i];
                if (xlSheet.Name.Length >= 5 && !excludeSheet.Contains(xlSheet.Name) && !xlSheet.Name.EndsWith("マスタ取込ログファイル"))
                {
                    _dicTableName.Add(xlSheet.Name.Substring(0, 4).ToUpper() + ".", xlSheet.Name);
                }
            }
        }

        private IRange SearchSheet(IWorksheet xlSheet, string what)
        {
            IRange findedRange = xlSheet.Range.Find(what, null, FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, true);
            return findedRange;
        }
        #region "convert"
        private void ConvertSheet(IWorksheet xlSheet, IWorkbook xlBookDB)
        {
            Dictionary<IRange, string[][]> dicS = ExtractDataSheet(xlSheet);
            IWorksheet newSheet = xlSheet.CopyAfter(xlSheet) as IWorksheet;
            newSheet.Name = xlSheet.Name + "_DB";
            OutputDataToSheet(newSheet, dicS, xlBookDB);
        }

        private void OutputDataToSheet(IWorksheet newSheet, Dictionary<IRange, string[][]> dicS, IWorkbook xlBookDB)
        {
            foreach (IRange whatIR in dicS.Keys)
            {
                string addre = whatIR.Address;
                string resVal = "";
                for (int i = 0; i < dicS[whatIR].Length; i++)
                {
                    if (dicS[whatIR][i] != null)
                    {
                        string sheetNm = dicS[whatIR][i][0];
                        string sheetCd = dicS[whatIR][i][1];
                        string cellData = dicS[whatIR][i][2];
                        string splitString = dicS[whatIR][i][3];
                        if (sheetNm.Length > 0)
                        {
                            IRange findedIR = SearchSheet(xlBookDB.Worksheets[sheetNm], cellData);
                            if (findedIR != null)
                            {
                                IRange mapIR = xlBookDB.Worksheets[sheetNm].Cells[findedIR.Row, findedIR.Column + 8];
                                resVal = resVal + sheetCd + mapIR.Value.ToString() + splitString;
                            }
                            else
                            {
                                resVal = resVal + sheetCd + cellData + splitString;
                            }



                        }
                        else
                        {
                            resVal = resVal + cellData + splitString;
                        }

                    }

                }
                newSheet.Cells[addre].Value = resVal;
            }
        }

        private bool IsDataReq(string value)
        {
            string[] arrayData = value.Split(_arraySplitString, StringSplitOptions.None);
            for (int i = 0; i < arrayData.Length; i++)
            {
                string ku = arrayData[i].Trim();
                if (ku.Length >= 5)
                {
                    if (_dicTableName.ContainsKey(ku.Substring(0, 5).ToUpper()))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        private Dictionary<IRange, string[][]> ExtractDataSheet(IWorksheet xlSheet)
        {
            Dictionary<IRange, string[][]> dicRa = new Dictionary<IRange, string[][]>();
            IRange rMax = xlSheet.UsedRange;
            for (int i = 0; i < rMax.Columns.ColumnCount; i++)
            {
                for (int j = 0; j < rMax.Rows.RowCount; j++)
                {
                    IRange item = rMax.Cells[j, i];
                    if (item.Value != null && item.Value.ToString().Length >= 5 && IsDataReq(item.Value.ToString()))
                    {
                        string[] arrayData = item.Value.ToString().Split(_arraySplitString, StringSplitOptions.None);
                        string curPosOffset = "";
                        string[][] arrayRes = new string[arrayData.Length][];
                        for (int k = 0; k < arrayData.Length; k++)
                        {
                            curPosOffset = curPosOffset + arrayData[k];
                            string ku = arrayData[k].Trim();
                            if (ku.Length >= 5)
                            {
                                string sheetCd = ku.Substring(0, 5).ToUpper();
                                if (_dicTableName.ContainsKey(sheetCd))
                                {
                                    string sheetName = _dicTableName[sheetCd];
                                    arrayRes[k] = new string[] { sheetName, sheetCd, CleanInput(ku.Substring(5)), "" };

                                }
                                else
                                {
                                    arrayRes[k] = new string[] { "", "", ku, "" };
                                }
                            }
                            else
                            {
                                arrayRes[k] = new string[] { "", "", ku, "" };
                            }
                            foreach (string spl in _arraySplitString)
                            {
                                if (curPosOffset.Length + spl.Length <= item.Value.ToString().Length && item.Value.ToString().Substring(0, curPosOffset.Length + spl.Length).Equals(curPosOffset + spl))
                                {
                                    curPosOffset = curPosOffset + spl;
                                    arrayRes[k][3] = spl;
                                    break;
                                }
                            }
                        }
                        dicRa.Add(item, arrayRes);
                    }
                }
            }
            return dicRa;
        }

        private string CleanInput(string v)
        {
            return v.Replace("ｺｰﾄﾞ", "コード");
        }
        #endregion

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                txtExcelName.Text = openFileDialog1.FileName;
                string sFileNameTarget = txtExcelName.Text;
                IWorkbook xlBookTarget = SpreadsheetGear.Factory.GetWorkbook(sFileNameTarget);
                string[] arraySheetName = ConfigurationManager.AppSettings.Get("ArraySheetConvert").Split(',');
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
