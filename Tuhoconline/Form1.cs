using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Anh.Tuhoconline
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public async Task<string> GetHtmlContent(int aiPage, string page)
        {
            string res = null;
            HttpResponseMessage response = null;
            try
            {
                //using (var client = new HttpClient(handler))
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://tuhoconline.net/");
                    HttpRequestHeaders reqH = client.DefaultRequestHeaders;
                    //Authen(ref reqH);
                    //reqH.Add(HttpRequestHeader.Cookie.ToString(), "NID=188=" + RenCookie());
                    //HTTP GET

                    string path = string.Format("{0}/{1}", page, aiPage);
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

        private void Authen(ref HttpRequestHeaders reqH)
        {
            //reqH.Add("authority", "translate.google.com.vn");
            //reqH.Add("content-type", "text/html; charset=UTF-8");
            //reqH.Add("accept", "*/*");
            //reqH.Accept.Clear();
            //reqH.Accept.Add(new MediaTypeWithQualityHeaderValue("text/html"));
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

        private async void btn共通単語_Click(object sender, EventArgs e)
        {
            int iPage = 1;
            string res = await GetHtmlContent(iPage, "3000-tu-vung-tieng-nhat-thong-dung.html");
            DataTable data = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.Find3000共通単語);
            while (iPage < 17)
            {
                iPage++;
                res = await GetHtmlContent(iPage, "3000-tu-vung-tieng-nhat-thong-dung.html");
                DataTable data1 = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.Find3000共通単語);
                data.Merge(data1);
            }
            await ExtractData_3000共通単語(data);
            OutPutExcel_3000共通単語(data);
        }

        private async Task ExtractData_3000共通単語(DataTable data)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<html><head><title>Google</title></head><body><table>");
            for (int i = 0; i < data.Rows.Count; i++)
            {
                sb.Append(data.Rows[i]["tooltipText"]);
            }
            sb.Append("</table></body</html>");
            string res = sb.ToString();
            await HtmlSourceCode.GetCode(res, data, HtmlSourceCode.ExtractData_3000共通単語);
        }

        private void OutPutExcel_3000共通単語(DataTable adtData)
        {
            string templatePath = Common.GetTemplate("3000共通単語_テンプレート.xls");
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(templatePath);
            try
            {
                SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets["3000共通単語"];
                SpreadsheetGear.IRange range = null;

                for (int i = 0, addressY = 2; i < adtData.Rows.Count; i++, addressY++)
                {
                    range = worksheet.Cells["A" + addressY];
                    range.Value = adtData.Rows[i]["id"];

                    range = worksheet.Cells["B" + addressY];
                    range.Value = adtData.Rows[i]["jp"];
                    if (adtData.Rows[i]["read"].ToString().Length > 0 || adtData.Rows[i]["tooltipText"].ToString().Length > 0)
                    {
                        range.AddComment(adtData.Rows[i]["read"].ToString().Length > 0 ? (adtData.Rows[i]["read"].ToString() + Environment.NewLine + adtData.Rows[i]["tooltipText"].ToString()) : adtData.Rows[i]["tooltipText"].ToString());
                        SpreadsheetGear.IComment icomment = range.Comment;
                        using (Graphics g = this.CreateGraphics())
                        {
                            string item = icomment.ToString();
                            SizeF sizeF = g.MeasureString(item, Font);
                            icomment.Shape.Width = sizeF.Width;
                            icomment.Shape.Height = sizeF.Height;
                        }
                    }

                    range = worksheet.Cells["C" + addressY];
                    range.Value = adtData.Rows[i]["read"];
                    range.WrapText = false;

                    range = worksheet.Cells["D" + addressY];
                    range.Value = adtData.Rows[i]["vi"];
                    range.WrapText = false;

                    range = worksheet.Cells["E" + addressY];
                    range.Value = adtData.Rows[i]["innerText"];
                    range.WrapText = false;

                    range = worksheet.Cells["F" + addressY];
                    range.Value = adtData.Rows[i]["outerHtml"];
                    range.WrapText = false;

                    range = worksheet.Cells["G" + addressY];
                    range.Value = adtData.Rows[i]["tooltipText"];
                    range.WrapText = false;
                }
                string outPath = "";
                Common.SaveExcelTemplate(workbook, "3000共通単語", "xls", out outPath);
                if (File.Exists(outPath))
                {
                    System.Diagnostics.Process.Start(outPath);
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                workbook.Close();
            }



        }

        private void OutPutExcel_1000共通単語(DataTable adtData)
        {
            string templatePath = Common.GetTemplate("1000共通単語_テンプレート.xls");
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(templatePath);
            try
            {
                SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets["1000共通単語"];
                SpreadsheetGear.IRange range = null;

                for (int i = 0, addressY = 2; i < adtData.Rows.Count; i++, addressY++)
                {
                    range = worksheet.Cells["A" + addressY];
                    range.Value = adtData.Rows[i]["id"];

                    range = worksheet.Cells["B" + addressY];
                    range.Value = adtData.Rows[i]["jp"];
                    if (adtData.Rows[i]["read"].ToString().Length > 0 || adtData.Rows[i]["tooltipText"].ToString().Length > 0)
                    {
                        range.AddComment(adtData.Rows[i]["read"].ToString().Length > 0 ? (adtData.Rows[i]["read"].ToString() + Environment.NewLine + adtData.Rows[i]["tooltipText"].ToString()) : adtData.Rows[i]["tooltipText"].ToString());
                        SpreadsheetGear.IComment icomment = range.Comment;
                        using (Graphics g = this.CreateGraphics())
                        {
                            string item = icomment.ToString();
                            SizeF sizeF = g.MeasureString(item, Font);
                            icomment.Shape.Width = sizeF.Width;
                            icomment.Shape.Height = sizeF.Height;
                        }
                    }

                    range = worksheet.Cells["C" + addressY];
                    range.Value = adtData.Rows[i]["read"];
                    range.WrapText = false;

                    range = worksheet.Cells["D" + addressY];
                    range.Value = adtData.Rows[i]["vi"];
                    range.WrapText = false;

                    range = worksheet.Cells["E" + addressY];
                    range.Value = adtData.Rows[i]["innerText"];
                    range.WrapText = false;

                    range = worksheet.Cells["F" + addressY];
                    range.Value = adtData.Rows[i]["outerHtml"];
                    range.WrapText = false;

                    range = worksheet.Cells["G" + addressY];
                    range.Value = adtData.Rows[i]["tooltipText"];
                    range.WrapText = false;
                }
                string outPath = "";
                Common.SaveExcelTemplate(workbook, "1000共通単語", "xls", out outPath);
                if (File.Exists(outPath))
                {
                    System.Diagnostics.Process.Start(outPath);
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                workbook.Close();
            }
        }

        private void OutPutExcel_2000共通単語(DataTable adtData)
        {
            string templatePath = Common.GetTemplate("2000共通単語_テンプレート.xls");
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(templatePath);
            try
            {
                SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets["2000共通単語"];
                SpreadsheetGear.IRange range = null;

                for (int i = 0, addressY = 2; i < adtData.Rows.Count; i++, addressY++)
                {
                    range = worksheet.Cells["A" + addressY];
                    range.Value = adtData.Rows[i]["id"];

                    range = worksheet.Cells["B" + addressY];
                    range.Value = adtData.Rows[i]["jp"];
                    if (adtData.Rows[i]["read"].ToString().Length > 0 || adtData.Rows[i]["tooltipText"].ToString().Length > 0)
                    {
                        range.AddComment(adtData.Rows[i]["read"].ToString().Length > 0 ? (adtData.Rows[i]["read"].ToString() + Environment.NewLine + adtData.Rows[i]["tooltipText"].ToString()) : adtData.Rows[i]["tooltipText"].ToString());
                        SpreadsheetGear.IComment icomment = range.Comment;
                        using (Graphics g = this.CreateGraphics())
                        {
                            string item = icomment.ToString();
                            SizeF sizeF = g.MeasureString(item, Font);
                            icomment.Shape.Width = sizeF.Width;
                            icomment.Shape.Height = sizeF.Height;
                        }
                    }

                    range = worksheet.Cells["C" + addressY];
                    range.Value = adtData.Rows[i]["read"];
                    range.WrapText = false;

                    range = worksheet.Cells["D" + addressY];
                    range.Value = adtData.Rows[i]["vi"];
                    range.WrapText = false;

                    range = worksheet.Cells["E" + addressY];
                    range.Value = adtData.Rows[i]["innerText"];
                    range.WrapText = false;

                    range = worksheet.Cells["F" + addressY];
                    range.Value = adtData.Rows[i]["outerHtml"];
                    range.WrapText = false;

                    range = worksheet.Cells["G" + addressY];
                    range.Value = adtData.Rows[i]["tooltipText"];
                    range.WrapText = false;
                }
                string outPath = "";
                Common.SaveExcelTemplate(workbook, "2000共通単語", "xls", out outPath);
                if (File.Exists(outPath))
                {
                    System.Diagnostics.Process.Start(outPath);
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                workbook.Close();
            }
        }

        private void OutPutExcel_N5Nguphap(DataTable adtData)
        {
            string templatePath = Common.GetTemplate("N5Nguphap_テンプレート.xls");
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(templatePath);
            try
            {
                SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets["N5Nguphap"];
                SpreadsheetGear.IRange range = null;
                //tb.Columns.Add("id");
                //tb.Columns.Add("maucau");
                //tb.Columns.Add("cachchia");
                //tb.Columns.Add("ynghia");
                //tb.Columns.Add("vidu");
                string dataA, dataB, dataC, dataD, dataE;
                string [] numofLineA, numofLineB, numofLineC, numofLineD, numofLineE;
                for (int i = 0, addressY = 2, plus = 0; i < adtData.Rows.Count;i++, addressY++,plus=0)
                {
                    dataA = adtData.Rows[i]["id"].ToString();
                    dataB = adtData.Rows[i]["maucau"].ToString();
                    dataC = adtData.Rows[i]["cachchia"].ToString();
                    dataD = adtData.Rows[i]["ynghia"].ToString();
                    dataE = adtData.Rows[i]["vidu"].ToString();
                    numofLineA = dataA.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineB = dataB.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineC = dataC.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineD = dataD.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineE = dataE.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    for (int plusP = 0; plusP < numofLineA.Length; plusP++)
                    {
                        range = worksheet.Cells["A" + (addressY+plusP)];
                        range.Value = numofLineA[plusP];   
                    }
                    for (int plusP = 0; plusP < numofLineB.Length; plusP++)
                    {
                        range = worksheet.Cells["B" + (addressY + plusP)];
                        range.Value = numofLineB[plusP];
                    }
                    for (int plusP = 0; plusP < numofLineC.Length; plusP++)
                    {
                        range = worksheet.Cells["C" + (addressY + plusP)];
                        range.Value = numofLineC[plusP];
                    }
                    for (int plusP = 0; plusP < numofLineD.Length; plusP++)
                    {
                        range = worksheet.Cells["D" + (addressY + plusP)];
                        range.Value = numofLineD[plusP];
                    }
                    for (int plusP = 0; plusP < numofLineE.Length; plusP++)
                    {
                        range = worksheet.Cells["E" + (addressY + plusP)];
                        range.Value = numofLineE[plusP];
                    }
                    plus = Math.Max(numofLineA.Length, numofLineB.Length);
                    plus = Math.Max(numofLineC.Length, plus);
                    plus = Math.Max(numofLineD.Length, plus);
                    plus = Math.Max(numofLineE.Length, plus);
                    addressY = addressY + plus - 1;
                }
                string outPath = "";
                Common.SaveExcelTemplate(workbook, "N5文法", "xls", out outPath);
                if (File.Exists(outPath))
                {
                    System.Diagnostics.Process.Start(outPath);
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                workbook.Close();
            }
        }

        private void OutPutExcel_N4Nguphap(DataTable adtData)
        {
            string templatePath = Common.GetTemplate("N4Nguphap_テンプレート.xls");
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(templatePath);
            try
            {
                SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets["N4Nguphap"];
                SpreadsheetGear.IRange range = null;
                //tb.Columns.Add("id");
                //tb.Columns.Add("maucau");
                //tb.Columns.Add("cachchia");
                //tb.Columns.Add("ynghia");
                //tb.Columns.Add("vidu");
                string dataA, dataB, dataC, dataD, dataE;
                string[] numofLineA, numofLineB, numofLineC, numofLineD, numofLineE;
                for (int i = 0, addressY = 2, plus = 0; i < adtData.Rows.Count; i++, addressY++, plus = 0)
                {
                    dataA = adtData.Rows[i]["id"].ToString();
                    dataB = adtData.Rows[i]["maucau"].ToString();
                    dataC = adtData.Rows[i]["cachchia"].ToString();
                    dataD = adtData.Rows[i]["ynghia"].ToString();
                    dataE = adtData.Rows[i]["vidu"].ToString();
                    numofLineA = dataA.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineB = dataB.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineC = dataC.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineD = dataD.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    numofLineE = dataE.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    for (int plusP = 0; plusP < numofLineA.Length; plusP++)
                    {
                        range = worksheet.Cells["A" + (addressY + plusP)];
                        range.Value = numofLineA[plusP];
                    }
                    for (int plusP = 0; plusP < numofLineB.Length; plusP++)
                    {
                        range = worksheet.Cells["B" + (addressY + plusP)];
                        range.Value = numofLineB[plusP];
                    }
                    for (int plusP = 0; plusP < numofLineC.Length; plusP++)
                    {
                        range = worksheet.Cells["C" + (addressY + plusP)];
                        range.Value = numofLineC[plusP];
                    }
                    for (int plusP = 0; plusP < numofLineD.Length; plusP++)
                    {
                        range = worksheet.Cells["D" + (addressY + plusP)];
                        range.Value = numofLineD[plusP];
                    }
                    for (int plusP = 0; plusP < numofLineE.Length; plusP++)
                    {
                        range = worksheet.Cells["E" + (addressY + plusP)];
                        range.Value = numofLineE[plusP];
                    }
                    plus = Math.Max(numofLineA.Length, numofLineB.Length);
                    plus = Math.Max(numofLineC.Length, plus);
                    plus = Math.Max(numofLineD.Length, plus);
                    plus = Math.Max(numofLineE.Length, plus);
                    addressY = addressY + plus - 1;
                }
                string outPath = "";
                Common.SaveExcelTemplate(workbook, "N4文法", "xls", out outPath);
                if (File.Exists(outPath))
                {
                    System.Diagnostics.Process.Start(outPath);
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                workbook.Close();
            }
        }
        private async void btn1000共通単語_Click(object sender, EventArgs e)
        {
            int iPage = 1;
            string res = await GetHtmlContent(iPage, "tu-vung-tieng-nhat-1000-tu-thong-dung-nhat.html");
            DataTable data = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.Find1000共通単語);
            while (iPage < 17)
            {
                iPage++;
                res = await GetHtmlContent(iPage, "tu-vung-tieng-nhat-1000-tu-thong-dung-nhat.html");
                DataTable data1 = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.Find1000共通単語);
                data.Merge(data1);
            }
            await ExtractData_3000共通単語(data);
            OutPutExcel_1000共通単語(data);
        }

        private async void btn2000共通単語_Click(object sender, EventArgs e)
        {
            int iPage = 1;
            string res = await GetHtmlContent(iPage, "2000-tu-vung-tieng-nhat-thong-dung.html");
            DataTable data = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.Find2000共通単語);
            while (iPage < 17)
            {
                iPage++;
                res = await GetHtmlContent(iPage, "2000-tu-vung-tieng-nhat-thong-dung.html");
                DataTable data1 = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.Find2000共通単語);
                data.Merge(data1);
            }
            await ExtractData_3000共通単語(data);
            OutPutExcel_2000共通単語(data);
        }

        private async void btnNguphapN5_Click(object sender, EventArgs e)
        {
            int iPage = 1;
            string res = await GetHtmlContent(iPage, "ngu-phap-tieng-nhat-n5-tong-hop.html");
            DataTable data = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.FindN5Nguphap);
            while (iPage < 12)
            {
                iPage++;
                res = await GetHtmlContent(iPage, "ngu-phap-tieng-nhat-n5-tong-hop.html");
                DataTable data1 = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.FindN5Nguphap);
                data.Merge(data1);
            }
            OutPutExcel_N5Nguphap(data);
        }

        private async void btnNguphapN4_Click(object sender, EventArgs e)
        {
            int iPage = 1;
            string res = await GetHtmlContent(iPage, "ngu-phap-n4-sach-mimi-bai-"+iPage+".html");
            DataTable data = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.FindN4Nguphap);
            while (iPage < 18)
            {
                iPage++;
                res = await GetHtmlContent(iPage, "ngu-phap-n4-sach-mimi-bai-" + iPage + ".html");
                DataTable data1 = await HtmlSourceCode.GetCode(res, iPage, HtmlSourceCode.FindN4Nguphap);
                data.Merge(data1);
            }
            OutPutExcel_N4Nguphap(data);
        }
    }

    public class HtmlSourceCode
    {
        public async static Task<string> Code(string Url, int aiPage, Func<HtmlDocument, int, string> findCode)
        {
            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            Task<string> t = Task.Run(() =>
            {
                HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(Url);
                myRequest.Method = "GET";
                int loop = 10;
                WebResponse myResponse = null;
                while (loop > 0 && myResponse == null)
                {
                    try
                    {
                        myResponse = myRequest.GetResponse();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        loop--;
                    }
                }
                if (myResponse == null)
                {
                    throw new System.Net.WebException(Url);
                }
                StreamReader sr = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.UTF8);
                string result = sr.ReadToEnd();
                sr.Close();
                myResponse.Close();
                return result;
            }).ContinueWith<HtmlDocument>((r) =>
            {
                WebBrowser browser = new WebBrowser();
                browser.ScriptErrorsSuppressed = true;
                browser.DocumentText = r.Result;
                browser.Document.OpenNew(true);
                browser.Document.Write(r.Result);
                browser.Refresh();
                browser.Document.Title = Url;
                return browser.Document;
            }, scheduler).ContinueWith<string>((r) =>
            {
                return findCode(r.Result, aiPage);
            }, scheduler);
            string res = await t;
            return res;
        }

        public async static Task GetCode(string content, DataTable adtOrgTable, Action<HtmlDocument, DataTable> updateDataTale)
        {
            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            Task t = Task.Run(() =>
            {
                content = content.Replace("Start your code after this line --><!-- End your code before this line", "");
                return content;
            }).ContinueWith<HtmlDocument>((r) =>
            {
                WebBrowser browser = new WebBrowser();
                browser.ScriptErrorsSuppressed = true;
                browser.DocumentText = "<html><head><title>Google</title></head><body></body</html>";
                browser.Document.OpenNew(true);
                browser.Document.Write(r.Result);
                browser.Refresh();
                return browser.Document;
            }, scheduler).ContinueWith((r) =>
            {
                updateDataTale(r.Result, adtOrgTable);
            }, scheduler);
            await t;
        }


        public async static Task<DataTable> GetCode(string content, int aiPage, Func<HtmlDocument, int, DataTable> findCode)
        {
            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            Task<DataTable> t = Task.Run(() =>
            {
                content = content.Replace("Start your code after this line --><!-- End your code before this line", "");
                return content;
            }).ContinueWith<HtmlDocument>((r) =>
            {
                WebBrowser browser = new WebBrowser();
                browser.ScriptErrorsSuppressed = true;
                browser.DocumentText = "<html><head><title>Google</title></head><body></body</html>";
                browser.Document.OpenNew(true);
                browser.Document.Write(r.Result);
                browser.Refresh();
                return browser.Document;
            }, scheduler).ContinueWith<DataTable>((r) =>
            {
                return findCode(r.Result, aiPage);
            }, scheduler);
            DataTable res = await t;
            return res;
        }

        private static HtmlElement GetElementById(HtmlElement search, string id)
        {
            HtmlElement el = search.Children.OfType<HtmlElement>().Where(r => r.Id == id).FirstOrDefault();
            if (el == null)
            {
                for (int i = 0; i < search.Children.Count; i++)
                {
                    HtmlElement elc = GetElementById(search.Children[i], id);
                    if (elc != null)
                    {
                        return elc;
                    }
                }
            }
            return el;
        }

        private static HtmlElement GetElementByTagName(HtmlElement search, string tagName)
        {
            HtmlElement el = search.Children.OfType<HtmlElement>().Where(r => r.TagName.ToLower() == tagName).FirstOrDefault();
            if (el == null)
            {
                for (int i = 0; i < search.Children.Count; i++)
                {
                    HtmlElement elc = GetElementByTagName(search.Children[i], tagName);
                    if (elc != null)
                    {
                        return elc;
                    }
                }
            }
            return el;
        }

        private static HtmlElement GetBody(HtmlDocument doc)
        {
            HtmlElement body = doc.All[1].Children.OfType<HtmlElement>().Where(r => r.TagName.ToLower() == "body").FirstOrDefault();
            if (body == null)
            {
                for (int i = 0; i < doc.All[1].Children.Count; i++)
                {
                    GetElementByTagName(doc.All[1].Children[i], "body");
                }
            }
            return body;
        }
        public static DataTable Find3000共通単語(HtmlDocument doc, int aiPage)
        {
            int iDay = 6 * (aiPage - 1) + 1;
            DataTable tb = new DataTable();
            tb.Columns.Add("id");
            tb.Columns.Add("jp");
            tb.Columns.Add("read");
            tb.Columns.Add("vi");
            tb.Columns.Add("innerText");
            tb.Columns.Add("outerHtml");
            tb.Columns.Add("tooltipText");
            char[] splitCharTranslate = { ':' };
            char[] splitCharReading = { '.', '(', ')' };

            //HtmlElement body = GetBody(doc);
            //HtmlElement item = GetElementById(body, "tu-vung-tieng-nhat-ngay-thu-" + iDay);
            HtmlElement item = doc.GetElementById("tu-vung-tieng-nhat-ngay-thu-" + iDay);
            if (item == null)
            {
                goto output;
            }
            bool continueSearch = true;
            HtmlElement p = item.Parent;
            string sInnerText = "";
            string sInnerTextStartChar = "";
            string sOuterHtml;
            string sdatatooltip = "";
            string sdatatooltips = "";
            while (continueSearch)
            {
                p = p.NextSibling;
                if (p == null)
                {
                    continueSearch = false;
                    break;
                }
                if (p.TagName.ToLower() != "p")
                {
                    continue;
                }
                sInnerText = p.InnerText;
                if (string.IsNullOrEmpty(sInnerText))
                {
                    continue;
                }
                sInnerTextStartChar = sInnerText.Substring(0, Math.Max(sInnerText.IndexOf('.'), 0));
                if (sInnerTextStartChar.Length == 0 || !Regex.IsMatch(sInnerTextStartChar, @"\d+"))
                {
                    continue;
                }
                //get data
                sOuterHtml = p.OuterHtml;
                DataRow newRow = tb.NewRow();
                string[] datas = p.InnerText.Split(splitCharTranslate, StringSplitOptions.RemoveEmptyEntries);
                if (datas.Length >= 2)
                {
                    string[] japanDatas = datas[0].Split(splitCharReading, StringSplitOptions.RemoveEmptyEntries);
                    if (japanDatas.Length >= 2)
                    {
                        newRow[0] = japanDatas[0];
                        newRow[1] = japanDatas[1];
                    }
                    if (japanDatas.Length >= 3)
                    {
                        newRow[2] = japanDatas[2];
                    }
                    newRow[3] = datas[1];
                    newRow[4] = sInnerText;
                    newRow[5] = sOuterHtml;
                    sdatatooltips = string.Format(@"<tr><td id=""{0}"">", japanDatas[0]);
                    foreach (HtmlElement elmentA in p.Children)
                    {
                        if (elmentA.TagName.ToLower() != "a")
                        {
                            continue;
                        }
                        sdatatooltip = elmentA.GetAttribute("data-cmtooltip");
                        if (!string.IsNullOrEmpty(sdatatooltip))
                        {
                            sdatatooltip = sdatatooltip + @"<br />";
                            sdatatooltips += sdatatooltip;
                        }

                    }
                    sdatatooltips += "</td></tr>";
                    newRow[6] = sdatatooltips;
                    goto addRow;
                }

                newRow[4] = sInnerText;
                newRow[5] = sOuterHtml;

            addRow:
                tb.Rows.Add(newRow);
            }

        output:
            return tb;
        }

        public static DataTable Find1000共通単語(HtmlDocument doc, int aiPage)
        {
            int iDay = 6 * (aiPage - 1) + 1;
            DataTable tb = new DataTable();
            tb.Columns.Add("id");
            tb.Columns.Add("jp");
            tb.Columns.Add("read");
            tb.Columns.Add("vi");
            tb.Columns.Add("innerText");
            tb.Columns.Add("outerHtml");
            tb.Columns.Add("tooltipText");
            char[] splitCharTranslate = { ':' };
            char[] splitCharReading = { '．', '(', ')', ']' };

            HtmlElement item = doc.GetElementById("ngay-" + iDay);
            if (item == null)
                item = doc.GetElementById("ngay-thu-" + iDay);
            if (item == null)
            {
                if (aiPage==6)
                {
                    item = doc.GetElementById("1000-tu-vung-tieng-nhat-thong-dung-nhat-tuan-6");
                    if (item!=null)
                    {
                        goto search;
                    }
                }
                goto output;
            }
            search:
            bool continueSearch = true;
            HtmlElement p = item.Parent;
            string sInnerText = "";
            string sInnerTextStartChar = "";
            string sOuterHtml;
            string sdatatooltip = "";
            string sdatatooltips = "";
            while (continueSearch)
            {
                p = p.NextSibling;
                if (p == null)
                {
                    continueSearch = false;
                    break;
                }
                if (p.TagName.ToLower() != "p")
                {
                    continue;
                }
                sInnerText = p.InnerText;
                if (string.IsNullOrEmpty(sInnerText))
                {
                    continue;
                }
                sInnerTextStartChar = sInnerText.Substring(0, Math.Max(sInnerText.IndexOf('．'), 0));
                if (sInnerTextStartChar.Length == 0 || !Regex.IsMatch(sInnerTextStartChar, @"\d+"))
                {
                    continue;
                }
                //get data
                sOuterHtml = p.OuterHtml;
                DataRow newRow = tb.NewRow();
                string[] datas = p.InnerText.Split(splitCharTranslate, StringSplitOptions.RemoveEmptyEntries);
                if (datas.Length >= 2)
                {
                    string[] japanDatas = datas[0].Split(splitCharReading, StringSplitOptions.RemoveEmptyEntries);
                    if (japanDatas.Length >= 2)
                    {
                        newRow[0] = japanDatas[0];
                        newRow[1] = japanDatas[1];
                    }
                    if (japanDatas.Length >= 3)
                    {
                        newRow[2] = japanDatas[2];
                    }
                    newRow[3] = datas[1];
                    newRow[4] = sInnerText;
                    newRow[5] = sOuterHtml;
                    sdatatooltips = string.Format(@"<tr><td id=""{0}"">", japanDatas[0]);
                    foreach (HtmlElement elmentA in p.Children)
                    {
                        if (elmentA.TagName.ToLower() != "a")
                        {
                            continue;
                        }
                        sdatatooltip = elmentA.GetAttribute("data-cmtooltip");
                        if (!string.IsNullOrEmpty(sdatatooltip))
                        {
                            sdatatooltip = sdatatooltip + @"<br />";
                            sdatatooltips += sdatatooltip;
                        }

                    }
                    sdatatooltips += "</td></tr>";
                    newRow[6] = sdatatooltips;
                    goto addRow;
                }

                newRow[4] = sInnerText;
                newRow[5] = sOuterHtml;

            addRow:
                tb.Rows.Add(newRow);
            }

        output:
            return tb;
        }

        public static DataTable Find2000共通単語(HtmlDocument doc, int aiPage)
        {
            int iDay = 6 * (aiPage - 1) + 101;
            DataTable tb = new DataTable();
            tb.Columns.Add("id");
            tb.Columns.Add("jp");
            tb.Columns.Add("read");
            tb.Columns.Add("vi");
            tb.Columns.Add("innerText");
            tb.Columns.Add("outerHtml");
            tb.Columns.Add("tooltipText");
            char[] splitCharTranslate = { ':' };
            char[] splitCharReading = { '[', ']' };
            int irowIndentity = 2001 + 60 * (aiPage - 1);
            HtmlElement item = doc.GetElementById("tu-vung-tieng-nhat-thong-dung-ngay-" + iDay);
            if (item == null)
            {
                goto output;
            }
            bool continueSearch = true;
            HtmlElement ol = item.Parent;
            HtmlElement p = null;
            string sInnerText = "";
            string sOuterHtml;
            string sdatatooltip = "";
            string sdatatooltips = "";
            while (continueSearch)
            {
                ol = ol.NextSibling;
                if (ol == null)
                {
                    continueSearch = false;
                    break;
                }
                if (ol.TagName.ToLower() != "ol")
                {
                    continue;
                }
                for (int i = 0; i < ol.Children.Count; i++)
                {
                    p = ol.Children[i];
                    if (p.Children.Count > 0)
                    {
                        p = p.FirstChild;
                    }
                    sInnerText = p.InnerText;
                    if (string.IsNullOrEmpty(sInnerText))
                    {
                        continue;
                    }

                    //get data
                    sOuterHtml = p.OuterHtml;
                    DataRow newRow = tb.NewRow();
                    newRow[0] = irowIndentity;
                    string[] datas = p.InnerText.Split(splitCharTranslate, StringSplitOptions.RemoveEmptyEntries);
                    if (datas.Length >= 2)
                    {
                        string[] japanDatas = datas[0].Split(splitCharReading, StringSplitOptions.RemoveEmptyEntries);
                        if (japanDatas.Length >= 1)
                        {
                            newRow[1] = japanDatas[0];
                        }
                        if (japanDatas.Length >= 2)
                        {
                            newRow[2] = japanDatas[1];
                        }
                        newRow[3] = datas[1];
                        newRow[4] = sInnerText;
                        newRow[5] = sOuterHtml;
                        sdatatooltips = string.Format(@"<tr><td id=""{0}"">", irowIndentity);
                        foreach (HtmlElement elmentA in p.Children)
                        {
                            if (elmentA.TagName.ToLower() != "a")
                            {
                                continue;
                            }
                            sdatatooltip = elmentA.GetAttribute("data-cmtooltip");
                            if (!string.IsNullOrEmpty(sdatatooltip))
                            {
                                sdatatooltip = sdatatooltip + @"<br />";
                                sdatatooltips += sdatatooltip;
                            }

                        }
                        sdatatooltips += "</td></tr>";
                        newRow[6] = sdatatooltips;
                        goto addRow;
                    }

                    newRow[4] = sInnerText;
                    newRow[5] = sOuterHtml;

                addRow:
                    tb.Rows.Add(newRow);
                    irowIndentity++;
                }
            }

        output:
            return tb;
        }

        public static DataTable FindN5Nguphap(HtmlDocument doc, int aiPage)
        {
            int iDay = 5 * (aiPage - 1) + 1;
            DataTable tb = new DataTable();
            tb.Columns.Add("id");
            tb.Columns.Add("maucau");
            tb.Columns.Add("cachchia");
            tb.Columns.Add("ynghia");
            tb.Columns.Add("vidu");

            so:
            if (iDay > 5*aiPage)
            {
                goto output;
            }
            HtmlElement item = doc.GetElementById("cau-truc-so-" + iDay);
            if (item == null)
            {
                iDay++;
                goto so;
            }
            bool continueSearch = true;
            HtmlElement p = item.Parent;
            StringBuilder sbTextPart = new StringBuilder();
            string sInnerText;
            int columnIndex = 1;
            DataRow newRow = tb.NewRow();
            newRow[0] = iDay;
            while (continueSearch)
            {
                p = p.NextSibling;
                if (p == null)
                {
                    continueSearch = false;
                    break;
                }
                if (p.TagName.ToLower() != "p")
                {
                    if (p.TagName.ToLower() == "h4" && p.FirstChild != null && !string.IsNullOrEmpty(p.FirstChild.GetAttribute("id")) && p.FirstChild.GetAttribute("id").ToLower() == "cau-truc-so-" + (iDay + 1))
                    {
                        newRow[columnIndex] = sbTextPart.ToString();
                        sbTextPart.Clear();
                        continueSearch = false;
                        break;
                    }
                    continue;
                }
               
                sInnerText = p.InnerText;
                if (string.IsNullOrEmpty(sInnerText))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sInnerText) && sInnerText == "Cách chia :")
                {
                    newRow[columnIndex] = sbTextPart.ToString();
                    columnIndex = 2;
                    sbTextPart.Clear();
                    continue;
                }
                if (!string.IsNullOrEmpty(sInnerText) && sInnerText=="Giải thích ý nghĩa :")
                {
                    newRow[columnIndex] = sbTextPart.ToString();
                    columnIndex = 3;
                    sbTextPart.Clear();
                    continue;
                }
                if (!string.IsNullOrEmpty(sInnerText) && sInnerText == "Ví dụ và ý nghĩa ví dụ :")
                {
                    newRow[columnIndex] = sbTextPart.ToString();
                    columnIndex = 4;
                    sbTextPart.Clear();
                    continue;
                }

                if (iDay == 5*aiPage && sInnerText.StartsWith("Trên đây là nội dung tổng hợp ngữ pháp tiếng Nhật N"))
                {
                    newRow[columnIndex] = sbTextPart.ToString();
                    sbTextPart.Clear();
                    continueSearch = false;
                    break;
                }
                sbTextPart.AppendLine(sInnerText);
            }
            tb.Rows.Add(newRow);
            iDay++;
            goto so;
        output:
            return tb;
        }

        public static DataTable FindN4Nguphap(HtmlDocument doc, int aiPage)
        {
            int iDay = 5 * (aiPage - 1) + 1;
            DataTable tb = new DataTable();
            tb.Columns.Add("id");
            tb.Columns.Add("maucau");
            tb.Columns.Add("cachchia");
            tb.Columns.Add("ynghia");
            tb.Columns.Add("vidu");

        so:
            if (iDay > 5 * aiPage)
            {
                goto output;
            }
            HtmlElement item = doc.GetElementById("cau-truc-so-" + iDay);
            if (item == null)
            {
                iDay++;
                goto so;
            }
            bool continueSearch = true;
            HtmlElement p = item.Parent;
            StringBuilder sbTextPart = new StringBuilder();
            string sInnerText;
            int columnIndex = 1;
            DataRow newRow = tb.NewRow();
            newRow[0] = iDay;
            int iCountCachchia = 0;
            while (continueSearch)
            {
                p = p.NextSibling;
                if (p == null)
                {
                    continueSearch = false;
                    break;
                }
                if (p.TagName.ToLower() != "p")
                {
                    if (p.TagName.ToLower() == "h4" && p.FirstChild != null && !string.IsNullOrEmpty(p.FirstChild.GetAttribute("id")) && p.FirstChild.GetAttribute("id").ToLower() == "cau-truc-so-" + (iDay + 1))
                    {
                        newRow[columnIndex] = sbTextPart.ToString();
                        sbTextPart.Clear();
                        continueSearch = false;
                        break;
                    }
                    continue;
                }

                sInnerText = p.InnerText;
                if (string.IsNullOrEmpty(sInnerText))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sInnerText) && (sInnerText.Trim() == "Cách chia :" || sInnerText.Trim() == "Cách chia "+(iCountCachchia+2)+":"))
                {
                    if (iCountCachchia > 0)
                    {
                        DataRow newRow1 = tb.NewRow();
                        newRow1.ItemArray = newRow.ItemArray;
                        tb.Rows.Add(newRow1);
                    }
                    iCountCachchia++;
                    newRow[columnIndex] = sbTextPart.ToString();
                    columnIndex = 2;
                    sbTextPart.Clear();
                    continue;
                }
                if (!string.IsNullOrEmpty(sInnerText) && (sInnerText.Trim() == "Ý nghĩa :"))
                {
                    newRow[columnIndex] = sbTextPart.ToString();
                    columnIndex = 3;
                    sbTextPart.Clear();
                    continue;
                }
                if (!string.IsNullOrEmpty(sInnerText) && sInnerText.Trim() == "Ví dụ minh họa :")
                {
                    newRow[columnIndex] = sbTextPart.ToString();
                    columnIndex = 4;
                    sbTextPart.Clear();
                    continue;
                }

                if (iDay == 5 * aiPage && sInnerText.StartsWith("Trên đây là nội dung các cấu trúc Ngữ pháp N"))
                {
                    newRow[columnIndex] = sbTextPart.ToString();
                    sbTextPart.Clear();
                    continueSearch = false;
                    break;
                }
                sbTextPart.AppendLine(sInnerText);
            }
            tb.Rows.Add(newRow);
            iCountCachchia = 0;
            iDay++;
            goto so;
        output:
            return tb;
        }

        public static void ExtractData_3000共通単語(HtmlDocument doc, DataTable adtOrgTable)
        {
            HtmlElement p = null;
            int iTemp = 0;
            HtmlElement htmlelementTBODY = doc.Body.FirstChild.FirstChild;
            for (int i = 0; i < htmlelementTBODY.Children.Count; i++)
            {
                p = htmlelementTBODY.Children[i];
                if (p.TagName.ToLower() != "tr")
                {
                    continue;
                }

                if (int.TryParse(p.FirstChild.Id, out iTemp))
                {
                    foreach (DataRow rowFinded in adtOrgTable.Select("id='" + iTemp + "'"))
                    {
                        rowFinded["tooltipText"] = p.FirstChild.InnerText;
                        break;
                    }
                }
            }
        }
    }
}
