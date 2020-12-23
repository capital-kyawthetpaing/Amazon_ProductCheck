using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data.SqlClient;
using System.Data;

namespace AmazonProductCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            create_Table();
            HtmlAgilityPack.HtmlDocument document1 = new HtmlAgilityPack.HtmlDocument();
            string htmlCode1;
            using (WebClient client = new WebClient())
            {
                var htmlData1 = client.DownloadData("https://www.amazon.co.jp/gp/bestsellers/sports/15337751/ref=zg_bs_nav_sg_1_sg");
                htmlCode1 = Encoding.UTF8.GetString(htmlData1);
            }
            document1.LoadHtml(htmlCode1);

            if (document1.DocumentNode.SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//li//a") != null)
            {
                HtmlNode[] nodes = document1.DocumentNode
                                                .SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//li//a").ToArray();
                treeView1.BeginUpdate();
                foreach (HtmlNode item in nodes)
                {
                    var parentNode = treeView1.Nodes.Add(item.Attributes["href"].Value, item.InnerHtml);

                    string hrefValue = item.Attributes["href"].Value;
                    HtmlAgilityPack.HtmlDocument document2 = new HtmlAgilityPack.HtmlDocument();
                    string htmlCode2;
                    using (WebClient client = new WebClient())
                    {
                        var htmlData2 = client.DownloadData(hrefValue);
                        htmlCode2 = Encoding.UTF8.GetString(htmlData2);
                    }
                    document2.LoadHtml(htmlCode2);

                    if (document2.DocumentNode.SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//ul//li//a") != null)
                    {
                        HtmlNode[] nodes2 = document2.DocumentNode
                                                        .SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//ul//li//a").ToArray();
                        if (nodes2 != null)
                        {
                            foreach (HtmlNode item1 in nodes2)
                            {
                                var subparent = parentNode.Nodes.Add(item1.Attributes["href"].Value, item1.InnerHtml);

                                string lnk = item1.Attributes["href"].Value;
                                HtmlAgilityPack.HtmlDocument document3 = new HtmlAgilityPack.HtmlDocument();
                                string htmlCode3;
                                using (WebClient client = new WebClient())
                                {
                                    var htmlData3 = client.DownloadData(lnk);
                                    htmlCode3 = Encoding.UTF8.GetString(htmlData3);
                                }
                                document3.LoadHtml(htmlCode3);
                                if(item1.InnerHtml == "子ども用自転車")
                                {
                                    if (document3.DocumentNode.SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//ul//li") != null)
                                    {
                                        HtmlNode[] nodes3 = document3.DocumentNode
                                                                        .SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//ul//li//a").ToArray();
                                        if (nodes3 != null)
                                        {
                                            foreach (HtmlNode item3 in nodes3)
                                            {
                                                subparent.Nodes.Add(item3.Attributes["href"].Value, item3.InnerHtml);
                                            }
                                        }
                                    }
                                }

                                if (document3.DocumentNode.SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//ul//ul//li") != null)
                                {
                                    HtmlNode[] nodes3 = document3.DocumentNode
                                                                    .SelectNodes("//div[@class='a-fixed-left-grid-col a-col-left']//ul[@id='zg_browseRoot']//ul//ul//ul//ul//ul//li//a").ToArray();
                                    if (nodes3 != null)
                                    {
                                        foreach (HtmlNode item3 in nodes3)
                                        {
                                            subparent.Nodes.Add(item3.Attributes["href"].Value, item3.InnerHtml);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                treeView1.EndUpdate();
            }

        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            // The code only executes if the user caused the checked state to change.
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count > 0)
                {
                    /* Calls the CheckAllChildNodes method, passing in the current 
                    Checked value of the TreeNode whose checked state changed. */
                    this.CheckAllChildNodes(e.Node, e.Node.Checked);
                }
            }
        }
        // Updates all child tree nodes recursively.
        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    // If the current node has child nodes, call the CheckAllChildsNodes method recursively.
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string categoryurl = string.Empty;
            string categoryname = string.Empty;
            categoryurl = treeView1.SelectedNode.Name;
            categoryname = treeView1.SelectedNode.Text;

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            string path = Environment.ExpandEnvironmentVariables("%LOCALAPPDATA%\\Google\\Chrome\\User Data");
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("user-data-dir=" + path);
            //options.AddArgument("historyItem ");
            options.AddArguments("profile-directory=Default");
            options.AddArgument("start-maximized");
            // options.AddArgument("--headlesss");
             IWebDriver Chrome = new ChromeDriver(service, options);

            //Firefox = new ChromeDriver(service, options);
            //ChromeDriverService serviceBB = ChromeDriverService.CreateDefaultService();
            //serviceBB.HideCommandPromptWindow = true;      
            //using (IWebDriver Chrome = new ChromeDriver(serviceBB))
            //{
            Chrome.Manage().Window.Maximize();
            if (!Amazon_Chrome(Chrome, categoryurl, categoryname))
                {
                    try
                    {
                        this.Refresh();
                        Refresh();
                        this.Hide();
                        Form1 ss = new Form1();
                        ss.Show();
                        Process[] chromeDriverProcesses = Process.GetProcessesByName("ChromeDriver");
                        foreach (var chromeDriverProcess in chromeDriverProcesses)
                        {
                            chromeDriverProcess.Kill();
                        }
                        return;
                    }
                    catch
                    {
                    }
               }
                else
                {
                    try
                    {
                        Chrome.Close();
                        Chrome.Quit();
                    }
                    catch
                    {
                    }
                }
           // }
        }
        private bool Amazon_Chrome(IWebDriver chrome, string categoryurl,string categoryname)
        {        
            try
            {               
                for (int i = 1; i <= 10; i++)
                {
                    if (i <= 50)
                    {
                        //chrome.Url = "https://www.amazon.co.jp/gp/bestsellers/sports/15314601/ref=zg_bs_nav_sg_2_15337751/358-4226875-7101947";
                        chrome.Url = categoryurl;
                        //Thread.Sleep(1000);


                       var itemcode_url = chrome.FindElement(By.XPath("//*[@id='zg-ordered-list']/li[" + i + "]/span/div/span/a")).GetAttribute("href");
                        chrome.Url = itemcode_url;
                            //Thread.Sleep(1000);

                            var itemname = chrome.FindElement(By.Id("imgTagWrapperId")).FindElement(By.TagName("img")).GetAttribute("alt");

                            var brandname = chrome.FindElement(By.XPath("//*[@id='productOverview_feature_div']/div/table/tbody/tr[1]/td[2]/span")).Text;

                            string price = string.Empty;
                            if (IsElementPresent(chrome, By.Id("price_inside_buybox")))
                            {
                                //do if exists
                                 price = chrome.FindElement(By.Id("price_inside_buybox")).Text;
                            }
                            else
                            {
                                //do if does not exists
                                price = "一時的に在庫切れ";
                            }
                            string ASIN = string.Empty;
                            if (IsElementPresent(chrome, By.XPath("//*[@id='productDetails_detailBullets_sections1']/tbody/tr[1]/td")))
                            {
                                //do if exists
                                 ASIN = chrome.FindElement(By.XPath("//*[@id='productDetails_detailBullets_sections1']/tbody/tr[1]/td")).Text;
                            }
                            else
                            {
                                //do if does not exists
                                 ASIN = chrome.FindElement(By.XPath("//*[@id='detailBullets_feature_div']/ul/li[4]/span/span[2]")).Text;
                            }
                            
                        chrome.Url = "https://sellercentral-japan.amazon.com/product-search/search?q=" + ASIN + "&ref_=xx_catadd_dnav_home";
                        //Thread.Sleep(1000);
                        string EAN = string.Empty;

                           
                             if (IsElementPresent(chrome, By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]")))
                            {
                                if (IsElementPresent(chrome, By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]/p[2]")))
                                {
                                    EAN= chrome.FindElement(By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]/p[2]")).Text;
                                }
                                else
                                {
                                    EAN = chrome.FindElement(By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]/p")).Text;
                                }

                            }
                            else
                            {
                                EAN = "";
                            }


                            if (EAN.Contains("EAN:"))
                            {
                                EAN = EAN.Replace("EAN:", "");
                            }
                            else
                            {
                                EAN = "";
                            }

                            DataResult(i.ToString(), categoryname, itemname,brandname,price,ASIN, EAN);

                    }
                    else
                    {
                        int j = 0;
                        j = i - 50;
                        //chrome.Url = "https://www.amazon.co.jp/-/en/gp/bestsellers/sports/15314601/ref=zg_bs_pg_2?ie=UTF8&pg=2";
                        chrome.Url = categoryurl + "&pg=2";

                        var itemcode_url = chrome.FindElement(By.XPath("//*[@id='zg-ordered-list']/li[" + i + "]/span/div/span/a")).GetAttribute("href");

                        chrome.Url = itemcode_url;


                        var itemname = chrome.FindElement(By.Id("imgTagWrapperId")).FindElement(By.TagName("img")).GetAttribute("alt");

                        var brandname = chrome.FindElement(By.XPath("//*[@id='productOverview_feature_div']/div/table/tbody/tr[1]/td[2]/span")).Text;

                        string price = string.Empty;
                        if (IsElementPresent(chrome, By.Id("price_inside_buybox")))
                        {
                            //do if exists
                            price = chrome.FindElement(By.Id("price_inside_buybox")).Text;
                        }
                        else
                        {
                            //do if does not exists
                            price = "一時的に在庫切れ";
                        }
                        string ASIN = string.Empty;
                        if (IsElementPresent(chrome, By.XPath("//*[@id='productDetails_detailBullets_sections1']/tbody/tr[1]/td")))
                        {
                            //do if exists
                            ASIN = chrome.FindElement(By.XPath("//*[@id='productDetails_detailBullets_sections1']/tbody/tr[1]/td")).Text;
                        }
                        else
                        {
                            //do if does not exists
                            ASIN = chrome.FindElement(By.XPath("//*[@id='detailBullets_feature_div']/ul/li[4]/span/span[2]")).Text;
                        }

                        chrome.Url = "https://sellercentral-japan.amazon.com/product-search/search?q=" + ASIN + "&ref_=xx_catadd_dnav_home";
                        //Thread.Sleep(1000);
                        string EAN = string.Empty;


                        if (IsElementPresent(chrome, By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]")))
                        {
                            if (IsElementPresent(chrome, By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]/p[2]")))
                            {
                                EAN = chrome.FindElement(By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]/p[2]")).Text;
                            }
                            else
                            {
                                EAN = chrome.FindElement(By.XPath("//*[@id='search-result']/div/kat-box/div/section[3]/div[1]/section[1]/div/section[1]/p")).Text;
                            }

                        }
                        else
                        {
                            EAN = "";
                        }


                        if (EAN.Contains("EAN:"))
                        {
                            EAN = EAN.Replace("EAN:", "");
                        }
                        else
                        {
                            EAN = "";
                        }

                        DataResult(i.ToString(), categoryname, itemname, brandname, price, ASIN, EAN);


                    }

                }
                chrome.Close();
                chrome.Quit();
                ReleaseOutputFile();
                //InsertData();
                Process[] firefoxDriverProcesses = Process.GetProcessesByName("geckodriver");
                foreach (var firefoxDriverProcesse in firefoxDriverProcesses)
                {
                    firefoxDriverProcesse.Kill();
                }
            }
            catch (ThreadInterruptedException e)
            {
                Environment.Exit(0);
            }


            return true;
        }
        private bool IsElementPresent(IWebDriver chrome,By by)
        {
            try
            {
                chrome.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        System.Data.DataTable dtResult;
        string outdir = @"D:\New PJ\Amazon\Export_XML";
        public static string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        private void create_Table()
        {
            dtResult = new System.Data.DataTable();
            dtResult.Columns.Add("サブカテゴリ名"); 
            dtResult.Columns.Add("順位");
            dtResult.Columns.Add("商品名"); 
            dtResult.Columns.Add("ブランド名");
            dtResult.Columns.Add("ASIN"); 
            dtResult.Columns.Add("EANCD");  
            dtResult.Columns.Add("参考価格");
            dtResult.Columns.Add("AmazonSKU");
            dtResult.Columns.Add("下代");
            dtResult.Columns.Add("取得日");
        }
        private void DataResult(string rank, string categoryname, string itemname,string brandname,string price,string ASIN, string EAN)
        {
            dtResult.Rows.Add();
            dtResult.Rows[dtResult.Rows.Count - 1]["サブカテゴリ名"] = categoryname;
            dtResult.Rows[dtResult.Rows.Count - 1]["順位"] = rank;
            dtResult.Rows[dtResult.Rows.Count - 1]["商品名"] = itemname;
            dtResult.Rows[dtResult.Rows.Count - 1]["ブランド名"] = brandname;
            dtResult.Rows[dtResult.Rows.Count - 1]["ASIN"] = ASIN;
            dtResult.Rows[dtResult.Rows.Count - 1]["EANCD"] = EAN;
            dtResult.Rows[dtResult.Rows.Count - 1]["参考価格"] = price;
            DataTable dtAmazonSKU = new DataTable();
            dtAmazonSKU = GetAmazonSKU(EAN);
            if (dtAmazonSKU.Rows.Count > 0)
            {
                dtResult.Rows[dtResult.Rows.Count - 1]["AmazonSKU"] = dtAmazonSKU.Rows[0]["AmazonSKU"].ToString();
                DataTable dtMakerStatus = new DataTable();
                dtMakerStatus = GetDataFromMaker_Status(dtAmazonSKU.Rows[0]["AmazonSKU"].ToString());
                if (dtMakerStatus.Rows.Count > 0)
                {
                    dtResult.Rows[dtResult.Rows.Count - 1]["下代"] = dtMakerStatus.Rows[0]["下代"].ToString();
                }

            }
            else
            {
                dtResult.Rows[dtResult.Rows.Count - 1]["AmazonSKU"] = "";
                dtResult.Rows[dtResult.Rows.Count - 1]["下代"] = "";
            }
            dtResult.Rows[dtResult.Rows.Count - 1]["取得日"] = System.DateTime.Now.ToString("MM/dd/yyyy");
        }
        private void ReleaseOutputFile()
        {
            ExportExcel(dtResult, outdir+ "\\" + DateTime.Now.ToString("yyyyMMddHHmmss").Replace("\\", "").Replace(" ", string.Empty).Replace("/", "").Replace(":", "") + "Amazon.xlsx");
               
        }
        private void ExportExcel(System.Data.DataTable dtOutput, string filename)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;

            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetBB;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheetBB = xlWorkBook.Worksheets.Add(misValue);

            xlWorkSheetBB.Name = "Amazon_Data";

            xlWorkSheetBB = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0; 
            int j = 0; 

            for (int k = 0; k < dtOutput.Columns.Count; k++)
            {
                xlWorkSheetBB.Cells[1, k + 1] = dtOutput.Columns[k].ToString();
            }
            if (dtOutput.Rows.Count > 0)
            {
                for (i = 0; i <= dtOutput.Rows.Count - 1; i++)
                {
                    for (j = 0; j <= dtOutput.Columns.Count - 1; j++)
                    {
                        xlWorkSheetBB.Cells[i + 2, j + 1] = dtOutput.Rows[i][j].ToString();
                    }
                }
            }          

            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(filename);
            xlWorkBook.Close(true);
            xlApp.Quit();

            releaseObject(xlWorkSheetBB);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        private void InsertData()
        {
            dtResult.TableName = "test";
            System.IO.StringWriter writer = new System.IO.StringWriter();
            dtResult.WriteXml(writer, XmlWriteMode.WriteSchema, false);
            string result = writer.ToString();
            Insert(result);
        }
        private void Insert(string xml)
        {
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("SP_Insert_AmazonSKU", con);
            try
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;
                cmd.Parameters.Add("@xml", SqlDbType.Xml).Value = xml;
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Connection.Close();
                cmd.Dispose();
            }

        }
        private static DataTable GetAmazonSKU(string EAN)
        {
            string sql = "Select AmazonSKU from Item_Sku where JANCD='"+EAN+"'";
            SqlDataAdapter da = new SqlDataAdapter();
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.CommandType = CommandType.Text;
            da.SelectCommand = cmd;
            da.SelectCommand.CommandType = CommandType.Text;
            DataTable dt = new DataTable();
            da.SelectCommand.Connection.Open();
            da.Fill(dt);
            da.SelectCommand.Connection.Close();
            return dt;
        }
        private static DataTable GetDataFromMaker_Status(string JANCD)
        {
            string sql = "Select 下代 from Maker_Status where JANコード='"+ JANCD+"'";
            SqlDataAdapter da = new SqlDataAdapter();
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.CommandType = CommandType.Text;
            da.SelectCommand = cmd;
            da.SelectCommand.CommandType = CommandType.Text;
            DataTable dt = new DataTable();
            da.SelectCommand.Connection.Open();
            da.Fill(dt);
            da.SelectCommand.Connection.Close();
            return dt;
        }
    }
}


