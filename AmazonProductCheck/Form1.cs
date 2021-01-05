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
using System.Threading.Tasks;
using System.IO;

namespace AmazonProductCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        async private void Form1_Load(object sender, EventArgs e)
        {
            Process[] firefoxDriverProcesses = Process.GetProcessesByName("ChromeDriver");
            foreach (var firefoxDriverProcesse in firefoxDriverProcesses)
            {
                firefoxDriverProcesse.Kill();
            }

            WaitDialog wd = new WaitDialog();
            wd.TopMost = true;
            wd.StartPosition = FormStartPosition.CenterScreen;
            wd.Show();
            this.Enabled = false;

            string url = "https://www.amazon.co.jp/gp/bestsellers/sports/15337751/ref=zg_bs_unv_sg_2_15314601_1";

            TreeView t1 = new TreeView();
            t1.Nodes.Add(url, "自転車");
            await Task.Run(() => GetCategory(url, t1.Nodes[0]));

            TreeNodeCollection myTreeNodeCollection = t1.Nodes;
            // Create an array of 'TreeNodes'.
            TreeNode[] myTreeNodeArray = new TreeNode[t1.Nodes.Count];
            // Copy the tree nodes to the 'myTreeNodeArray' array.
            t1.Nodes.CopyTo(myTreeNodeArray, 0);
            // Remove all the tree nodes from the 'myTreeViewBase' TreeView.
            t1.Nodes.Clear();
            // Add the 'myTreeNodeArray' to the 'myTreeViewCustom' TreeView.
            treeView1.Nodes.AddRange(myTreeNodeArray);

            wd.Close();
            this.Enabled = true;
            this.Show();
        }

        void LookupChecks(TreeNodeCollection nodes, List<TreeNode> list)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Checked)
                    list.Add(node);

                LookupChecks(node.Nodes, list);
            }
        }

        private void GetCategory(string url, TreeNode parent)
        {
            HtmlAgilityPack.HtmlDocument document1 = new HtmlAgilityPack.HtmlDocument();
            string htmlCode1;
            using (WebClient client = new WebClient())
            {
                var htmlData1 = client.DownloadData(url);
                htmlCode1 = Encoding.UTF8.GetString(htmlData1);
                document1.LoadHtml(htmlCode1);
            }

            var nc4 = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//ul//ul//ul//span[@class='zg_selected']");
            var nc3 = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//ul//ul//span[@class='zg_selected']");
            var nc2 = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//ul//span[@class='zg_selected']");
            var nc1 = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//span[@class='zg_selected']");
            HtmlNodeCollection nc = null;
            if (nc4 != null)
            {
                nc = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//ul//ul//ul//ul//a");
            }
            else if (nc3 != null)
            {
                nc = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//ul//ul//ul//a");
            }
            else if (nc2 != null)
            {
                nc = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//ul//ul//a");
            }
            else if (nc1 != null)
            {
                nc = document1.DocumentNode.SelectNodes("//ul[@id='zg_browseRoot']//ul//ul//ul//a");
            }

            if (nc == null)
                return;
            HtmlNode[] nodes = nc.ToArray();
            foreach (HtmlNode item in nodes)
            {
                var child = parent.Nodes.Add(item.Attributes["href"].Value, item.InnerHtml);
                GetCategory(item.Attributes["href"].Value, child);
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
            if(MessageBox.Show("Do you really want to exit?","Confirm",MessageBoxButtons.OKCancel,MessageBoxIcon.Question) == DialogResult.OK)
                Environment.Exit(0);
        }
        private bool Amazon_Chrome(IWebDriver chrome, List<TreeNode> catlist)
        {
            try
            {
                foreach (var category in catlist)
                {
                    string categoryurl = string.Empty;
                    string categoryname = string.Empty;
                    string itemname = string.Empty;
                    string brandname = string.Empty;
                    string price = string.Empty;
                    string ASIN = string.Empty;
                    string EAN = string.Empty;

                    int numbSelect = 0;
                    numbSelect = Convert.ToInt32(textBox1.Text);

                    var _list = new List<String>();
                    categoryurl = category.Name;

                    chrome.Url = categoryurl;
                    Thread.Sleep(1000);

                    for (int z = 1; z <= numbSelect; z++)
                        {
                            if (z <= 50)
                            {
                            if (IsElementPresent(chrome, By.XPath("//*[@id='zg-ordered-list']/li[" + z + "]/span/div/span/a")))
                            {
                                var itemcode_url = chrome.FindElement(By.XPath("//*[@id='zg-ordered-list']/li[" + z + "]/span/div/span/a")).GetAttribute("href");
                                _list.Add(itemcode_url);
                            }
                            else
                                break;
                            }
                            else
                            {
                                chrome.FindElement(By.XPath("//*[@id='zg-center-div']/div[2]/div/ul/li[3]/a")).Click();
                                Thread.Sleep(1000);
                                for(int w=1; w<=numbSelect-50; w++)
                                {
                                if (IsElementPresent(chrome, By.XPath("//*[@id='zg-ordered-list']/li[" + w + "]/span/div/span/a")))
                                {
                                    var itemcode_url = chrome.FindElement(By.XPath("//*[@id='zg-ordered-list']/li[" + w + "]/span/div/span/a")).GetAttribute("href");
                                    _list.Add(itemcode_url);
                                }
                                else
                                    break;
                            }
                                break;
                            }
                        }
 
                    for (int i = 0; i < _list.Count; i++)
                    {
                        categoryname = category.Text;

                        chrome.Url = _list[i];
                        Thread.Sleep(1000);

                    string result = string.Empty;
                    if (IsElementPresent(chrome, By.XPath("/html/body/center/span")))
                    {
                        result = chrome.FindElement(By.XPath("/html/body/center/span")).Text;
                    }
                    if (result.Contains("年齢確認"))
                    {
                        chrome.FindElement(By.XPath("/html/body/center/div[1]/a")).Click();
                    }

                    itemname = chrome.FindElement(By.Id("imgTagWrapperId")).FindElement(By.TagName("img")).GetAttribute("alt");

                        if (IsElementPresent(chrome, By.XPath("//*[@id='bylineInfo']")))
                        {
                            brandname = chrome.FindElement(By.XPath("//*[@id='bylineInfo']")).Text;
                        }
                        else
                        {
                            brandname = "";
                        }
                        if (brandname.Contains("ブランド: "))
                        {
                            brandname = brandname.Replace("ブランド: ", "");
                        }
                        else if(brandname.Contains("Brand:"))
                        {
                            brandname = brandname.Replace("Brand:", "");
                        }
                        else
                        {
                            brandname ="";
                        }

                        if (IsElementPresent(chrome, By.Id("price_inside_buybox")))
                        {
                            price = chrome.FindElement(By.Id("price_inside_buybox")).Text;
                            if (price.Contains("¥"))
                            {
                                price = price.Replace("¥", "");
                            }
                        }
                        else if (IsElementPresent(chrome, By.Id("newBuyBoxPrice")))
                        {
                            price = chrome.FindElement(By.Id("newBuyBoxPrice")).Text;
                            if (price.Contains("¥"))
                            {
                                price = price.Replace("¥", "");
                            }
                        }
                        else
                        {
                            price = "0";
                        }

                        if (IsElementPresent(chrome, By.XPath("//*[contains(text(),'ASIN')]")))
                        {
                            if (IsElementPresent(chrome, By.XPath("//*[@id='detailBullets_feature_div']/ul")))
                            {
                                if (chrome.FindElement(By.XPath("//*[@id='detailBullets_feature_div']/ul")).Text.Contains("ASIN"))
                                {
                                    ASIN = chrome.FindElement(By.XPath("//*[@id='detailBullets_feature_div']/ul")).Text;
                                }
                            }
                            else if (IsElementPresent(chrome, By.XPath("//*[@id='productDetails_detailBullets_sections1']/tbody/tr[1]/td")))
                            {
                                if (chrome.FindElement(By.XPath("//*[@id='productDetails_detailBullets_sections1']/tbody/tr[1]")).Text.Contains("ASIN"))
                                {
                                    ASIN = chrome.FindElement(By.XPath("//*[@id='productDetails_detailBullets_sections1']/tbody/tr[1]/td")).Text;
                                }
                            }
                        }
                        if (ASIN.Contains("ASIN"))
                        {
                            ASIN = ASIN.Substring(ASIN.IndexOf("ASIN : ") + 7).Substring(0, 10);
                        }


                        chrome.Url = "https://sellercentral-japan.amazon.com/product-search/search?q=" + ASIN + "&ref_=xx_catadd_dnav_home";
                        Thread.Sleep(2000);


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
                            chrome.Close();
                            chrome.Quit();
                            MessageBox.Show(this, "Please login to EAN code!");
                            return false;
                        }
                        if (EAN.Contains("EAN:"))
                        {
                            EAN = EAN.Replace("EAN:", "");
                        }
                        else
                        {
                            EAN = "";
                        }
                        int j = 0;
                        j = i + 1;
                        DataResult(j.ToString(), categoryname, itemname, brandname, price, ASIN, EAN);
                    }
                }
                chrome.Close();
                chrome.Quit();
                DataTable dt = InsertData();
                ReleaseOutputFile(dt);
                
                Process[] firefoxDriverProcesses = Process.GetProcessesByName("ChromeDriver");
                foreach (var firefoxDriverProcesse in firefoxDriverProcesses)
                {
                    firefoxDriverProcesse.Kill();
                }
                return true;
            }
            catch (ThreadInterruptedException e)
            {
                Environment.Exit(0);
                return false;
            }
           
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
            //DataTable dtAmazonSKU = new DataTable();
            //dtAmazonSKU = GetAmazonSKU(EAN);
            //if (dtAmazonSKU.Rows.Count > 0)
            //{
            //    dtResult.Rows[dtResult.Rows.Count - 1]["AmazonSKU"] = dtAmazonSKU.Rows[0]["AmazonSKU"].ToString();
            //    DataTable dtMakerStatus = new DataTable();
            //    dtMakerStatus = GetDataFromMaker_Status(dtAmazonSKU.Rows[0]["AmazonSKU"].ToString());
            //    if (dtMakerStatus.Rows.Count > 0)
            //    {
            //        dtResult.Rows[dtResult.Rows.Count - 1]["下代"] = dtMakerStatus.Rows[0]["下代"].ToString();
            //    }

            //}
            //else
            //{
            //    dtResult.Rows[dtResult.Rows.Count - 1]["AmazonSKU"] = "";
            //    dtResult.Rows[dtResult.Rows.Count - 1]["下代"] = "";
            //}
            dtResult.Rows[dtResult.Rows.Count - 1]["取得日"] = System.DateTime.Now.ToString();
        }
        private void ReleaseOutputFile(DataTable dtResult)

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
        private DataTable InsertData()
        {
            dtResult.TableName = "test";
            System.IO.StringWriter writer = new System.IO.StringWriter();
            dtResult.WriteXml(writer, XmlWriteMode.WriteSchema, false);
            string result = writer.ToString();
            return Insert(result);
        }
        private DataTable Insert(string xml)
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("SP_Insert_AmazonSKU", con);
            try
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;
                cmd.Parameters.Add("@xml", SqlDbType.Xml).Value = xml;
                var adp = new SqlDataAdapter();
                adp.SelectCommand = cmd;
                cmd.Connection.Open();
                adp.Fill(dt);
                return dt;

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

        private void button2_Click(object sender, EventArgs e)
        {
            create_Table();

            var list = new List<TreeNode>();
            LookupChecks(treeView1.Nodes, list);

            Process[] chromeDriverProcesse = Process.GetProcessesByName("ChromeDriver");
            foreach (var chromeDriverProces in chromeDriverProcesse)
            {
                chromeDriverProces.Kill();
            }

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            string path = Environment.ExpandEnvironmentVariables("%LOCALAPPDATA%\\Google\\Chrome\\User Data");
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("user-data-dir=" + path);
            options.AddArguments("profile-directory=Default");
            options.AddArgument("start-maximized");
            IWebDriver Chrome = new ChromeDriver(service, options);

            if (!Amazon_Chrome(Chrome, list))
            {
                try
                {                 
                    Process[] chromeDriverProcesses = Process.GetProcessesByName("ChromeDriver");
                    foreach (var chromeDriverProcess in chromeDriverProcesses)
                    {
                        chromeDriverProcess.Kill();
                    }
                }
                catch
                { }
            }
            else
            {
                try
                {
                    Chrome.Close();
                    Chrome.Quit();
                    Process[] firefoxDriverProcesses = Process.GetProcessesByName("ChromeDriver");
                    foreach (var firefoxDriverProcesse in firefoxDriverProcesses)
                    {
                        firefoxDriverProcesse.Kill();
                    }
                }
                catch
                { }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process.Start(Path.GetDirectoryName(@"D:\New PJ\Amazon\Export_XML\"));
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var textbox = sender as TextBox;
            int value;
            if (int.TryParse(textbox.Text, out value))
            {
                if (value > 100)
                    textbox.Text = "100";
                else if (value < 0)
                    textbox.Text = "0";
            }
        }
    }
}


