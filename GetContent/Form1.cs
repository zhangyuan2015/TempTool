using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace GetContent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var cookieStr = txtCookie.Text;
            if (string.IsNullOrEmpty(cookieStr))
            {
                txtRes.Text = "请填入Cookie";
                return;
            }
            CookieCollection cookies = new CookieCollection();
            foreach (var cookieItem in cookieStr.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var cookieItemArr = cookieItem.Split(new[] { '=' }, StringSplitOptions.RemoveEmptyEntries);
                if (cookieItemArr.Length == 1)
                    cookies.Add(new Cookie(cookieItemArr[0].Trim(), "", "/", "www.mrobao.com"));
                else if (cookieItemArr.Length == 2)
                    cookies.Add(new Cookie(cookieItemArr[0].Trim(), cookieItemArr[1].Trim(), "/", "www.mrobao.com"));
            }
            CookieContainer cookieContainer = new CookieContainer();
            cookieContainer.Add(cookies);

            WebHeaderCollection webHeaderCollection = new WebHeaderCollection();
            //webHeaderCollection.Add("Accept-Encoding", "gzip, deflate");
            webHeaderCollection.Add("Accept-Language", "zh-CN,zh;q=0.9");
            webHeaderCollection.Add("Upgrade-Insecure-Requests", "1");

            button1.Enabled = false;
            //Task.Run(() =>
            //{
            GetData(1, cookieContainer, webHeaderCollection);
            button1.Enabled = true;
            //});
        }

        public void Res(string res)
        {
            txtRes.Text += $"{DateTime.Now.ToShortTimeString()} - {res}{Environment.NewLine}";
        }

        int? totalPage = null;
        public void PageCount(int page, int totalPage)
        {
            txtPageCount.Text = $"{page} / {totalPage}";
        }

        int? totalRows = null;
        int orderCount = 0;
        public void OrderCount()
        {
            txtOrderCount.Text = $"{orderCount} / {totalRows}";
        }

        int oQ = 4;
        int pQ = 2;
        int pageSize = 10;
        List<订单> 订单集合 = new List<订单>();
        public void GetData(int pageIndex, CookieContainer cookieContainer, WebHeaderCollection webHeaderCollection)
        {
            try
            {
                if (pageIndex > totalPage)
                {
                    Res($"执行完毕");
                    return;
                }

                Res($"开始解析第 {pageIndex} 页");
                string pageParam = "";
                if (pageIndex > 1)
                {
                    int firstRow = (pageIndex - 1) * pageSize;
                    pageParam = $"firstRow={firstRow}&totalRows={totalRows}&";
                }

                var responseBody = HttpGet($"http://www.mrobao.com/main.php?{pageParam}m=product&s=admin_sellorder&key={txtKey.Text.Trim()}&buy_catid=&is_invoice=", cookieContainer, webHeaderCollection);
                if (responseBody.Contains("登录"))
                {
                    Res($"Cookie 错误，登录失败");
                    return;
                }
                var doc = new HtmlDocument();
                doc.LoadHtml(responseBody);

                //总页数
                var pageNodes = doc.DocumentNode.SelectNodes("//div[@class='pagination']/a");
                var pageNode = pageNodes[pageNodes.Count - 1/*倒数第二个-1*/- 1/*使用下标-1*/];
                if (totalPage == null)
                    totalPage = int.Parse(pageNode.InnerText.Trim('.'));
                PageCount(pageIndex, totalPage.Value);

                if (totalRows == null)
                {
                    var href = pageNode.GetAttributeValue("href", "");
                    totalRows = int.Parse(href.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault(a => a.Contains("totalRows")).Split(new[] { '=' }, StringSplitOptions.RemoveEmptyEntries)[1]);
                }

                //订单Node
                var orderNode = doc.DocumentNode.SelectSingleNode("//table[@class='table-list-style order']");
                //订单Node/Tr
                var orderTrNodes = orderNode.SelectNodes("tbody/tr");

                Res($"当前页订单数 {orderTrNodes.Count} ");

                订单 订单 = null;
                for (int i = 0; i < orderTrNodes.Count; i++)
                {
                    var orderTrNode = orderTrNodes[i];
                    if (i % oQ == 0)
                    {
                        订单 = new 订单();
                        orderCount++;
                        OrderCount();
                    }
                    else if (i % oQ == 1)
                    {
                        var 订单编号 = orderTrNode.SelectSingleNode("th/span[1]/span").InnerText.Trim();
                        订单.订单编号 = 订单编号;
                        Res($"解析订单 - {订单编号}");

                        var 下单时间 = orderTrNode.SelectSingleNode("th/span[2]/span").InnerText.Trim();
                        订单.下单时间 = 下单时间;
                        Res(下单时间);
                    }
                    else if (i % oQ == 2)
                    {
                        GetDataDtl(订单, cookieContainer, webHeaderCollection);
                    }
                    else if (i % oQ == 3)
                    {
                        if (orderTrNode.InnerText.Contains("发票号码"))
                        {
                            var 发票信息 = orderTrNode.SelectSingleNode("td/p").InnerText.Trim();
                            订单.发票信息 = new 发票信息 { 发票号码 = 发票信息 };
                            Res(发票信息);
                        }

                        订单集合.Add(订单);
                    }
                }

                Thread.Sleep(1000);

                pageIndex++;
                GetData(pageIndex, cookieContainer, webHeaderCollection);
            }
            catch (Exception ex)
            {
                txtRes.Text = ex.Message;
            }
        }

        public void GetDataDtl(订单 订单, CookieContainer cookieContainer, WebHeaderCollection webHeaderCollection)
        {
            try
            {
                var responseBody = HttpGet($"http://www.mrobao.com/main.php?m=product&s=admin_orderdetail&id={订单.订单编号}", cookieContainer, webHeaderCollection);
                var doc = new HtmlDocument();
                doc.LoadHtml(responseBody);

                //订单DtlNode
                var orderDtlNode = doc.DocumentNode.SelectSingleNode("//div[@class='order-detail']");

                var 收获信息 = orderDtlNode.SelectSingleNode("dl/dd[1]").InnerText.Trim();
                if (订单.收货信息 == null)
                    订单.收货信息 = new 收货信息();

                var 收货信息Arr = 收获信息.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                订单.收货信息.收件人 = 收货信息Arr[0];
                var 收获地址Arr = 收货信息Arr[收货信息Arr.Length - 1].Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                订单.收货信息.省 = 收获地址Arr[0];
                订单.收货信息.市 = 收获地址Arr[1];
                订单.收货信息.区 = 收获地址Arr[2];
                订单.收货信息.地址 = 收获地址Arr[3];
                订单.收货信息.联系电话 = new List<string>();
                for (int i = 1; i < (收货信息Arr.Length - 1); i++)
                {
                    订单.收货信息.联系电话.Add(收货信息Arr[i]);
                }

                var 发票信息 = orderDtlNode.SelectSingleNode("dl/dd[2]").InnerText.Trim();
                if (订单.发票信息 == null)
                    订单.发票信息 = new 发票信息();


            }
            catch (Exception ex)
            {
                txtRes.Text = ex.Message;
            }
        }

        /// <summary>
        /// GET请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="cookie"></param>
        /// <returns></returns>
        public static string HttpGet(string url, CookieContainer cookies, WebHeaderCollection headers)
        {
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);
            req.CookieContainer = cookies;
            req.Headers = headers;
            req.Method = "GET";
            req.ContentType = "text/html";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.115 Safari/537.36";
            req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9";
            req.Referer = "http://www.mrobao.com/main.php?m=product&s=admin_sellorder";

            WebResponse wr = req.GetResponse();
            Stream respStream = wr.GetResponseStream();
            StreamReader reader = new StreamReader(respStream, System.Text.Encoding.GetEncoding("utf-8"));
            string t = reader.ReadToEnd();
            wr.Close();
            return t;
        }
    }

    public class 订单
    {
        public string 订单编号 { get; set; }
        public string 下单时间 { get; set; }
        public string 付款时间 { get; set; }
        public string 发货时间 { get; set; }
        public string 物流名称 { get; set; }
        public string 物流单号 { get; set; }

        public 收货信息 收货信息 { get; set; }

        public 发票信息 发票信息 { get; set; }

        public 买家信息 买家信息 { get; set; }

        public List<订单商品> 订单商品 { get; set; }
    }

    public class 收货信息
    {
        public string 收件人 { get; set; }
        public List<string> 联系电话 { get; set; }
        public string 省 { get; set; }
        public string 市 { get; set; }
        public string 区 { get; set; }
        public string 地址 { get; set; }
    }

    public class 发票信息
    {
        public string 公司名称 { get; set; }
        public string 税号 { get; set; }
        public string 开户地址 { get; set; }
        public string 电话 { get; set; }
        public string 开户银行 { get; set; }
        public string 帐号 { get; set; }

        public string 发票号码 { get; set; }
        public string 金额 { get; set; }
        public string 开票时间 { get; set; }
    }

    public class 买家信息
    {
        public string 用户名 { get; set; }
        public string 昵称 { get; set; }
    }

    public class 订单商品
    {
        public string 商品名称 { get; set; }
        public string 状态 { get; set; }
        public string 单价_元 { get; set; }
        public string 数量 { get; set; }
        public string 商品总价_元 { get; set; }
    }
}