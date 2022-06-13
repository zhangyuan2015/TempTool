using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Windows.Forms;
using HtmlAgilityPack;
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
            try
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
                var responseBody = HttpGet($"http://www.mrobao.com/main.php?m=product&s=admin_sellorder&key={txtKey.Text.Trim()}&buy_catid=&is_invoice=", cookieContainer, webHeaderCollection);

                var doc = new HtmlDocument();
                doc.LoadHtml(responseBody);
                var orderNode = doc.DocumentNode.SelectSingleNode("//table[@class='table-list-style order']");
                var orderTrNodes = orderNode.SelectNodes("//tbody/tr");

                List<订单> 订单集合 = new List<订单>();

                int q = 4;
                订单 订单 = null;
                for (int i = 0; i < orderTrNodes.Count; i++)
                {

                    if (i % q == 0)
                    {
                        订单 = new 订单();
                        //var a = orderTrNodes[i].InnerText;
                    }
                    else if (i % q == 1)
                    {
                        var 订单编号 = orderTrNodes[i].SelectSingleNode("//th/span[1]/span").InnerText.Trim();
                        var 下单时间 = orderTrNodes[i].SelectSingleNode("//th/span[2]/span").InnerText.Trim();

                        订单.订单编号 = 订单编号;
                        订单.下单时间 = 下单时间;
                    }
                    else if (i % q == 2)
                    {
                        var a = orderTrNodes[i].InnerText;
                    }
                    else if (i % q == 3)
                    {
                        if (orderTrNodes[i].InnerText.Contains("发票号码"))
                        {
                            var a = orderTrNodes[i].SelectSingleNode("//td");
                            var 发票信息 = orderTrNodes[i].SelectSingleNode("//td/p").InnerText.Trim();
                            订单.发票信息 = new 发票信息 { 发票号码 = 发票信息 };
                        }

                        订单集合.Add(订单);
                    }
                }

                var totalPage = doc.DocumentNode.SelectSingleNode("//div[@class='pagination']/a[10]").InnerText.Trim('.');
                txtPageCount.Text = $"1 / {totalPage}";

                //string Title = headNode.InnerText;
                txtRes.Text = responseBody;
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

        public 收货地址 收货地址 { get; set; }

        public 发票信息 发票信息 { get; set; }

        public 买家信息 买家信息 { get; set; }

        public List<订单商品> 订单商品 { get; set; }
    }

    public class 收货地址
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
        public string 商品 { get; set; }
        public string 状态 { get; set; }
        public string 单价_元 { get; set; }
        public string 数量 { get; set; }
        public string 商品总价_元 { get; set; }
    }
}