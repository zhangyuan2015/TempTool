using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
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

            button1.Enabled = false;
            Task.Run(() =>
            {
                try
                {
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

                    GetData(1, cookieContainer, webHeaderCollection);

                    Export(订单集合);
                }
                catch (Exception ex)
                {
                    Res($"异常：{ex.Message}");
                }
                finally
                {
                    button1.Enabled = true;
                }
            });
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
                Res("Cookie 错误，登录失败");
                return;
            }
            var doc = new HtmlDocument();
            doc.LoadHtml(responseBody);

            //总页数
            var pageNodes = doc.DocumentNode.SelectNodes("//div[@class='pagination']/a");
            if (pageNodes != null)
            {
                var pageNode = pageNodes[pageNodes.Count - 1/*倒数第二个-1*/- 1/*使用下标-1*/];
                if (totalPage == null)
                    totalPage = int.Parse(pageNode.InnerText.Trim('.'));
                PageCount(pageIndex, totalPage.Value);
                if (totalRows == null)
                {
                    var href = pageNode.GetAttributeValue("href", "");
                    totalRows = int.Parse(href.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault(a => a.Contains("totalRows")).SplitGetValue('='));
                }
            }
            else
            {
                totalPage = 1;
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
                    订单 = new 订单() { 买家信息 = new 买家信息(), 发票信息 = new 发票信息(), 收货信息 = new 收货信息(), 订单商品 = new List<订单商品>() };
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
                    //
                }
                else if (i % oQ == 3)
                {
                    if (orderTrNode.InnerText.Contains("发票号码"))
                    {
                        var 发票信息 = orderTrNode.SelectSingleNode("td/p").InnerText.Replace("\t", ",").Replace("\n", ",");
                        var 发票信息Arr = 发票信息.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var 发票信息item in 发票信息Arr)
                        {
                            if (发票信息item.Contains(nameof(订单.发票信息.发票号码)))
                                订单.发票信息.发票号码 = 发票信息item.SplitGetValue('：');
                            else if (发票信息item.Contains(nameof(订单.发票信息.金额)))
                                订单.发票信息.金额 = 发票信息item.SplitGetValue('：');
                            else if (发票信息item.Contains(nameof(订单.发票信息.开票时间)))
                                订单.发票信息.开票时间 = 发票信息item.SplitGetValue('：');
                        }
                    }

                    GetDataDtl(订单, cookieContainer, webHeaderCollection);
                    订单集合.Add(订单);
                }
            }

            Thread.Sleep(1000);

            pageIndex++;
            GetData(pageIndex, cookieContainer, webHeaderCollection);
        }

        public void GetDataDtl(订单 订单, CookieContainer cookieContainer, WebHeaderCollection webHeaderCollection)
        {
            try
            {
                string 订单详情Url = $"http://www.mrobao.com/main.php?m=product&s=admin_orderdetail&id={订单.订单编号}";
                订单.订单详情Url = 订单详情Url;

                var responseBody = HttpGet(订单详情Url, cookieContainer, webHeaderCollection);
                var doc = new HtmlDocument();
                doc.LoadHtml(responseBody);

                //订单DtlNode
                var orderDtlNode = doc.DocumentNode.SelectSingleNode("//div[@class='order-detail']");

                var orderDtlDLNodes = orderDtlNode.SelectNodes("dl");
                foreach (var orderDtlDLNode in orderDtlDLNodes)
                {
                    if (orderDtlDLNode.InnerText.Contains("收货地址："))
                    {
                        //收货信息
                        var 收获信息 = orderDtlDLNode.SelectSingleNode("dd[1]").InnerText.Trim();
                        var 收货信息Arr = 收获信息.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        订单.收货信息.收件人 = 收货信息Arr[0].FormatString();
                        var 收获地址Arr = 收货信息Arr[收货信息Arr.Length - 1].Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        订单.收货信息.省 = 收获地址Arr[0].FormatString();
                        订单.收货信息.市 = 收获地址Arr[1].FormatString();
                        订单.收货信息.区 = 收获地址Arr[2].FormatString();
                        订单.收货信息.地址 = 收获地址Arr[3].FormatString();
                        订单.收货信息.联系电话 = new List<string>();
                        for (int i = 1; i < (收货信息Arr.Length - 1); i++)
                        {
                            订单.收货信息.联系电话.Add(收货信息Arr[i].FormatString());
                        }

                        //发票信息
                        var 发票信息 = orderDtlDLNode.SelectSingleNode("dd[2]").InnerText.Trim();
                        var 发票信息Arr = 发票信息.Replace("\n", ",").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var 发票信息item in 发票信息Arr)
                        {
                            if (发票信息item.Contains(nameof(订单.发票信息.公司名称)))
                                订单.发票信息.公司名称 = 发票信息item.SplitGetValue('：');
                            else if (发票信息item.Contains(nameof(订单.发票信息.税号)))
                                订单.发票信息.税号 = 发票信息item.SplitGetValue('：');
                            else if (发票信息item.Contains("开户地址、电话"))
                                订单.发票信息.开户地址_电话 = 发票信息item.SplitGetValue('：');
                            else if (发票信息item.Contains(nameof(订单.发票信息.开户银行)))
                                订单.发票信息.开户银行 = 发票信息item.SplitGetValue('：');
                            else if (发票信息item.Contains(nameof(订单.发票信息.帐号)))
                                订单.发票信息.帐号 = 发票信息item.SplitGetValue('：');
                        }
                    }
                    else if (orderDtlDLNode.InnerText.Contains("买家信息"))
                    {
                        //买家信息
                        var 买家信息Arr = orderDtlDLNode.SelectNodes("dd").Select(a => a.InnerText);
                        foreach (var 买家信息item in 买家信息Arr)
                        {
                            if (买家信息item.Contains(nameof(订单.买家信息.用户名)))
                                订单.买家信息.用户名 = 买家信息item.SplitGetValue('：');
                            else if (买家信息item.Contains(nameof(订单.买家信息.昵称)))
                                订单.买家信息.昵称 = 买家信息item.SplitGetValue('：');
                        }
                    }
                    else if (orderDtlDLNode.InnerText.Contains("订单信息"))
                    {
                        //订单信息
                        var 订单信息Arr = orderDtlDLNode.SelectNodes("dd").Select(a => a.InnerText);
                        foreach (var 订单信息item in 订单信息Arr)
                        {
                            if (订单信息item.Contains(nameof(订单.付款时间)))
                                订单.付款时间 = 订单信息item.SplitGetValue('：');
                            else if (订单信息item.Contains(nameof(订单.发货时间)))
                                订单.发货时间 = 订单信息item.SplitGetValue('：');
                            else if (订单信息item.Contains(nameof(订单.物流名称)))
                                订单.物流名称 = 订单信息item.SplitGetValue('：');
                            else if (订单信息item.Contains(nameof(订单.物流单号)))
                                订单.物流单号 = 订单信息item.SplitGetValue('：');
                        }
                    }
                }

                var productNodes = orderDtlNode.SelectNodes("table/tr");
                productNodes.RemoveAt(0);
                foreach (var productNode in productNodes)
                {
                    订单商品 订单商品 = new 订单商品
                    {
                        商品图片 = productNode.SelectSingleNode("td[1]/div[2]/a").GetAttributeValue("href", ""),
                        商品名称 = productNode.SelectSingleNode("td[1]/div[2]/a").InnerText.FormatString(),
                        状态 = productNode.SelectSingleNode("td[2]").InnerText.FormatString(),
                        单价_元 = productNode.SelectSingleNode("td[3]").InnerText.FormatString(),
                        数量 = productNode.SelectSingleNode("td[4]").InnerText.FormatString()
                    };
                    订单.订单商品.Add(订单商品);
                }
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

        public void Export(List<订单> 订单集合)
        {
            Res($"共获取到 {订单集合.Count} 订单");

            if (!订单集合.Any())
                return;

            Res($"开始导出");

            string[] tableHeader =
            {
                "订单编号",
                "下单时间",
                "付款时间",
                "发货时间",
                "物流名称",
                "物流单号",
                "订单详情Url",
                "收件人",
                "联系电话",
                "省",
                "市",
                "区",
                "地址",
                "公司名称",
                "税号",
                "开户地址_电话",
                "开户银行",
                "帐号",
                "发票号码",
                "金额",
                "开票时间",
                "用户名",
                "昵称",
                "商品图片",
                "商品名称",
                "状态",
                "单价_元",
                "数量"
            };
            var tableHeaderList = tableHeader.ToList();

            var exportInfoList = new List<ExportInfo>();
            foreach (var 订单 in 订单集合)
            {
                foreach (var 商品 in 订单.订单商品)
                {
                    var exportInfo = new ExportInfo
                    {
                        订单编号 = 订单.订单编号,
                        下单时间 = 订单.下单时间,
                        付款时间 = 订单.付款时间,
                        发货时间 = 订单.发货时间,
                        物流名称 = 订单.物流名称,
                        物流单号 = 订单.物流单号,
                        订单详情Url = 订单.订单详情Url,

                        收件人 = 订单.收货信息.收件人,
                        联系电话 = String.Join(",", 订单.收货信息.联系电话),
                        省 = 订单.收货信息.省,
                        市 = 订单.收货信息.市,
                        区 = 订单.收货信息.区,
                        地址 = 订单.收货信息.地址,

                        公司名称 = 订单.发票信息.公司名称,
                        税号 = 订单.发票信息.税号,
                        开户地址_电话 = 订单.发票信息.开户地址_电话,
                        开户银行 = 订单.发票信息.开户银行,
                        帐号 = 订单.发票信息.帐号,
                        发票号码 = 订单.发票信息.发票号码,
                        金额 = 订单.发票信息.金额,
                        开票时间 = 订单.发票信息.开票时间,

                        用户名 = 订单.买家信息.用户名,
                        昵称 = 订单.买家信息.昵称,

                        商品图片 = 商品.商品图片,
                        商品名称 = 商品.商品名称,
                        状态 = 商品.状态,
                        单价_元 = 商品.单价_元,
                        数量 = 商品.数量
                    };

                    exportInfoList.Add(exportInfo);
                }
            }
            string path = Directory.GetCurrentDirectory();
            string name = DateTime.Now.ToString("yyyyMMddHHmmss.xlsx");
            ExportExcel($"{path}\\{name}", tableHeaderList, exportInfoList);

            Res($"导出完成：{path}/{name}");
        }

        public static void ExportExcel<T>(string filePath, List<string> tableHeader, IEnumerable<T> dataList)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("sheet1");
                //Font font = new Font("Calibri", 11, FontStyle.Bold);

                //设置表头
                for (int i = 1; i <= tableHeader.Count; i++)
                {
                    ws.Cells[1, i].Value = tableHeader[i - 1];
                    ws.Cells[1, i].AutoFitColumns(0);
                    //ws.Cells[1, i].Style.Font.SetFromFont(font);
                }
                //冻结首行
                ws.View.FreezePanes(2, 1);
                ws.Cells["A2"].LoadFromCollection(dataList);

                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    pck.SaveAs(fileStream);
                }
            }
        }
    }

    public static class StringUtils
    {
        public static string FormatString(this string str)
        {
            return (str ?? "").Replace("\n", "").Trim();
        }

        public static string SplitGetValue(this string str, char separator)
        {
            var sreArr = (str ?? "").Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
            if (sreArr.Length == 2)
                return sreArr[1].FormatString();
            return "";
        }
    }

    public class ExportInfo
    {
        public string 订单编号 { get; set; }
        public string 下单时间 { get; set; }
        public string 付款时间 { get; set; }
        public string 发货时间 { get; set; }
        public string 物流名称 { get; set; }
        public string 物流单号 { get; set; }
        public string 订单详情Url { get; set; }

        public string 收件人 { get; set; }
        public string 联系电话 { get; set; }
        public string 省 { get; set; }
        public string 市 { get; set; }
        public string 区 { get; set; }
        public string 地址 { get; set; }

        public string 公司名称 { get; set; }
        public string 税号 { get; set; }
        public string 开户地址_电话 { get; set; }
        public string 开户银行 { get; set; }
        public string 帐号 { get; set; }
        public string 发票号码 { get; set; }
        public string 金额 { get; set; }
        public string 开票时间 { get; set; }

        public string 用户名 { get; set; }
        public string 昵称 { get; set; }

        public string 商品图片 { get; set; }
        public string 商品名称 { get; set; }
        public string 状态 { get; set; }
        public string 单价_元 { get; set; }
        public string 数量 { get; set; }
    }

    public class 订单
    {
        public string 订单编号 { get; set; }
        public string 下单时间 { get; set; }
        public string 付款时间 { get; set; }
        public string 发货时间 { get; set; }
        public string 物流名称 { get; set; }
        public string 物流单号 { get; set; }

        public string 订单详情Url { get; set; }

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
        public string 开户地址_电话 { get; set; }
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
        public string 商品图片 { get; set; }
        public string 商品名称 { get; set; }
        public string 状态 { get; set; }
        public string 单价_元 { get; set; }
        public string 数量 { get; set; }
    }
}