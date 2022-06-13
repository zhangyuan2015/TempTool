using System;
using System.IO;
using System.Net;
using System.Windows.Forms;

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
}