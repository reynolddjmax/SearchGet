using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Web;
using System.IO;

namespace SearchGet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {



            int n = 1;
            foreach (var item in richTextBox1.Lines)
            {

                string a1 = "";
                string a2 = "";
                string a3 = "";
                string a4 = "";

                if (checkBox1.Checked)
                {
                    a1 = AskBaiDu(item);
                    this.label2.Text = item + " : 验证百度";
                    Application.DoEvents();
                }

                if (checkBox2.Checked)
                {
                    a2 = AskSoGou(item);
                    this.label2.Text = item + " : 验证搜狗";
                    Application.DoEvents();
                }

                if (checkBox3.Checked)
                {
                    a3 = AskSo(item);
                    this.label2.Text = item + " : 验证360";
                    Application.DoEvents();
                }

                if (checkBox4.Checked)
                {
                    a4 = AskYz(item);
                    this.label2.Text = item + " : 验证神马";
                    Application.DoEvents();
                }









                dataGridView1.Rows.Add(item,a1 , a2 , a3 , a4);


                this.label1.Text = n.ToString();
                n++;
                Application.DoEvents();

            }

            this.label1.Text = "完成";

        }


        string AskBaiDu(string Key)
        {
            try
            {
                string url = @"http://www.baidu.com/";
                string encodedKeyword = HttpUtility.UrlEncode(Key, Encoding.GetEncoding(936));
                //百度使用codepage 936字符编码来作为查询串，果然专注于中文搜索……
                //更不用说，还很喜欢微软
                //谷歌能正确识别UTF-8编码和codepage这两种情况，不过本身网页在HTTP头里标明是UTF-8的
                //估计谷歌也不讨厌微软（以及微软的专有规范）
                string query = "s?wd=" + encodedKeyword;

                HttpWebRequest req;
                HttpWebResponse response;
                Stream stream;
                req = (HttpWebRequest)WebRequest.Create(url + query);
                response = (HttpWebResponse)req.GetResponse();
                stream = response.GetResponseStream();
                int count = 0;
                byte[] buf = new byte[8192];
                string decodedString = null;
                StringBuilder sb = new StringBuilder();

                Console.WriteLine("正在读取网页{0}的内容……", url + query);
                do
                {
                    count = stream.Read(buf, 0, buf.Length);
                    if (count > 0)
                    {
                        decodedString = Encoding.GetEncoding("utf-8").GetString(buf, 0, count);
                        sb.Append(decodedString);
                    }
                } while (count > 0);




                string aa = sb.ToString();

                if (aa.IndexOf("很抱歉，没有找到与") != -1 || aa.IndexOf("没有找到该URL。您可以直接访问") != -1)
                {
                    return "无结果";
                }







                if (checkBox5.Checked)
                {
                    this.richTextBox2.Text = aa;
                }

                return "有结果" + aa.Length.ToString();





            }
            catch (Exception ex)
            {
                
                return  ex.Message;
            }

        }


        string AskSoGou(string Key)
        {
            try
            {
                string url = @"http://www.sogou.com/";
                string encodedKeyword = HttpUtility.UrlEncode(Key, Encoding.GetEncoding(936));
                //百度使用codepage 936字符编码来作为查询串，果然专注于中文搜索……
                //更不用说，还很喜欢微软
                //谷歌能正确识别UTF-8编码和codepage这两种情况，不过本身网页在HTTP头里标明是UTF-8的
                //估计谷歌也不讨厌微软（以及微软的专有规范）
                string query = "web?query=" + encodedKeyword;

                HttpWebRequest req;
                HttpWebResponse response;
                Stream stream;
                req = (HttpWebRequest)WebRequest.Create(url + query);
                response = (HttpWebResponse)req.GetResponse();
                stream = response.GetResponseStream();
                int count = 0;
                byte[] buf = new byte[8192];
                string decodedString = null;
                StringBuilder sb = new StringBuilder();

                Console.WriteLine("正在读取网页{0}的内容……", url + query);
                do
                {
                    count = stream.Read(buf, 0, buf.Length);
                    if (count > 0)
                    {
                        decodedString = Encoding.GetEncoding("utf-8").GetString(buf, 0, count);
                        sb.Append(decodedString);
                    }
                } while (count > 0);




                string aa = sb.ToString();

                if (aa.IndexOf("搜狗已为您找到约0条相关结果") != -1 && aa.IndexOf("未收录") != -1)
                {
                    return "无结果";
                }




                if (checkBox5.Checked)
                {
                    this.richTextBox2.Text = aa;
                }

                return "有结果" + aa.Length.ToString();
            }
            catch (Exception ex)
            {

                return ex.Message;
            }



        }

        string AskSo(string Key)
        {
            try
            {
                string url = @"http://www.so.com/";
                string encodedKeyword = HttpUtility.UrlEncode(Key, Encoding.GetEncoding(936));
                //百度使用codepage 936字符编码来作为查询串，果然专注于中文搜索……
                //更不用说，还很喜欢微软
                //谷歌能正确识别UTF-8编码和codepage这两种情况，不过本身网页在HTTP头里标明是UTF-8的
                //估计谷歌也不讨厌微软（以及微软的专有规范）
                string query = "s?q=" + encodedKeyword;

                HttpWebRequest req;
                HttpWebResponse response;
                Stream stream;
                req = (HttpWebRequest)WebRequest.Create(url + query);
                response = (HttpWebResponse)req.GetResponse();
                stream = response.GetResponseStream();
                int count = 0;
                byte[] buf = new byte[8192];
                string decodedString = null;
                StringBuilder sb = new StringBuilder();

                Console.WriteLine("正在读取网页{0}的内容……", url + query);
                do
                {
                    count = stream.Read(buf, 0, buf.Length);
                    if (count > 0)
                    {
                        decodedString = Encoding.GetEncoding("utf-8").GetString(buf, 0, count);
                        sb.Append(decodedString);
                    }
                } while (count > 0);


                
                string aa = sb.ToString();

                if (aa.IndexOf("找不到该URL，可以直接访问") != -1 || aa.IndexOf("检查输入是否正确") != -1)
                {
                    return "无结果";
                }


                if (checkBox5.Checked)
                {
                    this.richTextBox2.Text = aa;
                }


                return "有结果" + aa.Length.ToString();
            }
            catch (Exception ex)
            {

                return ex.Message;
            }



        }
        string AskYz(string Key)
        {
            try
            {
                string url = @"http://m.sm.cn/";
                string encodedKeyword = HttpUtility.UrlEncode(Key, Encoding.GetEncoding(936));
                //百度使用codepage 936字符编码来作为查询串，果然专注于中文搜索……
                //更不用说，还很喜欢微软
                //谷歌能正确识别UTF-8编码和codepage这两种情况，不过本身网页在HTTP头里标明是UTF-8的
                //估计谷歌也不讨厌微软（以及微软的专有规范）
                string query = "s?q=" + encodedKeyword;

                HttpWebRequest req;
                HttpWebResponse response;
                Stream stream;
                req = (HttpWebRequest)WebRequest.Create(url + query);

                req.UserAgent = "Mozilla/5.0 (iPhone 6s; CPU iPhone OS 10_3_1 like Mac OS X) AppleWebKit/603.1.30 (KHTML, like Gecko) Version/10.0";


                response = (HttpWebResponse)req.GetResponse();
                stream = response.GetResponseStream();
                int count = 0;
                byte[] buf = new byte[8192];
                string decodedString = null;
                StringBuilder sb = new StringBuilder();

                Console.WriteLine("正在读取网页{0}的内容……", url + query);
                do
                {
                    count = stream.Read(buf, 0, buf.Length);
                    if (count > 0)
                    {
                        decodedString = Encoding.GetEncoding("utf-8").GetString(buf, 0, count);
                        sb.Append(decodedString);
                    }
                } while (count > 0);




                string aa = sb.ToString();


                if (aa.IndexOf("抱歉") != -1 && aa.IndexOf("没有找到与") != -1 && aa.IndexOf("请检查输入文字是否有误") != -1)
                {
                    return "无结果";
                }



                if (checkBox5.Checked)
                {
                    this.richTextBox2.Text = aa;
                }


                return "有结果" + aa.Length.ToString();
            }
            catch (Exception ex)
            {

                return ex.Message;
            }




        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Rows.Clear();
        }
    }
}
