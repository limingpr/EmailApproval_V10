using OThinker.H3;
using OThinker.H3.Portal;
using OThinker.Organization;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Windows.Forms;
namespace EmailApproval_V10
{
    public partial class Form1 : Form
    {
        //邮件发送地址
        public string Address = "";

        //邮件标题
        public string Titles = "";
        //邮件发送内容
        public string HyLink = "";

        //邮件审批内容
        public string ApproveContent = "";

        //表单页面
        public string SheetUrl = "";
        //开始时间
        public DateTime TimeSpan = System.DateTime.Now;

        public int NowIndex = 0;

        public DataTable dt = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            timer2.Enabled = false;
            timer3.Enabled = false;
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtWorkTime.Text))
            {
                MessageBox.Show("轮训待办任务间隔时间不能为空！");

            }
            if (string.IsNullOrEmpty(txtFormTime.Text))
            {
                MessageBox.Show("表单加载时间不能为空！");

            }

            timer1.Interval = int.Parse(txtWorkTime.Text) * 1000;
            timer2.Interval = int.Parse(txtFormTime.Text) * 1000;
            timer1.Enabled = true;
            TimeSpan = System.DateTime.Now.AddMinutes(-int.Parse(txtLastTime.Text));
        }

        /// <summary>
        /// 定时读取WorkItem表数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            NowIndex = 0;
            var sqlworkitem = string.Format(@"SELECT TOP 100
	*
FROM ot_workitem t
WHERE t.receiveTime > '{0}'
ORDER BY t.ReceiveTime DESC", TimeSpan);

            dt = OThinker.H3.Controllers.AppUtility.Engine.Query.QueryTable(sqlworkitem);
            if (dt.Rows.Count > 0)
            {
                timer3.Enabled = true;

            }
            else
            {
                timer3.Enabled = false;

            }
        }


        public void OnUpdated(IEngine Engine, OThinker.H3.WorkItem.WorkItem WorkItem)
        {

            #region 邮件发送
            //获取参与者信息
            var user = Engine.Organization.GetUnit(WorkItem.Participant) as User;

            var instance = Engine.InstanceManager.GetInstanceContext(WorkItem.InstanceId);

            Address = user.Email;
            Titles = string.Format("流程：{0} {1} ", instance.InstanceName, WorkItem.DisplayName);
            Engine.LogWriter.Write("\n邮件发送标题:" + user.Email + Titles + ",");

            var Content = Engine.SettingManager.GetCustomSetting(OThinker.H3.Settings.CustomSetting.Setting_EmailNotificationContent);
            Content = Content.Replace("{InstanceName}", instance.InstanceName).Replace("{WorkItemID}", WorkItem.WorkItemID).Replace("{InstanceId}", instance.InstanceId);

            int sheetStart = Content.IndexOf(OThinker.H3.Portal.NotificationParser.Tag_SheetUrl_Start);
            int sheetEnd = Content.IndexOf(OThinker.H3.Portal.NotificationParser.Tag_SheetUrl_End);
            if (sheetStart == -1 || sheetEnd == -1 || sheetEnd < sheetStart)
            {

            }
            else
            {
                SheetUrl = Content.Substring(sheetStart + OThinker.H3.Portal.NotificationParser.Tag_SheetUrl_Start.Length, sheetEnd - (sheetStart + OThinker.H3.Portal.NotificationParser.Tag_SheetUrl_Start.Length));
                webBrowser1.Navigate(SheetUrl);

                while (webBrowser1.ReadyState != WebBrowserReadyState.Complete) Application.DoEvents();

                timer2.Enabled = false;

            }
            HyLink = Content.Substring(0, sheetStart);
            ApproveContent = Content.Substring(sheetEnd + NotificationParser.Tag_SheetUrl_End.Length);

            #endregion

        }


        /// <summary>
        /// 读取表单信息后定时发送
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            SendMail(true);
            timer3.Enabled = true;
        }

        /// <summary>
        /// 轮询已查询出的待办任务DataTable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer3_Tick(object sender, EventArgs e)
        {
            timer3.Enabled = false;
            if (NowIndex == dt.Rows.Count)
            {
                timer1.Enabled = true;

            }
            else
            {
                var workitemid = dt.Rows[NowIndex]["ObjectID"].ToString();
                var workitem = OThinker.H3.Controllers.AppUtility.Engine.WorkItemManager.GetWorkItem(workitemid);

                OnUpdated(OThinker.H3.Controllers.AppUtility.Engine, workitem);
                NowIndex++;

            }



        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var workitemid = textBox1.Text;
            var workitem = OThinker.H3.Controllers.AppUtility.Engine.WorkItemManager.GetWorkItem(workitemid);
            if (workitem != null)
            {
                OnUpdated(OThinker.H3.Controllers.AppUtility.Engine, workitem);

            }

        }

        /// <summary>
        /// 邮件发送内容
        /// </summary>
        private void SendMail(bool IsRefresh)
        {
            if (webBrowser1.Document != null)
            {
                var content = webBrowser1.Document.Body.InnerHtml;
                var table = FormatTable(content);

                //var htmlcontent = webBrowser1.DocumentText;
                //var headstart = htmlcontent.IndexOf("<head>");
                //var headend = htmlcontent.IndexOf("</head>");
                //var headcontent = htmlcontent.Substring(headstart + "<head>".Length, headend - (headstart + "<head>".Length));
                //var scripts = content.Split(new string[] { "<script>var init = [];</script>" }, StringSplitOptions.None);
                //content = scripts[0] + "<script>var init = [];</script>" + headcontent + scripts[1];
                //content = ParseSheet(content, SheetUrl);
                //content.Replace("$.MvcSheet.Init()", "");
                //var toolbar = content.Split(new string[] { "id=\"main-navbar\"" }, StringSplitOptions.None);
                //content = toolbar[0] + "id=\"main-navbar\"" + "   style=\"visibility:hidden\"  " + toolbar[1];
                OThinker.H3.Portal.MessageService s = new OThinker.H3.Portal.MessageService();
                System.IO.StringWriter sw = new System.IO.StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                table.RenderControl(htw);

                var mailcontent = GetHtmlText("SheetMail.html");
                mailcontent = mailcontent.Replace("{MailContent}", sw.ToString());
                s.SendEmail("lim@authine.com", Titles, HyLink + mailcontent + ApproveContent);
                //s.SendEmail(Address, Titles, HyLink + content+ApproveContent);
                //邮件发送后清空webBrowser，避免重复发送
                if (IsRefresh)
                {
                    webBrowser1.Navigate("about:blank");

                }


            }

        }

        /// <summary>
        /// 将div表单内容转换为Table
        /// </summary>
        /// <param name="divContent"></param>
        /// <returns></returns>
        private HtmlTable FormatTable(string divContent)
        {
            //获取第一层的div标签
            Regex reg = new Regex(@"(?is)<div[^>]*>(?><div[^>]*>(?<o>)|</div>(?<-o>)|(?:(?!</?div\b).)*)*(?(o)(?!))</div>");
            //获取 class维row 、row tableContent、bannerTitle的div内容
            Regex reg3 = new Regex(@"(?is)<div (class=" + "\"row\"|class=" + "\"nav-icon fa fa-chevron-right bannerTitle\"|class=" + "\"row tableContent\"" + @")[^>]*>(?><div[^>]*>(?<o>)|</div>(?<-o>)|(?:(?!</?div\b).)*)*(?(o)(?!))</div>");
            MatchCollection mc = reg.Matches(divContent);
            MatchCollection list2 = reg3.Matches(mc[2].Value);

            var table = new HtmlTable();
            table.ID = "table1";
            table.Attributes.Add("class", "table");

            var firstrow = new HtmlTableRow();
            var firstCell = new HtmlTableCell();
            firstCell.Attributes.Add("class", "processTitle");
            firstCell.ColSpan = 12;
            firstCell.InnerText = webBrowser1.DocumentTitle;
            firstCell.Style.Add("min-height", "35px");
            firstrow.Cells.Add(firstCell);
            table.Rows.Add(firstrow);

            foreach (Match item in list2)
            {
                var row = new HtmlTableRow();

                var cells = reg.Matches(item.Value.Substring(4, item.Value.Length - 10));
                if (cells.Count > 0)
                {
                    foreach (Match itemcell in cells)
                    {
                        var cell = new HtmlTableCell();
                        var classname = GetTitleContent(itemcell.Value, "div", "class");
                        var colspan = classname.Replace("col-md-", "");
                        cell.ColSpan = int.Parse(colspan);
                        cell.InnerHtml = itemcell.Value;
                        if (colspan == "2")
                        {
                            cell.Width = "15%";

                        }


                        row.Cells.Add(cell);

                    }
                }
                else
                {
                    var cell = new HtmlTableCell();
                    var classname = GetTitleContent(item.Value, "div", "class");
                    if (item.Value.Contains("bannerTitle"))
                    {
                        cell.ColSpan = 12;
                        cell.InnerHtml = item.Value;

                    }
                    row.Cells.Add(cell);


                }
                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>  
        /// 获取字符中指定标签的值  
        /// </summary>  
        /// <param name="str">字符串</param>  
        /// <param name="title">标签</param>  
        /// <param name="attrib">属性名</param>  
        /// <returns>属性</returns>  
        public static string GetTitleContent(string str, string title, string attrib)
        {

            string tmpStr = string.Format("<{0}[^>]*?{1}=(['\"\"]?)(?<url>[^'\"\"\\s>]+)\\1[^>]*>", title, attrib); //获取<title>之间内容  

            Match TitleMatch = Regex.Match(str, tmpStr, RegexOptions.IgnoreCase);

            string result = TitleMatch.Groups["url"].Value;
            return result;
        }

        /// <summary>
        /// 读取HTML模板内容
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public string GetHtmlText(string filename)
        {
            string stmp = Assembly.GetExecutingAssembly().Location;
            stmp = stmp.Substring(0, stmp.LastIndexOf('\\'));//删除文件名
            var path = stmp.Replace(@"\bin\Debug", "") + @"\Template\" + filename;
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Read);
            StreamReader sr = new StreamReader(fs, Encoding.Default);
            var content = sr.ReadToEnd();
            sr.Close();
            fs.Close();
            return content;


        }

        /// <summary>
        /// 测试发送
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTest_Click(object sender, EventArgs e)
        {
            SendMail(false);

        }
    }
}
