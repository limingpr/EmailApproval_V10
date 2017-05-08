using OThinker.H3;
using OThinker.H3.Instance;
using OThinker.H3.Notification;
using OThinker.H3.WorkItem;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace OThinker.H3.Portal
{
    public class NotificationParser
    {
        private IEngine _Engine;
        protected IEngine Engine
        {
            get
            {
                return
                this._Engine;
            }
        }

        public NotificationParser(IEngine Engine)
        {
            this._Engine = Engine;
        }
        public NotificationParser()
        {
        }

        public const string Tag_ReceiveTime = "{ReceiveTime}";
        public const string Tag_RepeatItem_Start = "{RepeateItem}";
        public const string Tag_ApprovalSeparator = "*";
        public const string Tag_ActivityCode = "{ActivityCode}";
        public const string Tag_ItemFlag = "{ItemFlag}";
        public const string Tag_WorkItemCount = "{WorkItemCount}";
        public const string Tag_CurrentTime = "{CurrentTime}";
        public const string Tag_CurrentDate = "{CurrentDate}";
        public const string Tag_SeqNo = "{SeqNo}";
        public const string Tag_SheetUrl_End = "{/SheetUrl}";
        public const string Tag_SheetUrl_Start = "{SheetUrl}";
        public const string Tag_EmailApprove_End = "{/EmailApprove}";
        public const string Tag_ExceptionManager = "{ExceptionManager}";
        public const string Tag_ActivityName = "{ActivityName}";
        public const string Tag_EmailApprove_Start = "{EmailApprove}";
        public const string Tag_WorkflowCode = "{WorkflowCode}";
        public const string Tag_RepeatItem_End = "{/RepeateItem}";
        public const string Tag_DisplayName = "{DisplayName}";
        public const string Tag_Summary = "{Summary}";
        public const string Tag_WorkflowName = "{WorkflowName}";
        public const string Tag_ReceivedTime = "{ReceiveTime}";
        public const string Tag_Urger = "{Urger}";
        public const string Tag_WorkItemID = "{WorkItemID}";
        public const string Tag_Urgency = "{Urgency}";
        public const string Tag_Exception = "{Exception}";
        public const string Tag_InstanceId = "{InstanceId}";
        public const string Tag_InstanceName = "{InstanceName}";
        public const string Tag_ItemType = "{ItemType}";


        private string GetHtmlSnap(string Url)
        {
            System.Net.WebClient client = new System.Net.WebClient();
            client.Credentials = System.Net.CredentialCache.DefaultCredentials;
            var stream = client.OpenRead(Url);
            var reader = new System.IO.StreamReader(stream);
            var html = reader.ReadToEnd();
            reader.Close();
            return html;
        }
        public string GetStringByUrl(string strUrl)
        {
            WebRequest wrt = System.Net.WebRequest.Create(strUrl);
            WebResponse wrse = wrt.GetResponse();
            Stream strM = wrse.GetResponseStream();
            StreamReader SR = new StreamReader(strM, Encoding.GetEncoding("utf-8"));
            string strallstrm = SR.ReadToEnd();
            return strallstrm;
        }
        public void GetStringByWebUrl(string strUrl)
        {

            WebBrowser web = new WebBrowser();
            web.Navigate(strUrl);
            web.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(web_DocumentCompleted);

        }
        void web_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser web = (WebBrowser)sender;
            Thread.Sleep(10);
            var content = web.Document.Body.InnerHtml;

            content = ParseSheet(content, web.Url.AbsoluteUri.ToString());
            OThinker.H3.Portal.MessageService s = new OThinker.H3.Portal.MessageService();
            s.SendEmail("lim@authine.com", "Web消息测试", content);


        }
        private string ParseMvcSheet(OThinker.H3.WorkItem.WorkItem WorkItem, string Content)
        {
            string result = Content;
            while (true)
            {
                if (result == null)
                {
                    break;
                }
                int sheetStart = result.IndexOf(NotificationParser.Tag_SheetUrl_Start);
                int sheetEnd = result.IndexOf(NotificationParser.Tag_SheetUrl_End);
                if (sheetStart == -1 || sheetEnd == -1 || sheetEnd < sheetStart)
                {
                    break;
                }
                string url = result.Substring(sheetStart + NotificationParser.Tag_SheetUrl_Start.Length, sheetEnd - (sheetStart + NotificationParser.Tag_SheetUrl_Start.Length));
                string html = null;
                try
                {

                    GetStringByWebUrl(url);

                    html = GetHtmlSnap(url);
                    html = GetStringByUrl(url);


                    string httpHeader = url.Substring(0, url.LastIndexOf("/") + 1);
                    var builder = new StringBuilder();
                    if (httpHeader != string.Empty)
                    {
                        var siteName = httpHeader.Substring(httpHeader.Substring(0, httpHeader.Length - 1).LastIndexOf("/"));
                        string siteHeader = httpHeader.Replace(siteName, string.Empty);
                        var hrefs = html.Split(new string[] { "href=\"" }, StringSplitOptions.None);
                        builder.Append(hrefs[0]);
                        for (int i = 1; i < hrefs.Length; i++)
                        {
                            if (hrefs[i].ToLower().StartsWith("http"))
                            {
                                builder.Append("href=\"" + hrefs[i]);
                            }
                            else
                            {
                                if (hrefs[i].StartsWith(siteName))
                                {
                                    builder.Append("href=\"" + siteHeader + hrefs[i]);

                                }
                                else
                                {
                                    builder.Append("href=\"" + httpHeader + hrefs[i]);

                                }
                            }
                        }
                        html = builder.ToString();


                        builder = new StringBuilder();

                        hrefs = html.Split(new string[] { "src=\"" }, StringSplitOptions.None);
                        builder.Append(hrefs[0]);
                        for (int i = 1; i < hrefs.Length; i++)
                        {
                            if (hrefs[i].ToLower().StartsWith("http"))
                            {
                                builder.Append("src=\"" + hrefs[i]);
                            }
                            else
                            {
                                if (hrefs[i].StartsWith(siteName))
                                {
                                    builder.Append("src=\"" + siteHeader + hrefs[i]);

                                }
                                else
                                {
                                    builder.Append("src=\"" + httpHeader + hrefs[i]);

                                }
                            }
                        }


                    }

                    html = builder.ToString();
                }
                catch
                {

                }
                result = result.Substring(0, sheetStart) + html + result.Substring(sheetEnd + NotificationParser.Tag_SheetUrl_End.Length);


            }


            return result;

        }

        private string ParseSheet(OThinker.H3.WorkItem.WorkItem WorkItem, string Content)
        {
            string result = Content;
            while (true)
            {
                if (result == null)
                {
                    break;
                }
                int sheetStart = result.IndexOf(NotificationParser.Tag_SheetUrl_Start);
                int sheetEnd = result.IndexOf(NotificationParser.Tag_SheetUrl_End);
                if (sheetStart == -1 || sheetEnd == -1 || sheetEnd < sheetStart)
                {
                    break;
                }
                string url = result.Substring(sheetStart + NotificationParser.Tag_SheetUrl_Start.Length, sheetEnd - (sheetStart + NotificationParser.Tag_SheetUrl_Start.Length));
                string html = null;
                try
                {

                    GetStringByWebUrl(url);

                    html = GetHtmlSnap(url);
                    html = GetStringByUrl(url);


                    string httpHeader = url.Substring(0, url.LastIndexOf("/") + 1);
                    var builder = new StringBuilder();
                    if (httpHeader != string.Empty)
                    {
                        var siteName = httpHeader.Substring(httpHeader.Substring(0, httpHeader.Length - 1).LastIndexOf("/"));
                        string siteHeader = httpHeader.Replace(siteName, string.Empty);
                        var hrefs = html.Split(new string[] { "href=\"" }, StringSplitOptions.None);
                        builder.Append(hrefs[0]);
                        for (int i = 1; i < hrefs.Length; i++)
                        {
                            if (hrefs[i].ToLower().StartsWith("http"))
                            {
                                builder.Append("href=\"" + hrefs[i]);
                            }
                            else
                            {
                                if (hrefs[i].StartsWith(siteName))
                                {
                                    builder.Append("href=\"" + siteHeader + hrefs[i]);

                                }
                                else
                                {
                                    builder.Append("href=\"" + httpHeader + hrefs[i]);

                                }
                            }
                        }
                        html = builder.ToString();


                        builder = new StringBuilder();

                        hrefs = html.Split(new string[] { "src=\"" }, StringSplitOptions.None);
                        builder.Append(hrefs[0]);
                        for (int i = 1; i < hrefs.Length; i++)
                        {
                            if (hrefs[i].ToLower().StartsWith("http"))
                            {
                                builder.Append("src=\"" + hrefs[i]);
                            }
                            else
                            {
                                if (hrefs[i].StartsWith(siteName))
                                {
                                    builder.Append("src=\"" + siteHeader + hrefs[i]);

                                }
                                else
                                {
                                    builder.Append("src=\"" + httpHeader + hrefs[i]);

                                }
                            }
                        }


                    }

                    html = builder.ToString();
                }
                catch
                {

                }
                result = result.Substring(0, sheetStart) + html + result.Substring(sheetEnd + NotificationParser.Tag_SheetUrl_End.Length);
            }


            return result;

        }
        public string ParseSheet(string html, string url)
        {
            string result = "";

            try
            {



                string httpHeader = url.Substring(0, url.LastIndexOf("/") + 1);
                var builder = new StringBuilder();
                if (httpHeader != string.Empty)
                {
                    var siteName = httpHeader.Substring(httpHeader.Substring(0, httpHeader.Length - 1).LastIndexOf("/"));
                    string siteHeader = httpHeader.Replace(siteName, string.Empty);
                    var hrefs = html.Split(new string[] { "href=\"" }, StringSplitOptions.None);
                    builder.Append(hrefs[0]);
                    for (int i = 1; i < hrefs.Length; i++)
                    {
                        if (hrefs[i].ToLower().StartsWith("http"))
                        {
                            builder.Append("href=\"" + hrefs[i]);
                        }
                        else
                        {
                            if (hrefs[i].StartsWith(siteName))
                            {
                                builder.Append("href=\"" + siteHeader + hrefs[i]);

                            }
                            else
                            {
                                builder.Append("href=\"" + httpHeader + hrefs[i]);

                            }
                        }
                    }
                    html = builder.ToString();


                    builder = new StringBuilder();

                    hrefs = html.Split(new string[] { "src=\"" }, StringSplitOptions.None);
                    builder.Append(hrefs[0]);
                    for (int i = 1; i < hrefs.Length; i++)
                    {
                        if (hrefs[i].ToLower().StartsWith("http"))
                        {
                            builder.Append("src=\"" + hrefs[i]);
                        }
                        else
                        {
                            if (hrefs[i].StartsWith(siteName))
                            {
                                builder.Append("src=\"" + siteHeader + hrefs[i]);

                            }
                            else
                            {
                                builder.Append("src=\"" + httpHeader + hrefs[i]);

                            }
                        }
                    }


                }

                html = builder.ToString();

            }
            catch
            {

            }
            result = html;

            return result;

        }


        public string ParseContent(OThinker.H3.WorkItem.WorkItem WorkItem, string Content)
        {
            string c = Content;
            if (c == null || WorkItem == null)
            {
                return null;
            }
            c = c.Replace(Tag_WorkItemID, WorkItem.WorkItemID);
            c = c.Replace(Tag_ReceiveTime, WorkItem.ReceiveTime.ToShortDateString());
            c = c.Replace(Tag_DisplayName, WorkItem.DisplayName);
            c = c.Replace(Tag_Summary, WorkItem.ItemSummary);
            c = c.Replace(Tag_InstanceId, WorkItem.InstanceId);
            c = c.Replace(Tag_WorkflowCode, WorkItem.WorkflowCode);
            c = c.Replace(Tag_ActivityName, WorkItem.DisplayName);

            var Context = Engine.InstanceManager.GetInstanceContext(WorkItem.InstanceId);
            c = c.Replace(Tag_InstanceName, Context.InstanceName);
            c = c.Replace(Tag_SeqNo, Context.SequenceNo);
            c = ParseSheet(WorkItem, c);
            if (c.IndexOf("{") > -1 && c.IndexOf("}") > -1)
            {
                var instanceData = new OThinker.H3.Instance.InstanceData(Engine, WorkItem.InstanceId, WorkItem.TokenId, WorkItem.ActivityCode, WorkItem.Participant)
                    ;
                c = instanceData.ParseText(c);

            }
            return c;
        }

        public string ParseContent(OThinker.H3.Notification.Notification Notification, string Content)
        {
            string c = Content;
            if (c == null || Notification == null)
            {
                return null;
            }
            c = c.Replace(Tag_InstanceId, Notification.InstanceId);

            var Context = Engine.InstanceManager.GetInstanceContext(Notification.InstanceId);
            c = c.Replace(Tag_InstanceName, Context.InstanceName);
            c = c.Replace(Tag_SeqNo, Context.SequenceNo);
            //c = ParseSheet(Notification.Url, c);

            //c = ParseSheet(Notification, c);
            if (c.IndexOf("{") > -1 && c.IndexOf("}") > -1)
            {
                var instanceData = new InstanceData(
    Engine,
     Notification.InstanceId,
   null);
                c = instanceData.ParseText(c);

            }

            return c;

        }






    }


}
