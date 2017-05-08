using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OThinker.Clusterware;
using OThinker.Data.Database;
using System.Configuration;
using System.Net.Mail;

namespace OThinker.H3.Portal
{
    public class MessageService
    {
        public MessageService() { }

        private int _Port = 0;
        /// <summary>
        /// 获取邮件发送端口号
        /// </summary>
        private int Port
        {
            get
            {
                return this._Port;
            }
        }

        /// <summary>
        /// 发送电子邮件
        /// </summary>
        /// <param name="Address">邮件地址</param>
        /// <param name="Subject">邮件标题</param>
        /// <param name="Body">邮件内容</param>
        public void SendEmail(string Address, string Subject, string Body)
        {
            //ConnectionSetting setting = new OThinker.Clusterware.ConnectionSetting()
            //{
            //    Address = "127.0.0.1",
            //    ObjectUri = OThinker.H3.Configs.ProductInfo.EngineUri,
            //    Port = Port
            //};
            //OThinker.Clusterware.MasterConnection connection = new Clusterware.MasterConnection(
            //            setting,
            //            UserName,
            //            Password);


            string smtp = "smtp.wewowhealth.com";
            if (!string.IsNullOrEmpty(smtp))
            {
                string from ="OA";
                string userName = "oa@wewowhealth.com";// from;// @"coli\coli_workflow";
                string password = "wewowoa2016";

                try
                {
                    // 发送该邮件
                    System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(smtp);
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(userName, password);
                    client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    client.EnableSsl = false;
                    client.Port = 25;// VesselCustomSetting.GetValue<int>(this.GetSettingValue(VesselCustomSetting.Setting_SmtpPort), 25);
                    Encoding subjectEncoding = null;
                    // 默认值是UTF8
                    subjectEncoding = System.Text.Encoding.UTF8;

                    Encoding bodyEncoding = subjectEncoding;

                    this.SendMailBySmtp(client,
                        userName,
                        from,
                        subjectEncoding,
                        bodyEncoding,
                        Address,
                        Subject,
                        Body);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="SmtpClient">SMTP对象</param>
        /// <param name="UserName">发送的邮箱账号</param>
        /// <param name="From">发送的邮箱显示名称</param>
        /// <param name="SubjectEncoding">编码</param>
        /// <param name="BodyEncoding">编码</param>
        /// <param name="Address">接收邮箱的地址</param>
        /// <param name="Title">邮件标题</param>
        /// <param name="Content">邮件内容</param>
        private void SendMailBySmtp(System.Net.Mail.SmtpClient SmtpClient,
            string UserName,
            string From,
            Encoding SubjectEncoding,
            Encoding BodyEncoding,
            string Address,
            string Title,
            string Content)
        {
            if (string.IsNullOrEmpty(Address))
            {
                return;
            }
            MailAddress fromAddress = new MailAddress(UserName, From, SubjectEncoding);
            MailAddress toAddress = new MailAddress(Address, Address, SubjectEncoding);
            MailMessage message = new MailMessage(fromAddress, toAddress)
            {
                Subject = Title,
                Body = Content,
                SubjectEncoding = SubjectEncoding,
                BodyEncoding = BodyEncoding,
                IsBodyHtml = true
            };
            //System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(
            //    From,
            //    Address,
            //    Title,
            //    Content);
            // 编码
            //message.SubjectEncoding = SubjectEncoding;
            //message.BodyEncoding = BodyEncoding;
            //message.IsBodyHtml = true;

            SmtpClient.Send(message);
        }

    }
}
