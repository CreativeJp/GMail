using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GMailSync.Helper
{
    #region Usings

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web;
    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Gmail.v1;
    using Google.Apis.Gmail.v1.Data;
    using System.IO;
    using Google.Apis.Services;
    using System.Text.RegularExpressions;
    using System.Data.SqlClient;
    using System.Text;
    using System.Net.Mail;
    using System.Net.Mime;
    using Newtonsoft.Json;
    using System.Threading;
    using System.Security.Cryptography.X509Certificates;
    using System.Configuration;

    #endregion

    public class GMailHelper
    {
        #region Private Members

        private static string[] Scopes = { GmailService.Scope.GmailModify };
        private static string _applicationName = "Jp's Synchronization Application";
        private static string _userId = "me";
        private static string _processedLbl = "ProcessedMail";
        private static string _docRootDirectory = HttpContext.Current.Server.MapPath("~/SynchDocuments");
        private static string _serviceAccountKey = HttpContext.Current.Server.MapPath(ConfigurationManager.AppSettings["GoogleServiceAccountKey"]);
        private static string _serviceAccountUser = Convert.ToString(ConfigurationManager.AppSettings["GoogleServiceAccountEmail"]);

        #endregion

        #region Helper Actions

        public static GmailService InitilizeService(UserCredential credential)
        {
            try
            {
                var gmailService = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = _applicationName,
                });
                return gmailService;
            }
            catch (Exception e)
            {
                throw;
            }
        }

        public static GmailService InitilzeServiceByServiceAccount()
        {
            try
            {
                var json = System.IO.File.ReadAllText(_serviceAccountKey);
                var cred = JsonConvert.DeserializeObject<ServiceAccountCred>(json);
                GmailService service = new GmailService();

                var saCredential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(cred.client_email)
                {
                    Scopes = new[] { GmailService.Scope.GmailModify },
                    User = _serviceAccountUser
                }.FromPrivateKey(cred.private_key));

                if (saCredential.RequestAccessTokenAsync(CancellationToken.None).Result)
                {
                    var initilizer = new BaseClientService.Initializer() { HttpClientInitializer = saCredential, };

                    Google.Apis.Auth.OAuth2.Responses.TokenResponse toke = saCredential.Token;
                    service = new GmailService(initilizer);
                }
                return service;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                return null;
            }
        }

        public static List<Google.Apis.Gmail.v1.Data.Message> ListInboxMessages(GmailService service, String query)
        {
            List<Google.Apis.Gmail.v1.Data.Message> result = new List<Google.Apis.Gmail.v1.Data.Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(_userId);
            request.Q = query;
            request.LabelIds = "INBOX";

            do
            {
                try
                {
                    ListMessagesResponse response = request.Execute();
                    if (response.Messages != null && response.Messages.Count > 0)
                    {
                        result.AddRange(response.Messages);
                        request.PageToken = response.NextPageToken;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return result;
        }

        public static void ProcessMessage(GmailService service, String messageId)
        {
            try
            {
                Google.Apis.Gmail.v1.Data.Message message = GetMessage(service, _userId, messageId);

                string mailSubject = message.Payload.Headers.FirstOrDefault(x => x.Name.ToLower() == "subject").Value;
                //string sender = message.Payload.Headers.FirstOrDefault(x => x.Name.ToLower() == "from").Value;//.Replace("<", "").Replace(">", "");
                string sender = message.Payload.Headers.FirstOrDefault(x => x.Name.ToLower() == "return-path").Value.Replace("<", "").Replace(">", "");
                string receiver = message.Payload.Headers.FirstOrDefault(x => x.Name.ToLower() == "delivered-to").Value;//.Replace("<", "").Replace(">", "");
                string msgId = message.Payload.Headers.FirstOrDefault(x => x.Name.ToLower() == "message-id").Value;

                //Get Lable Id for "Ab.Processed"
                Label abProcessedLable = GetLabelByName(service, _processedLbl);
                if (abProcessedLable == null) // Create New Lable if Not created yet
                {
                    abProcessedLable = CreateLabel(service, _processedLbl);
                }

                // Extract Domain & Patient
                string strErrorMsg = "";
                string domain, patientNumber = string.Empty;

                string strRegexSubject = Regex.Replace(mailSubject, "[,|:|;]", "?");
                var subjectArray = strRegexSubject.Split('?');

                if (subjectArray == null || subjectArray.Length == 0)
                {
                    strErrorMsg = "Mail subject not provided OR Empty mail subject.";
                }
                else if (subjectArray.Length < 3)
                {
                    strErrorMsg = "Invalid Mail Subject: " + mailSubject;
                }
                else
                {
                    domain = subjectArray[1].Replace(")", "").Replace("(", "").Replace("]", "").Replace("[", "").Trim();
                    patientNumber = subjectArray[2].Replace(")", "").Replace("(", "").Replace("]", "").Replace("[", "").Trim();

                    IList<MessagePart> attachments = message.Payload.Parts.Where(x => !string.IsNullOrEmpty(x.Filename)).ToList();
                    if (attachments != null)
                    {
                        foreach (MessagePart part in attachments)
                        {
                            string folderPath = _docRootDirectory + "/" + domain + "/Docfolder/" + patientNumber + "/";
                            string filePath = Path.Combine(folderPath, part.Filename);


                            // TODO : Proccessing Mail & Save Attachment Logic 

                            if (string.IsNullOrEmpty(strErrorMsg))
                            {
                                String attId = part.Body.AttachmentId;
                                MessagePartBody attachPart = service.Users.Messages.Attachments.Get(_userId, messageId, attId).Execute();
                                String attachData = attachPart.Data.Replace('-', '+').Replace('_', '/');

                                byte[] data = Convert.FromBase64String(attachData);
                                if (!Directory.Exists(folderPath))
                                {
                                    Directory.CreateDirectory(folderPath);
                                }
                                File.WriteAllBytes(filePath, data);
                            }

                        }
                    }
                    else
                    {
                        strErrorMsg = "Error: Attachment not found.";
                    }

                    //Mark mail as read
                    MoifyMessage(service, _userId, messageId, new List<String>() { abProcessedLable.Id }, new List<String>() { "UNREAD", "INBOX" });
                }


                // SEND FeedBack Success/Failure.
                string feedbackMailSubject = string.IsNullOrEmpty(strErrorMsg) ? ("Accepted: " + mailSubject) : ("Error: " + mailSubject);
                string feedbackMailBody = string.IsNullOrEmpty(strErrorMsg) ? "File Accepted" : strErrorMsg;
                SendFeedBackMail(sender, feedbackMailSubject, feedbackMailBody);
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
        }

        public static Google.Apis.Gmail.v1.Data.Message GetMessage(GmailService service, String userId, String messageId)
        {
            try
            {
                return service.Users.Messages.Get(userId, messageId).Execute();
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public static Google.Apis.Gmail.v1.Data.Message MoifyMessage(GmailService service, String userId, String messageId, List<String> labelsToAdd, List<String> labelsToRemove)
        {
            ModifyMessageRequest mods = new ModifyMessageRequest();
            mods.RemoveLabelIds = labelsToRemove;
            mods.AddLabelIds = labelsToAdd;

            try
            {
                return service.Users.Messages.Modify(mods, userId, messageId).Execute();
            }
            catch (Exception e)
            {
                return null;
            }
        }

        #endregion

        #region Private Action

        public static Label GetLabelByName(GmailService service, String labelName)
        {
            try
            {
                labelName = labelName.ToLower();
                return service.Users.Labels.List(_userId).Execute().Labels.FirstOrDefault(x => x.Name.ToLower() == labelName);
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public static Label CreateLabel(GmailService service, String labelName)
        {
            try
            {
                Label labelProcessed = new Label() { Name = labelName };
                service.Users.Labels.Create(labelProcessed, _userId).Execute();
                return GetLabelByName(service, labelName);
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private static void SendFeedBackMail(string toEmail, string subject, string mailBody)
        {
            string FromEmail = "";
            string SmtpServerName = "";
            string SmtpPort = "587";
            string SmtpUserName = "";
            string SmtpPassword = "";

            StringBuilder HTML = new StringBuilder();
            HTML.Append(mailBody);

            try
            {
                System.Net.Mail.MailMessage MailMsg = new System.Net.Mail.MailMessage();
                MailMsg.Subject = subject;

                MailMsg.To.Add(new MailAddress(toEmail, ""));
                MailMsg.From = new MailAddress(FromEmail, "");
                MailMsg.IsBodyHtml = true;

                MailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(Convert.ToString(HTML), null, MediaTypeNames.Text.Plain));
                MailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(Convert.ToString(HTML), null, MediaTypeNames.Text.Html));
                SmtpClient SmtpClient = new SmtpClient(SmtpServerName, Convert.ToInt32(SmtpPort));
                SmtpClient.Credentials = new System.Net.NetworkCredential(SmtpUserName, SmtpPassword);
                SmtpClient.Send(MailMsg);
            }
            catch (Exception ex)
            {
                return;
            }
        }

        #endregion
    }

    public class ServiceAccountCred
    {
        public string type { get; set; }
        public string project_id { get; set; }
        public string private_key_id { get; set; }
        public string private_key { get; set; }
        public string client_email { get; set; }
        public string client_id { get; set; }
        public string auth_uri { get; set; }
        public string token_uri { get; set; }
        public string auth_provider_x509_cert_url { get; set; }
        public string client_x509_cert_url { get; set; }
    }
}