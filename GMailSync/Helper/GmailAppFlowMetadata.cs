namespace GMailSync.Helper
{
    #region Usings

    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Auth.OAuth2.Flows;
    using Google.Apis.Auth.OAuth2.Mvc;
    using Google.Apis.Auth.OAuth2.Requests;
    using Google.Apis.Gmail.v1;
    using Google.Apis.Util.Store;
    using System.IO;
    using System.Web;

    #endregion

    public class GmailAppFlowMetadata : FlowMetadata
    {
        public GmailAppFlowMetadata()
        {
            if (!Directory.Exists(credPath))
            {
                Directory.CreateDirectory(credPath);
            }
        }

        #region Private Members

        private static string _apiCredentialDirectory = HttpContext.Current.Server.MapPath("~/GoogleAPI/");
        private static string _client_secret_json = HttpContext.Current.Server.MapPath("~/GoogleAPI/client_secret.json");
        private static string credPath = _apiCredentialDirectory + "/credentials/";
        private static string[] Scope = { GmailService.Scope.GmailModify };

        #endregion

        #region Overridee Props & Actions


        private static readonly IAuthorizationCodeFlow flow =
            new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
            {
                ClientSecrets = GetClientIDSecret(),
                Scopes = Scope,
                DataStore = new FileDataStore(credPath + "/GMailApi-Credential.json")
            });


        public override IAuthorizationCodeFlow Flow
        {
            get { return flow; }
        }

        public override string GetUserId(System.Web.Mvc.Controller controller)
        {
            //var user = controller.Session["user"];
            //if (user == null)
            //{
            //    user = Guid.NewGuid();
            //    controller.Session["user"] = user;
            //}
            //return user.ToString();
            return "TestUser";
        }

        #endregion

        #region Actions

        private static ClientSecrets GetClientIDSecret()
        {
            using (var stream = new FileStream(_client_secret_json, FileMode.Open, FileAccess.Read))
            {
                return GoogleClientSecrets.Load(stream).Secrets;
            }
        }

        #endregion
    }

    public class dsAuthorizationCodeFlow : GoogleAuthorizationCodeFlow
    {
        public dsAuthorizationCodeFlow(Initializer initializer)
            : base(initializer) { }

        public override AuthorizationCodeRequestUrl
                       CreateAuthorizationCodeRequest(string redirectUri)
        {
            return base.CreateAuthorizationCodeRequest("http://localhost:50376/AuthCallback/SynchPatientDocumentsAsync");
        }
    }
}