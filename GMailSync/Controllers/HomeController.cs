using GMailSync.Helper;
using Google.Apis.Auth.OAuth2.Mvc;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace GMailSync.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        #region Service Account
        
        public ActionResult MailSync()
        {
            try
            {
                var service = GMailHelper.InitilzeServiceByServiceAccount();
                var lstMessages = GMailHelper.ListInboxMessages(service, "Save");
                foreach (var objMessage in lstMessages)
                {
                    GMailHelper.ProcessMessage(service, objMessage.Id);
                }
                string msg = lstMessages.Count > 0 ? string.Format("{0} mail(s) synchronized.", lstMessages.Count) : "No more mail to synchronize";
                return Json(new { status = "success", message = msg }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = "failed", message = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        } 

        #endregion

        #region Oauth ClientID

        public async Task<ActionResult> IndexAsync(CancellationToken cancellationToken)
        {
            var result = await new AuthorizationCodeMvcApp(this, new GmailAppFlowMetadata()).
                AuthorizeAsync(cancellationToken);

            if (result.Credential != null)
            {
                var service = GMailHelper.InitilizeService(result.Credential);
                var lstMessages = GMailHelper.ListInboxMessages(service, "Save");
                foreach (var objMessage in lstMessages)
                {
                    GMailHelper.ProcessMessage(service, objMessage.Id);
                }
                string msg = lstMessages.Count > 0 ? "" : "No more mail to Synchronize";
                if (Request.HttpMethod == "GET")
                    return RedirectToAction("SynchSuccess", "Home");
                else
                    return Json(new { status = "success", message = msg }, JsonRequestBehavior.AllowGet);
            }
            else
            {
                return Json(new { status = "failed", redirectURL = result.RedirectUri }, JsonRequestBehavior.AllowGet);
            }
        }

        [AcceptVerbs(HttpVerbs.Get)]
        public ActionResult SynchSuccess()
        {
            return View();
        }

        #endregion
    }
}