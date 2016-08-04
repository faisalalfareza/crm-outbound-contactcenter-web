using MVC_CRUD.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static MVC_CRUD.Models.AuthModels;

namespace MVC_CRUD.Controllers
{
    public class AuthController : Controller
    {
        //
        // GET: /Auth/

        public ActionResult Login()
        {
            //if (Session["UserId"] != null && Session["UserName"] != null && Session["RoleId"] != null)
            //{
            //    return RedirectToAction("crmAgent", "Home");
            //}

            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Login(Login model)
        {
            if (model.Password == null && model.UserName == null)
            {
                return RedirectToAction("Login");
            }
            using (DB_CRM_CCEntities db = new DB_CRM_CCEntities())
            {
                String Password = Helper.EncodePassword(model.Password, "th1siScRmc0nT4Ctc3nTeR!!!");
                var loginValid = db.TR_User.Where(x => x.Email.Equals(model.UserName) && x.UserPass.Equals(Password)).FirstOrDefault();
                if (loginValid != null)
                {
                    Session["UserId"] = loginValid.UserId;
                    Session["UserName"] = loginValid.UserName;
                    Session["RoleId"] = loginValid.RoleId;

                    if (Session["RoleId"].Equals(2))
                    {
                        return RedirectToAction("callStatus", "Home");
                    }
                    else if (Session["RoleId"].Equals(1))
                    {
                        return RedirectToAction("masterUser", "Home");
                    }
                    else if (Session["RoleId"].Equals(3))
                    {
                        return RedirectToAction("masterExpired", "Home");
                    }
                    return RedirectToAction("crmAgent", "Home");
                }

            }

            return View();
        }

        public ActionResult Logout()
        {
            Session.Abandon();
            return RedirectToAction("Login", "Auth");
        }

    }
}
