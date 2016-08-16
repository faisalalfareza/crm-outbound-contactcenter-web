using BimaExpress.UI.Common;
using MVC_CRUD.Models;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
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

            Session["CaptchaImageText"] = GenerateRandomCode();
            return View();
        }

        public ActionResult ChangePassword()
        {
            if (Session["UserId"] != null && Session["IsVerified"] != null)
            {
                Session["CaptchaImageText"] = GenerateRandomCode();
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }


        private string GenerateRandomCode()
        {
            var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
            var random = new Random();
            string c = new string(
                Enumerable.Repeat(chars, 6)
                          .Select(s => s[random.Next(s.Length)])
                          .ToArray());
            return c;
        }


        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult LoginCheck(Login model)
        {
            if (model.Password == null || model.UserName == null || model.Captcha == null)
            {
                return RedirectToAction("Login");
            }
            else
            {
                if (Session["CaptchaImageText"].ToString().ToUpper() == model.Captcha.ToUpper())
                {
                    using (DB_CRM_CCEntitiesNew db = new DB_CRM_CCEntitiesNew())
                    {
                        String Password = Helper.EncodePassword(model.Password, "th1siScRmc0nT4Ctc3nTeR!!!");
                        var loginValid = db.TR_User.Where(x => x.Email.Equals(model.UserName) && x.UserPass.Equals(Password) && x.UserStatus == 1).FirstOrDefault();

                        if (loginValid != null)
                        {
                            Session["UserId"] = loginValid.UserId;
                            Session["UserName"] = loginValid.UserName;
                            Session["RoleId"] = loginValid.RoleId;

                            if (loginValid.IsVerified == false)
                            {
                                Session["IsVerified"] = 0;
                                Session["Email"] = model.UserName;
                                return Redirect("~/Auth/ChangePassword");
                            }
                            else {
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
                        else
                        {
                            return Redirect("~/Auth/Login?msg=User Not Registered");
                        }
                    }
                }
                else {
                    return Redirect("~/Auth/Login?msg=Invalid Captcha");
                }
            }

            //return View();
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult ChangePasswordAct(int? UserId, string NewPassword, string VerifryPassword, string Captcha)
        {
            if (NewPassword == null || VerifryPassword == null || UserId == null || Captcha == null)
            {
                return RedirectToAction("ChangePassword");
            }
            else
            {
                if (Session["CaptchaImageText"].ToString().ToUpper() == Captcha.ToUpper())
                {
                    if (NewPassword == VerifryPassword)
                    {
                        using (DB_CRM_CCEntitiesNew db = new DB_CRM_CCEntitiesNew())
                        {
                            var dtUser = db.TR_User.Where(p => p.UserId == UserId);
                            if (dtUser.Count() > 0)
                            {
                                var dtUsers = dtUser.FirstOrDefault();
                                string password = Helper.EncodePassword(NewPassword, "th1siScRmc0nT4Ctc3nTeR!!!");

                                dtUsers.UserPass = password;
                                dtUsers.IsVerified = true;
                                db.SaveChanges();

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
                            else {
                                return Redirect("~/Auth/Login?msg=User Not Registered");
                            }
                        }
                    }
                    else
                    {
                        return Redirect("~/Auth/ChangePassword?msg=Password Invalid");
                    }
                }
                else
                {
                    return Redirect("~/Auth/ChangePassword?msg=Invalid Captcha");
                }
            }

            //return View();
        }

        public ActionResult Logout()
        {
            Session.Abandon();
            return RedirectToAction("Login", "Auth");
        }

        public ActionResult Captcha()
        {
            if (Session["CaptchaImageText"] != null)
            {
                CaptchaText ci = new CaptchaText(this.Session["CaptchaImageText"].ToString(), 310, 40,
                                          "Franklin Gothic Demi Cond");
                this.Response.Clear();
                this.Response.ContentType = "image/jpeg";
                ci.Image.Save(this.Response.OutputStream, ImageFormat.Jpeg);
                ci.Dispose();
            }

            return View();
        }

    }
}
