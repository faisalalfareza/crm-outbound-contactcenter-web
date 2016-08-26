
using MVC_CRUD.Models;
using System;
using System.Collections.Generic;
using System.IO;
//using System.Data.Objects.SqlClient;
using System.Data.Entity.SqlServer;
using System.Linq;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections.Specialized;
using static MVC_CRUD.Models.contactCenterModels;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace MVC_CRUD.Controllers
{
    public class HomeController : Controller
    {
        DB_CRM_CCEntities db = new DB_CRM_CCEntities();
        public ActionResult crmAgent()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "4")
            {
                int UserId = Convert.ToInt32(Session["UserId"].ToString());
                var result = (from contact in db.TR_Contact.Where(p=> p.UserId == UserId)
                             join callH in (from call in db.TT_CallHistory
                                            group call by call.ContactId into grouping
                                            select new
                                            {
                                                ContactId = grouping.Key,
                                                BeginCall = grouping.Min(prod => prod.CallDate),
                                                CallBack = grouping.Max(prod => prod.CallDate)
                                            }) on contact.ContactId equals callH.ContactId
                             join custP in db.TT_CustomerProject on contact.CustProId equals custP.CustProId
                             join cust in db.TR_Customer on custP.CustomerId equals cust.CustomerId
                             //where contact.UserId == UserId
                             select new contactCenterModels.HistoryCall
                             {
                                 ContactName = contact.ContactName,
                                 ContactId = contact.ContactId,
                                 CustomerName = contact.CustomerName,
                                 CustProName = custP.CustProName,
                                 CustomerContactId = contact.CustomerContactId,
                                 CallStatus = contact.CallStatus,
                                 SubStatus = contact.SubStatus,
                                 BeginCall = callH.BeginCall,
                                 CallBack = callH.CallBack,
                                 AgingAgent = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), DateTime.Now),
                                 AgingData = SqlFunctions.DateDiff("d", contact.CreatedOn, DateTime.Now),
                                 Reach = (from h in db.TT_CallHistory
                                          where h.ContactId == contact.ContactId
                                          select h).Count()
                             }).ToList();

                //ViewBag.closing = db.TT_CallHistory.Count(h => h.CallStatus.Equals("Closing") && h.UserId == UserId);
                //ViewBag.prospect = db.TT_CallHistory.Count(h => h.CallStatus.Equals("Prospect") && h.UserId == UserId);
                //ViewBag.notprospect = db.TT_CallHistory.Count(h => h.CallStatus != "Closing" && h.CallStatus != "Prospect" && h.UserId == UserId);
                //ViewBag.newdata = db.TR_Contact.Count(h => h.ContactStatus == 1 && h.UserId == UserId);

                ViewBag.closing = db.TR_Contact.Count(h => h.CallStatus.Equals("Closing") && h.UserId == UserId);
                ViewBag.prospect = db.TR_Contact.Count(h => h.CallStatus.Equals("Prospect") && h.UserId == UserId);
                ViewBag.notprospect = db.TR_Contact.Count(h => h.CallStatus != "Closing" && h.CallStatus != "Prospect" && h.UserId == UserId);
                ViewBag.newdata = db.TR_Contact.Count(h => h.ContactStatus == 1 && h.UserId == UserId);

                return View(result);
            }
            else {
                return RedirectToAction("Login", "Auth");
            }
        }

        public ActionResult crmNewcall()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "4")
            {
                int UserId = Convert.ToInt32(Session["UserId"].ToString());
                var result = (from contact in db.TR_Contact
                              join custP in db.TT_CustomerProject on contact.CustProId equals custP.CustProId
                              join cust in db.TR_Customer on custP.CustomerId equals cust.CustomerId
                              where contact.ContactStatus == 1 && contact.UserId == UserId
                              select new contactCenterModels.HistoryCall()
                              {
                                  ContactName = contact.ContactName,
                                  ContactId = contact.ContactId,
                                  CustomerContactId = contact.CustomerContactId,
                                  CustomerName = cust.CustomerName,
                                  CustProName = custP.CustProName,
                                  AgingData = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), DateTime.Now)
                              }).ToList();
                return View(result);
            } else {
                return RedirectToAction("Login", "Auth");
            }
        }

        public ActionResult call(int id)
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "4")
            {
                IEnumerable<contactCenterModels.HistoryDetail>
                historydetail = (from callH in db.TT_CallHistory
                                 join contact in db.TR_Contact on callH.ContactId equals contact.ContactId
                                 join agent in db.TR_User on callH.UserId equals agent.UserId
                                 where contact.ContactId.Equals(id)
                                 select new contactCenterModels.HistoryDetail()
                                 {
                                     CallDate = callH.CallDate,
                                     ContactPhone = contact.ContactPhone,
                                     CallStatus = callH.CallStatus,
                                     SubStatus = callH.SubStatus,
                                     Remarks = callH.Remarks,
                                     AgingAgent = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), callH.CallDate),//1,//,
                                     UserName = agent.UserName
                                 }).OrderByDescending(x => x.CallDate);
                TR_Contact call = db.TR_Contact.Where(p => p.ContactId.Equals(id)).FirstOrDefault();

                TT_CustomerProject paramProject = db.TT_CustomerProject.Where(p => p.CustProId == call.CustProId).FirstOrDefault();
                String CustomerName = (from h in db.TR_Customer
                                       where h.CustomerId == paramProject.CustomerId
                                       select h.CustomerName).FirstOrDefault();
                ViewBag.call = call;

                ViewBag.historydetail = historydetail.ToList();
                ViewBag.param = paramProject;
                ViewBag.CustomerName = CustomerName + " - " + paramProject.CustProName;
                return View();
            } else {
                return RedirectToAction("Login", "Auth");
            }
        }

        public ActionResult callclose(int id)
        {
            if (Session["UserId"] != null && (Session["RoleId"].ToString() == "4"  || Session["RoleId"].ToString() == "2"))
            {
                IEnumerable<contactCenterModels.HistoryDetail>
                    historydetail = (from callH in db.TT_CallHistory
                                     join contact in db.TR_Contact on callH.ContactId equals contact.ContactId
                                     join agent in db.TR_User on callH.UserId equals agent.UserId
                                     where contact.ContactId.Equals(id)
                                     select new contactCenterModels.HistoryDetail()
                                     {
                                         CallDate = callH.CallDate,
                                         ContactPhone = contact.ContactPhone,
                                         CallStatus = callH.CallStatus,
                                         SubStatus = callH.SubStatus,
                                         Remarks = callH.Remarks,
                                         AgingAgent = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), callH.CallDate),//1,//,
                                         UserName = agent.UserName
                                     }).OrderByDescending(x => x.CallDate);
                TR_Contact call = db.TR_Contact.Where(p => p.ContactId.Equals(id)).FirstOrDefault();
                TT_CallHistory callhis = db.TT_CallHistory.Where(p => p.ContactId == (id)).OrderByDescending(p=> p.CallDate).FirstOrDefault();

                TT_CustomerProject paramProject = db.TT_CustomerProject.Where(p => p.CustProId == call.CustProId).FirstOrDefault();
                String CustomerName = (from h in db.TR_Customer
                                       where h.CustomerId == paramProject.CustomerId
                                       select h.CustomerName).FirstOrDefault();
                ViewBag.call = call;
                ViewBag.callhis = callhis;
                ViewBag.historydetail = historydetail.ToList();
                ViewBag.param = paramProject;
                ViewBag.CustomerName = CustomerName + " - " + paramProject.CustProName;
                return View();
            } else {
                 return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        public JsonResult dialCallPost()
        {
            //String status = "";
            //get data collumn number dari form upload
            NameValueCollection nvc = Request.Form;
            String CallDuration = Convert.ToString(nvc["CallDuration"]);
            int id = Convert.ToInt32(nvc["id"]);
            String callstatus = Convert.ToString(nvc["callstatus"]);
            String substatus = Convert.ToString(nvc["substatus"]);
            String Remarks = Convert.ToString(nvc["Remarks"]);
            String Param1 = Convert.ToString(nvc["Param1"]);
            String Param2 = Convert.ToString(nvc["Param2"]);
            String Param3 = Convert.ToString(nvc["Param3"]);
            String Param4 = Convert.ToString(nvc["Param4"]);
            String Param5 = Convert.ToString(nvc["Param5"]);

            var dtContact = db.TR_Contact.Where(p => p.ContactId.Equals(id)).FirstOrDefault();
            if (callstatus.Equals("Closing"))
            {
                dtContact.ContactStatus = 3;
            }
            else
            {
                dtContact.ContactStatus = 2;
            }
            dtContact.Param1 = Param1;
            dtContact.Param2 = Param2;
            dtContact.Param3 = Param3;
            dtContact.Param4 = Param4;
            dtContact.Param5 = Param5;
            dtContact.CallStatus = callstatus;
            dtContact.SubStatus = substatus;
            dtContact.Remarks = Remarks;
            db.SaveChanges();

            TT_CallHistory dtcallHistory = new TT_CallHistory();
            dtcallHistory.CallDuration = CallDuration;
            dtcallHistory.CallStatus = callstatus;
            dtcallHistory.SubStatus = substatus;
            dtcallHistory.Remarks = Remarks;
            dtcallHistory.CallDate = DateTime.Now;
            dtcallHistory.ModifiedDate = DateTime.Now;
            dtcallHistory.UserId = Convert.ToInt32(Session["UserId"].ToString());
            dtcallHistory.ContactId = id;

            db.TT_CallHistory.Add(dtcallHistory);
            db.SaveChanges();

            if (callstatus.Equals("Closing"))
            {
                TT_CallResult dtcallResult = new TT_CallResult();
                dtcallResult.Param1 = Param1;
                dtcallResult.Param2 = Param2;
                dtcallResult.Param3 = Param3;
                dtcallResult.Param4 = Param4;
                dtcallResult.Param5 = Param5;
                dtcallResult.ContactId = id;
                dtcallResult.CallLogId = dtcallHistory.CallLogId;

                db.TT_CallResult.Add(dtcallResult);
                db.SaveChanges();
            }

            return Json(new
            {
                redirectUrl = Url.Action("crmAgent", "Home"),
                isRedirect = true
            });
        }

        public ActionResult masterAgent()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                IEnumerable<contactCenterModels.User>
                user = (from User in db.TR_User
                        join Role in db.TR_Role on User.RoleId equals Role.RoleId
                        join mgr in db.TR_User on User.UserManagerId equals mgr.UserId
                        where Role.RoleId == 4
                        select new contactCenterModels.User
                        {
                            UserId = User.UserId,
                            UserName = User.UserName,
                            Email = User.Email,
                            Role = Role.Rolename,
                            Team = mgr.UserName,
                            Level = User.UserSkill,
                            Active = User.UserStatus == 1 ? "Active" : "Inactive"
                        }).ToList();
                ViewBag.User = user;
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        public ActionResult upload()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                var CustPro = db.TT_CustomerProject.Where(x => x.status == 1);
                ViewBag.GroupId = new SelectList(CustPro, "CustProId", "CustProName");
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        /* Master Customer*/
        public ActionResult masterCustomer()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                return View(db.TR_Customer);
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        public ActionResult createCustomer(TR_Customer tabelCust)
        {
            if (ModelState.IsValid)
            {
                db.TR_Customer.Add(tabelCust);
                db.SaveChanges();
                return RedirectToAction("masterCustomer");
            }

            return View(db.TR_Customer);
        }

        public ActionResult updateCustomer(int? CustomerId, String CustomerName, String CustomerDesc)
        {
            TR_Customer tabelCust = db.TR_Customer.Find(CustomerId);
            if (tabelCust != null)
            {
                var dtCust = db.TR_Customer.Where(p => p.CustomerId == CustomerId).FirstOrDefault();
                dtCust.CustomerName = CustomerName;
                dtCust.CustomerDesc = CustomerDesc;
                db.SaveChanges();
            }
            return RedirectToAction("masterCustomer");
        }

        public ActionResult deleteCustomer(int id)
        {
            TR_Customer tabelCust = db.TR_Customer.Find(id);
            if (tabelCust == null)
            {
                return HttpNotFound();
            }
            db.TR_Customer.Remove(tabelCust);
            db.SaveChanges();

            return RedirectToAction("masterCustomer");
        }

        /* Master CustPro*/
        public ActionResult masterCustomerProject()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                ViewBag.CustomerName = db.TR_Customer.ToList();
                var result = (from CustP in db.TT_CustomerProject
                              join cust in db.TR_Customer on CustP.CustomerId equals cust.CustomerId
                              select new contactCenterModels.Customer()
                              {
                                  CustProId = CustP.CustProId,
                                  CustProName = CustP.CustProName,
                                  CustProExpired = CustP.CustProExpired,
                                  CustomerName = cust.CustomerName,
                                  Param1 = CustP.Param1,
                                  Param2 = CustP.Param2,
                                  Param3 = CustP.Param3,
                                  Param4 = CustP.Param4,
                                  Param5 = CustP.Param5
                              }).ToList();
                return View(result);
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        public ActionResult createCustomerProject(TT_CustomerProject tabelCustPro)
        {
            if (ModelState.IsValid)
            {
                tabelCustPro.CreatedOn = DateTime.Now;
                tabelCustPro.status = DateTime.Now >= tabelCustPro.CustProExpired ? 0 : 1;
                tabelCustPro.CreatedBy = Convert.ToString(Session["UserName"]);

                tabelCustPro.Param1 = tabelCustPro.Param1 == "" ? null : tabelCustPro.Param1;
                tabelCustPro.Param2 = tabelCustPro.Param2 == "" ? null : tabelCustPro.Param2;
                tabelCustPro.Param3 = tabelCustPro.Param3 == "" ? null : tabelCustPro.Param3;
                tabelCustPro.Param4 = tabelCustPro.Param4 == "" ? null : tabelCustPro.Param4;
                tabelCustPro.Param5 = tabelCustPro.Param5 == "" ? null : tabelCustPro.Param5;

                db.TT_CustomerProject.Add(tabelCustPro);
                db.SaveChanges();
                return RedirectToAction("masterCustomerProject");
            }

            return View();
        }

        public ActionResult updateCustomerProject(int? CustProId, String CustProName, DateTime CustProExpired, String Param1, String Param2, String Param3, String Param4, String Param5)
        {
            TT_CustomerProject tabelCustPro = db.TT_CustomerProject.Find(CustProId);
            if (tabelCustPro != null)
            {
                var dtCustPro = db.TT_CustomerProject.Where(p => p.CustProId == CustProId).FirstOrDefault();
                dtCustPro.CustProName = CustProName;
                dtCustPro.CustProExpired = CustProExpired;
                dtCustPro.status = DateTime.Now >= CustProExpired ? 0 : 1;
                dtCustPro.Param1 = Param1 == "" ? null : Param1;
                dtCustPro.Param2 = Param2 == "" ? null : Param2;
                dtCustPro.Param3 = Param3 == "" ? null : Param3;
                dtCustPro.Param4 = Param4 == "" ? null : Param4;
                dtCustPro.Param5 = Param5 == "" ? null : Param5;
                db.SaveChanges();
            }
            return RedirectToAction("masterCustomerProject");
        }

        public ActionResult updateCustomerProjectExpDate(int? CustProId, DateTime CustProExpired)
        {
            TT_CustomerProject tabelCustPro = db.TT_CustomerProject.Find(CustProId);
            if (tabelCustPro != null)
            {
                var dtCustPro = db.TT_CustomerProject.Where(p => p.CustProId == CustProId).FirstOrDefault();
                dtCustPro.CustProExpired = CustProExpired;
                dtCustPro.status = DateTime.Now >= CustProExpired ? 0 : 1;
                db.SaveChanges();
            }
            return RedirectToAction("masterExpired");
        }

        public ActionResult deleteCustomerProject(int id)
        {
            TT_CustomerProject tabelCustPro = db.TT_CustomerProject.Find(id);
            if (tabelCustPro == null)
            {
                return HttpNotFound();
            }
            db.TT_CustomerProject.Remove(tabelCustPro);
            db.SaveChanges();

            return RedirectToAction("masterCustomerProject");
        }

        public ActionResult withdraw()
        {
            if (Session["UserId"] != null && (Session["RoleId"].ToString() == "3" || Session["RoleId"].ToString() == "2"))
            {
                String UserName = Session["UserName"].ToString();
                int UserIds = Convert.ToInt32(Session["UserId"].ToString());
                int[] UserId = (from user in db.TR_User
                                where user.UserManagerId == UserIds
                                select user.UserId).ToArray();
                var custPro = (from uPro in db.TT_UserProject
                               join custP in db.TT_CustomerProject on uPro.CustProId equals custP.CustProId
                               where UserId.Contains((int)uPro.UserId) && custP.status == 1
                               select new CustomerProject
                               {
                                   CustProId = custP.CustProId,
                                   CustProName = custP.CustProName
                               }).Distinct().ToList().OrderBy(p=> p.CustomerName);
                ViewBag.GroupId = new SelectList(custPro, "CustProId", "CustProName");
                return View();
            } else {
                return RedirectToAction("Login", "Auth");
            }
        }

        public int GetTotalWithdraw(int CustProId, string From, string To, string Status)
        {
            DateTime DateFroms = DateTime.Parse(From);
            DateTime DateTos = DateTime.Parse(To);
            int UserId = Convert.ToInt32(Session["UserId"]);

            var data = new List<TR_Contact>();
            if (Status == "notselected")
            {
                data = (from c in db.TR_Contact.Where(x => x.CustProId == CustProId && x.CreatedOn >= DateFroms && x.CreatedOn <= DateTos && x.ContactStatus != 6)
                        join u in db.TR_User on c.UserId equals u.UserId
                        where u.UserManagerId == UserId
                        select c).ToList();
            }
            else
            {
                data = (from c in db.TR_Contact.Where(x => x.CustProId == CustProId && x.CreatedOn >= DateFroms && x.CreatedOn <= DateTos && x.ContactStatus != 6 && x.CallStatus == Status)
                        join u in db.TR_User on c.UserId equals u.UserId
                        where u.UserManagerId == UserId
                        select c).ToList();
            }

            return data.Count();
        }

        public int GetTotalWithdrawExp(int CustProId, string From, string To, string Status)
        {
            DateTime DateFroms = DateTime.Parse(From);
            DateTime DateTos = DateTime.Parse(To);
           
            var data = new List<TR_Contact>();
            if (Status == "notselected")
            {
                data = (from c in db.TR_Contact.Where(x => x.CustProId == CustProId && x.CreatedOn >= DateFroms && x.CreatedOn <= DateTos && x.ContactStatus != 6)
                        select c).ToList();
            }
            else
            {
                data = (from c in db.TR_Contact.Where(x => x.CustProId == CustProId && x.CreatedOn >= DateFroms && x.CreatedOn <= DateTos && x.ContactStatus != 6 && x.CallStatus == Status)
                        select c).ToList();
            }

            return data.Count();
        }


        public ActionResult createAgent()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                ViewBag.Manager = db.TR_User.Where(x => x.RoleId == 2).ToList();
                ViewBag.CustProject = (from custP in db.TT_CustomerProject
                                       join cust in db.TR_Customer on custP.CustomerId equals cust.CustomerId
                                       where custP.status == 1
                                       select new contactCenterModels.Customer
                                       {
                                           CustProId = custP.CustProId,
                                           CustProName = custP.CustProName,
                                           CustProExpired = custP.CustProExpired,
                                           CustomerName = cust.CustomerName
                                       }
                                      ).ToList();
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        public ActionResult updateAgent(int? id)
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                ViewBag.Manager = db.TR_User.Where(x => x.RoleId == 2).ToList();
                ViewBag.User = db.TR_User.Where(x => x.UserId == id).FirstOrDefault();
                ViewBag.CustProjectAda = (from userProject in db.TT_UserProject
                                          join custP in db.TT_CustomerProject on userProject.CustProId equals custP.CustProId
                                          join cust in db.TR_Customer on custP.CustomerId equals cust.CustomerId
                                          where userProject.UserId == id
                                          select new contactCenterModels.Customer
                                          {
                                              CustProId = custP.CustProId,
                                              CustProName = custP.CustProName,
                                              CustProExpired = custP.CustProExpired,
                                              CustomerName = cust.CustomerName
                                          }
                                      ).ToList();
                ViewBag.CustProjectTiada = (from custP in db.TT_CustomerProject
                                            join cust in db.TR_Customer on custP.CustomerId equals cust.CustomerId
                                            where !(from userProject in db.TT_UserProject
                                                    join custP in db.TT_CustomerProject on userProject.CustProId equals custP.CustProId
                                                    where userProject.UserId == id && custP.status == 1
                                                    select custP.CustProId).Contains(custP.CustProId)
                                            select new contactCenterModels.Customer
                                            {
                                                CustProId = custP.CustProId,
                                                CustProName = custP.CustProName,
                                                CustProExpired = custP.CustProExpired,
                                                CustomerName = cust.CustomerName
                                            }
                                      ).ToList();
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        public ActionResult addAgent(String Email, String RoleId, int UserManager, String UserSkill, ICollection<int> CustProId)
        {
            String UserPass = Helper.EncodePassword("Password1!", "th1siScRmc0nT4Ctc3nTeR!!!");
            String alias = Email.Substring(0, Email.IndexOf('@')).Replace('.', ' ');

            var dtUser = db.TR_User.Where(p => p.Email == Email);

            if (dtUser.Count() < 1)
            {
                TR_User user = new TR_User();
                user.RoleId = 4;
                user.Email = Email;
                user.UserName = alias;
                user.UserPass = UserPass;
                user.UserManagerId = UserManager;
                user.UserSkill = UserSkill;
                user.UserStatus = 1;
                user.CreatedOn = DateTime.Now;
                user.CreatedBy = Convert.ToString(Session["UserName"]);
                user.IsVerified = false;
                db.TR_User.Add(user);
                db.SaveChanges();

                if (CustProId != null)
                    foreach (var i in CustProId)
                    {
                        TT_UserProject userProject = new TT_UserProject();
                        userProject.CustProId = i;
                        userProject.CreatedBy = Convert.ToString(Session["UserName"]);
                        userProject.UserId = user.UserId;
                        userProject.CreatedOn = DateTime.Now;
                        db.TT_UserProject.Add(userProject);
                        db.SaveChanges();
                    }

                return RedirectToAction("masterAgent");
            }
            else {
                return Redirect("~/Home/createAgent?msg=Warning! Duplicate Email");
            }
        }

        [HttpPost]
        public ActionResult editAgent(int UserId, String Email, int UserManager, String UserSkill, int UserStatus)
        {
            String alias = Email.Substring(0, Email.IndexOf('@')).Replace('.', ' ');

            var dtUser = db.TR_User.Where(p => p.Email == Email && p.UserId != UserId);

            if (dtUser.Count() < 1)
            {
                TR_User user = db.TR_User.Where(x => x.UserId == UserId).FirstOrDefault();
                user.UserName = alias;
                user.Email = Email;
                user.UserManagerId = UserManager;
                user.UserSkill = UserSkill;
                user.UserStatus = UserStatus;
                db.SaveChanges();

                if (UserStatus == 0)
                {
                    var Contact = db.TR_Contact.Where(x => x.UserId == UserId).ToList();
                    foreach (var i in Contact)
                    {
                        i.UserId = 0;
                        i.ContactStatus = 0;
                    }
                    db.SaveChanges();
                }

                return RedirectToAction("masterAgent");
            }
            else {
                return RedirectToAction("updateAgent/"+ UserId);
            }  
        }

        [HttpPost]
        public ActionResult ExportToExcel(int GroupId, String callstatus, String DateFrom, String DateTo)
        {
            DateTime DateFroms = DateTime.Parse(DateFrom);
            DateTime DateTos = DateTime.Parse(DateTo);
            int UserId = Convert.ToInt32(Session["UserId"]);

            var data = new List<TR_Contact>();
            if (callstatus == "notselected")
            {
                data = (from c in db.TR_Contact.Where(x => x.CustProId == GroupId && x.CreatedOn >= DateFroms && x.CreatedOn <= DateTos && x.ContactStatus != 6)
                        join u in db.TR_User on c.UserId equals u.UserId
                        where u.UserManagerId == UserId
                        select c).ToList();
            }
            else
            {
                data = (from c in db.TR_Contact.Where(x => x.CustProId == GroupId && x.CreatedOn >= DateFroms && x.CreatedOn <= DateTos && x.ContactStatus != 6 && x.CallStatus == callstatus)
                        join u in db.TR_User on c.UserId equals u.UserId
                        where u.UserManagerId == UserId
                        select c).ToList();
            }

            foreach (var item in data)
            {
                item.ContactStatus = 6;
              //  item.UserId = 0;

                var history = db.TT_CallHistory.Where(x => x.ContactId == item.ContactId).ToList();
                foreach (var item1 in history)
                {
                    db.TT_CallHistory.Remove(item1);
                }
            }
            db.SaveChanges();

            String custProName = (from custPro in db.TT_CustomerProject
                                  where custPro.CustProId == GroupId
                                  select custPro.CustProName).FirstOrDefault();

            // set the data source
            GridView gridview = new GridView();
            gridview.DataSource = data;
            gridview.DataBind();

            // Clear all the content from the current response
            Response.ClearContent();
            Response.Buffer = true;

            // set the header
            Response.AddHeader("content-disposition", "attachment;filename = Withdraw " + custProName + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";

            // create HtmlTextWriter object with StringWriter
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    // render the GridView to the HtmlTextWriter
                    gridview.RenderControl(htw);
                    // Output the GridView content saved into StringWriter
                    Response.Output.Write(sw.ToString());
                    Response.Flush();
                    Response.End();
                }
            }

            return RedirectToAction("withdraw");
        }

        public ActionResult deleteAgent(int id)
        {
            var UserProject = db.TT_UserProject.Where(x => x.UserId == id).ToList();
            foreach (var item in UserProject)
            {
                db.TT_UserProject.Remove(item);
                db.SaveChanges();
            }

            TR_User tabelUser = db.TR_User.Find(id);
            if (tabelUser == null)
            {
                return HttpNotFound();
            }
            db.TR_User.Remove(tabelUser);
            db.SaveChanges();

            return RedirectToAction("masterAgent");
        }

        [HttpPost]
        public JsonResult deleteUserProject()
        {
            NameValueCollection nvc = Request.Form;
            int? CustProId = Convert.ToInt32(nvc["CustProId"]);
            int UserId = Convert.ToInt32(nvc["Id"]);

            TT_UserProject UserProject = db.TT_UserProject.Where(x => x.CustProId == CustProId && x.UserId == UserId).FirstOrDefault();
            db.TT_UserProject.Remove(UserProject);
            db.SaveChanges();

            return Json(new
            {
                Success = true
            });
        }

        [HttpPost]
        public JsonResult addUserProject()
        {
            NameValueCollection nvc = Request.Form;
            int? CustProId = Convert.ToInt32(nvc["CustProId"]);
            int UserId = Convert.ToInt32(nvc["Id"]);

            TT_UserProject UserProject = new TT_UserProject();
            UserProject.CustProId = CustProId;
            UserProject.CreatedOn = DateTime.Now;
            UserProject.CreatedBy = Convert.ToString(Session["UserName"]);
            UserProject.UserId = UserId;
            db.TT_UserProject.Add(UserProject);
            db.SaveChanges();

            return Json(new
            {
                Success = true
            });
        }

        /*  Start : Master User */

        public ActionResult masterUser()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                IEnumerable<contactCenterModels.User>
                user = (from User in db.TR_User.Where(x => x.RoleId != 1 && x.RoleId != 4)
                        from mgr in db.TR_User.Where(x=> x.UserId == User.UserManagerId).DefaultIfEmpty()
                        join Role in db.TR_Role on User.RoleId equals Role.RoleId
                        select new contactCenterModels.User
                        {
                            UserId = User.UserId,
                            Email = User.Email,
                            UserName = User.UserName,
                            Role = Role.Rolename,
                            Team = mgr.UserName,
                            Level = User.UserSkill,
                            Active = User.UserStatus == 1 ? "Active" : "Inactive"
                        }).ToList();
                ViewBag.User = user;

                return View();
            }
            else {
                return RedirectToAction("Login", "Auth");
            }
        }


        public ActionResult createUser()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                ViewBag.Manager = db.TR_User.Where(x => x.RoleId == 1).ToList();
                ViewBag.CustProject = db.TT_CustomerProject.ToList();
                ViewBag.Role = db.TR_Role.Where(x => x.RoleId != 4 && x.RoleId != 1).ToList();

                return View();
            }
            else {
                return RedirectToAction("Login", "Auth");
            }
            
        }

        [HttpPost]
        public ActionResult addUser(String Email, String Username, int RoleId, String UserSkill)
        {
            String UserPass = Helper.EncodePassword("Password1!", "th1siScRmc0nT4Ctc3nTeR!!!");
            
            var dtUser = db.TR_User.Where(p => p.Email == Email);

            if (dtUser.Count() < 1)
            {
                TR_User user = new TR_User();
                user.RoleId = RoleId;
                user.Email = Email;
                user.UserName = Username;
                user.UserPass = UserPass;
                //user.UserSkill = UserSkill;
                user.UserStatus = 1;
                user.CreatedOn = DateTime.Now;
                user.CreatedBy = Convert.ToString(Session["UserName"]);
                user.IsVerified = false;
                db.TR_User.Add(user);

                db.SaveChanges();
                return RedirectToAction("masterUser");
            }
            else {
                return Redirect("~/Home/createUser?msg=Warning! Duplicate Email");
            }
            
        }

        public ActionResult updateUser(int? id)
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                ViewBag.Manager = db.TR_User.Where(x => x.RoleId == 1).ToList();
                ViewBag.User = db.TR_User.Where(x => x.UserId == id).FirstOrDefault();

                ViewBag.Role = db.TR_Role.Where(x => x.RoleId != 4 && x.RoleId != 1).ToList();
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        public ActionResult editUser(int UserId, String Email, String Username, int RoleId, String UserPass, String UserSkill, int UserStatus)
        {
            //UserPass = Helper.EncodePassword(UserPass, "th1siScRmc0nT4Ctc3nTeR!!!");
            //String alias = Email.Substring(0, Email.IndexOf('@')).Replace('.', ' ');

            if (UserId != 1) //Admin
            {
                var dtUser = db.TR_User.Where(p => p.Email == Email && p.UserId != UserId);

                if (dtUser.Count() < 1)
                {
                    TR_User user = db.TR_User.Where(x => x.UserId == UserId).FirstOrDefault();
                    user.RoleId = RoleId;
                    user.UserName = Username;
                    user.Email = Email;
                    //user.UserPass = UserPass;
                    //user.UserSkill = UserSkill;
                    user.UserStatus = UserStatus;

                    db.SaveChanges();
                    return RedirectToAction("masterUser");
                }
                else
                {
                    return RedirectToAction("updateUser/" + UserId);
                }
            }
            else {
                return RedirectToAction("updateUser/" + UserId);
            }
            
            
        }

        public ActionResult deleteUser(int id)
        {
            if (id != 1) //Admin
            {
                TR_User tabelUser = db.TR_User.Find(id);
                if (tabelUser == null)
                {
                    return HttpNotFound();
                }
                db.TR_User.Remove(tabelUser);
                db.SaveChanges();

                return RedirectToAction("masterUser");
            }
            else {
                return RedirectToAction("masterUser");
            }
        }


        public ActionResult callStatus()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "2")
            {
                string UserName = Session["UserName"].ToString();
                int UserIds = Convert.ToInt32(Session["UserId"].ToString());

                int[] UserId = (from user in db.TR_User
                                where user.UserManagerId == UserIds
                                select user.UserId).ToArray();
                var closing = from contact in db.TR_Contact
                              join callH in (from call in db.TT_CallHistory
                                             group call by call.ContactId into grouping
                                             select new
                                             {
                                                 ContactId = grouping.Key,
                                                 BeginCall = grouping.Min(prod => prod.CallDate),
                                                 CallBack = grouping.Max(prod => prod.CallDate)
                                             }) on contact.ContactId equals callH.ContactId
                              join manager in db.TR_User on contact.UserId equals manager.UserId
                              join custP in db.TT_CustomerProject on contact.CustProId equals custP.CustProId
                              where UserId.Contains((int)contact.UserId) && contact.ContactStatus == 3
                              select new HistoryCall()
                              {
                                  ContactName = contact.ContactName,
                                  ContactId = contact.ContactId,
                                  CustomerName = contact.CustomerName,
                                  CustProName = custP.CustProName,
                                  CustomerContactId = contact.CustomerContactId,
                                  CallStatus = contact.CallStatus,
                                  SubStatus = contact.SubStatus,
                                  BeginCall = callH.BeginCall,
                                  CallBack = callH.CallBack,
                                  AgingData = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), DateTime.Now),
                                  Reach = (from h in db.TT_CallHistory
                                           where h.ContactId == contact.ContactId
                                           select h).Count()
                              };
                var suspect = from contact in db.TR_Contact
                              join callH in (from call in db.TT_CallHistory
                                             group call by call.ContactId into grouping
                                             select new
                                             {
                                                 ContactId = grouping.Key,
                                                 BeginCall = grouping.Min(prod => prod.CallDate),
                                                 CallBack = grouping.Max(prod => prod.CallDate)
                                             }) on contact.ContactId equals callH.ContactId
                              join manager in db.TR_User on contact.UserId equals manager.UserId
                              join custP in db.TT_CustomerProject on contact.CustProId equals custP.CustProId
                              where UserId.Contains((int)contact.UserId) && contact.ContactStatus == 4
                              select new HistoryCall()
                              {
                                  ContactName = contact.ContactName,
                                  ContactId = contact.ContactId,
                                  CustomerName = contact.CustomerName,
                                  CustProName = custP.CustProName,
                                  CustomerContactId = contact.CustomerContactId,
                                  CallStatus = contact.CallStatus,
                                  SubStatus = contact.SubStatus,
                                  BeginCall = callH.BeginCall,
                                  CallBack = callH.CallBack,
                                  AgingData = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), DateTime.Now),
                                  Reach = (from h in db.TT_CallHistory
                                           where h.ContactId == contact.ContactId
                                           select h).Count()
                              };

                var active = (from contact in db.TR_Contact.Where(p => p.CallStatus != "Closing")
                              join callH in (from call in db.TT_CallHistory
                                             group call by call.ContactId into grouping
                                             select new
                                             {
                                                 ContactId = grouping.Key,
                                                 BeginCall = grouping.Min(prod => prod.CallDate),
                                                 CallBack = grouping.Max(prod => prod.CallDate)
                                             }) on contact.ContactId equals callH.ContactId
                              join manager in db.TR_User on contact.UserId equals manager.UserId
                              join custP in db.TT_CustomerProject on contact.CustProId equals custP.CustProId
                              join cust in db.TR_Customer on custP.CustomerId equals cust.CustomerId
                              select new HistoryCall()
                              {
                                  ContactName = contact.ContactName,
                                  ContactId = contact.ContactId,
                                  CustomerName = contact.CustomerName,
                                  CustProName = custP.CustProName,
                                  CustomerContactId = contact.CustomerContactId,
                                  CallStatus = contact.CallStatus,
                                  SubStatus = contact.SubStatus,
                                  BeginCall = callH.BeginCall,
                                  CallBack = callH.CallBack,
                                  AgingAgent = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), DateTime.Now),
                                  AgingData = SqlFunctions.DateDiff("d", contact.CreatedOn, DateTime.Now),
                                  Reach = (from h in db.TT_CallHistory
                                           where h.ContactId == contact.ContactId
                                           select h).Count()
                              });

                ViewBag.closing = closing.ToList();
                ViewBag.suspect = suspect.ToList();
                ViewBag.active = active.ToList();
                return View();
            } else {
                return RedirectToAction("Login", "Auth");
            }
        }


        public ActionResult report(int? CustProId, String Periode)
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "2")
            {
                String UserName = Session["UserName"].ToString();
                int UserIds = Convert.ToInt32(Session["UserId"].ToString());

                int[] UserId = (from user in db.TR_User
                                where user.UserManagerId == UserIds
                                select user.UserId).ToArray();
                ViewBag.CustProId = (from uPro in db.TT_UserProject
                                     join custP in db.TT_CustomerProject on uPro.CustProId equals custP.CustProId
                                     where UserId.Contains((int)uPro.UserId)
                                     && custP.status == 1
                                     select new CustomerProject
                                     {
                                         CustProId = custP.CustProId,
                                         CustProName = custP.CustProName
                                     }).Distinct().ToList().OrderBy(p => p.CustomerName);
                return View();
            } else {
                return RedirectToAction("Login", "Auth");
            }
        }

        public ActionResult changePassword(int? UserId, String UserPass)
        {
            String password = Helper.EncodePassword(UserPass, "th1siScRmc0nT4Ctc3nTeR!!!");
            TR_User tabelPass = db.TR_User.Find(UserId);
            if (tabelPass != null)
            {
                var dtPass = db.TR_User.Where(p => p.UserId == UserId).FirstOrDefault();
                dtPass.UserPass = password;
                db.SaveChanges();
            }
            return RedirectToAction("Logout", "Auth");
        }
        [HttpPost]
        public JsonResult targetSet()
        {
            NameValueCollection nvc = Request.Form;
            String Periode = Convert.ToString(nvc["Periode"]);
            int CustProId = Convert.ToInt32(nvc["CustProId"]);
            int target = 0;

            var result = new List<contactCenterModels.Performance>();
            var productivity = new List<contactCenterModels.Productivity>();
            DateTime dateNow = DateTime.Now;
            String StartDate = DateTime.Now.ToString("yyyy-MM-dd"), EndDate = StartDate;
            if (Periode.Equals("Daily"))
            {
                EndDate = dateNow.AddDays(1).ToString("yyyy-MM-dd");
            }
            else if (Periode.Equals("Weekly"))
            {
                EndDate = dateNow.AddDays(7).ToString("yyyy-MM-dd");
            }
            else if (Periode.Equals("Monthly"))
            {
                EndDate = dateNow.AddDays(30).ToString("yyyy-MM-dd");
            }
            //Defined SQL CONNECTION
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;

            using (SqlConnection dbconnection = connection)
            {
                dbconnection.Open();
                
                //membuka koneksi database

                //mengambil data resource dari tabel TR_User
                String query = @"( select a.TargetData/b.UserId as Jumlah 
                                    from TR_TargetSetting a, 
                            (
                                                select count(UserId) as UserId
                                                from TT_UserProject
                                                where CustProId = @CustProId
                            )b
                                    where a.CustProId = @CustProId )";
                using (SqlCommand cmd = new SqlCommand(query, dbconnection))
                {
                    cmd.Parameters.AddWithValue("@CustProId", CustProId);
                    target = Convert.ToInt32(cmd.ExecuteScalar());
                }

                query = "SELECT a.UserId, b.UserName FROM TT_UserProject a join TR_User b on a.UserId = b.UserId WHERE CustProId = " + CustProId + "AND UserManagerId = '" + Session["UserId"] + "'";
                using (SqlDataAdapter cmd = new SqlDataAdapter(query, dbconnection))
                {
                    DataSet User = new DataSet("UserId");
                    cmd.FillSchema(User, SchemaType.Source, "TT_UserProject");
                    cmd.Fill(User, "TT_UserProject");

                    DataTable tblUser;
                    tblUser = User.Tables["TT_UserProject"];

                    for (int i = 0; i < tblUser.Rows.Count; i++)
                    {
                        string query2 = @" SELECT a.UserId ,
                                  SUM(CONVERT(INT, jamDuration)) as Jam ,
                                  SUM(CONVERT(INT, MenitDuration)) as Menit ,
                                  SUM(CONVERT(INT, DetikDuration)) as Detik ,
                                  COUNT(a.UserId) As Attempt,
                                  b.jumlah as Utillization
                                FROM(
                                   SELECT UserId,
                                     ContactId,
                                     --COUNT(ContactId) as Utillization,
                                     SUBSTRING(CallDuration, 1, 2)  AS jamDuration,
                                     SUBSTRING(CallDuration, 4, 2) AS MenitDuration,
                                     SUBSTRING(CallDuration, 7, 2)  AS DetikDuration
                                   FROM TT_CallHistory

                                  Where CallDate between @StartDate  AND @EndDate

                                  ) a join
                                  (
                                    SELECT UserId, sum(1) as jumlah
                                    FROM(
                                    SELECT UserId, contactId
                                    FROM TT_CallHistory
                                    GROUP BY UserId, ContactId
                                    )a
                                GROUP BY UserId
                                ) b on a.UserId = b.UserId
                                Where a.UserId = @User
                            GROUP BY a.UserId, b.jumlah";

                        using (SqlCommand command = new SqlCommand(query2, dbconnection))
                        {
                            var data = new contactCenterModels.Productivity();

                            command.Parameters.AddWithValue("@User", tblUser.Rows[i][0].ToString());
                            command.Parameters.AddWithValue("@StartDate", StartDate);
                            command.Parameters.AddWithValue("@EndDate", EndDate);
                            SqlDataReader reader = command.ExecuteReader();

                            if (!reader.HasRows)
                            {
                                data.Periode = Periode;
                                data.UserName = tblUser.Rows[i][1].ToString();
                                data.callAtemp = 0;
                                data.Utillization = 0;
                                data.talkTime = "00:00:00";
                                productivity.Add(data);
                                //Console.WriteLine((i + 1) + ". | " + tblUser.Rows[i][0].ToString() + " | " + tblUser.Rows[i][1].ToString() + "| 0% | " + target + "| 0 | 0 | 0 | 0 | 0 | 0 | 0");
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    data.Periode = Periode;
                                    data.UserName = tblUser.Rows[i][1].ToString();
                                    data.callAtemp = Convert.ToInt32(reader["Attempt"].ToString());
                                    data.Utillization = Convert.ToInt32(reader["Utillization"].ToString());

                                    int Detik = Convert.ToInt32(reader["Detik"].ToString());
                                    int Menit = Convert.ToInt32(reader["Menit"].ToString());
                                    int Jam = Convert.ToInt32(reader["Jam"].ToString());

                                    Menit = Menit + (Detik / 60);
                                    Detik = Detik % 60;
                                    Jam = Jam + (Menit / 60);
                                    Menit = Menit % 60;

                                    data.talkTime = Jam + ":" + Menit + ":" + Detik;
                                    productivity.Add(data);
                                    //Console.WriteLine((i + 1) + ". | " + reader["UserId"].ToString() + " | " + reader["UserName"].ToString() + " | " + reader["Achievement"].ToString() + "% | " + target + " | " + reader["Closing"].ToString() + " | " + reader["Prospect"].ToString() + " | " + reader["Promising"].ToString() + " | " + reader["Contacted"].ToString() + " | " + reader["Connected"].ToString() + " | " + reader["NotConnected"].ToString() + " | " + reader["Total"].ToString());
                                }
                            }

                            reader.Close();
                        }


                        String query1 = @"Select a.UserId ,
                                            b.UserName ,
                                            Closing*100/@Target as Achievement ,
                                            @Target as Target ,
                                            Closing ,
                                            Prospect ,
                                            Promising ,
                                            Contacted ,
                                            Connected ,
                                            NotConnected ,
                                            Total = Closing + Prospect + Promising + Contacted + Connected + NotConnected
                                            from 
                                            (
                                            Select UserId ,
                                                Sum(Case when CallStatus='Closing' then 1 else 0 END) as Closing ,
                                                Sum(Case when CallStatus='Prospect' then 1 else 0 END) as Prospect ,
                                                Sum(Case when CallStatus='Promising' then 1 else 0 END) as Promising ,
                                                Sum(Case when CallStatus='Contacted' then 1 else 0 END) as Contacted ,
                                                Sum(Case when CallStatus='Connected' then 1 else 0 END) as Connected ,
                                                Sum(Case when CallStatus='Not Connected' then 1 else 0 END) as NotConnected
                                            from ( 
                                            select UserId ,CallStatus
                                            from TT_CallHistory callh join
                                                (
                                                select ContactId, max(CallDate) as CallDate
                                                from TT_CallHistory
                                                group by ContactId
                                                )a
                                            on callh.CallDate=a.CallDate and callh.ContactId=a.ContactId
                                                        where callh.CallDate between @StartDate and @EndDate
                                            )a
                                                group by a.UserId
                                            )a join TR_User b on a.UserId=b.UserId
                                            WHERE a.UserId=@User";

                        using (SqlCommand command = new SqlCommand(query1, dbconnection))
                        {
                            var data = new contactCenterModels.Performance();

                            command.Parameters.AddWithValue("@User", tblUser.Rows[i][0].ToString());
                            command.Parameters.AddWithValue("@Target", target);
                            command.Parameters.AddWithValue("@StartDate", StartDate);
                            command.Parameters.AddWithValue("@EndDate", EndDate);
                            SqlDataReader reader = command.ExecuteReader();

                            if (!reader.HasRows)
                            {
                                data.Periode = Periode;
                                data.UserName = tblUser.Rows[i][1].ToString();
                                data.Achievment = "0";
                                data.Target = target;
                                data.Closing = 0;
                                data.Prospect = 0;
                                data.Contacted = 0;
                                data.Connected = 0;
                                data.NotConnected = 0;
                                result.Add(data);
                                //Console.WriteLine((i + 1) + ". | " + tblUser.Rows[i][0].ToString() + " | " + tblUser.Rows[i][1].ToString() + "| 0% | " + target + "| 0 | 0 | 0 | 0 | 0 | 0 | 0");
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    data.Periode = Periode;
                                    data.UserName = reader["UserName"].ToString();
                                    //data.Achievment = Convert.ToInt32(reader["Achievement"].ToString());
                                    data.Target = target;
                                    data.Closing = Convert.ToInt32(reader["Closing"].ToString());
                                    data.Prospect = Convert.ToInt32(reader["Prospect"].ToString());
                                    data.Contacted = Convert.ToInt32(reader["Contacted"].ToString());
                                    data.Connected = Convert.ToInt32(reader["Connected"].ToString());
                                    data.NotConnected = Convert.ToInt32(reader["NotConnected"].ToString()); ;
                                    result.Add(data);
                                    //Console.WriteLine((i + 1) + ". | " + reader["UserId"].ToString() + " | " + reader["UserName"].ToString() + " | " + reader["Achievement"].ToString() + "% | " + target + " | " + reader["Closing"].ToString() + " | " + reader["Prospect"].ToString() + " | " + reader["Promising"].ToString() + " | " + reader["Contacted"].ToString() + " | " + reader["Connected"].ToString() + " | " + reader["NotConnected"].ToString() + " | " + reader["Total"].ToString());
                                }
                            }

                            reader.Close();
                        }
                    }
                }
            }
            var hasil = new
            {
                Performance = result.ToList(),
                Productivity = productivity.ToList()
            };

            return Json(hasil, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult targetSet2(string From, string To, string ProjectId)
        {
            NameValueCollection nvc = Request.Form;
            DateTime dtFrom = Convert.ToDateTime(nvc["DateFrom"]);
            DateTime dtTo = Convert.ToDateTime(nvc["DateTo"]);
            int CustProId = Convert.ToInt32(nvc["CustProId"]);
            
            int target = 0;

            var result = new List<contactCenterModels.Performance>();
            var productivity = new List<contactCenterModels.Productivity>();
            DateTime dateNow = DateTime.Now;
            String StartDate = dtFrom.ToString("yyyy-MM-dd");
            String EndDate = dtTo.ToString("yyyy-MM-dd");

            //Defined SQL CONNECTION
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;

            using (SqlConnection dbconnection = connection)
            {
                dbconnection.Open();

                //membuka koneksi database

                //mengambil data resource dari tabel TR_User
                String query = @"( select a.TargetData/b.UserId as Jumlah 
                                    from TR_TargetSetting a, 
                            (
                                                select count(UserId) as UserId
                                                from TT_UserProject
                                                where CustProId = @CustProId
                            )b
                                    where a.CustProId = @CustProId )";
                using (SqlCommand cmd = new SqlCommand(query, dbconnection))
                {
                    cmd.Parameters.AddWithValue("@CustProId", CustProId);
                    target = Convert.ToInt32(cmd.ExecuteScalar());
                }

                query = "SELECT a.UserId, b.UserName, b.UserSkill FROM TT_UserProject a join TR_User b on a.UserId = b.UserId WHERE CustProId = " + CustProId + "AND UserManagerId = '" + Session["UserId"] + "'";
                using (SqlDataAdapter cmd = new SqlDataAdapter(query, dbconnection))
                {
                    DataSet User = new DataSet("UserId");
                    cmd.FillSchema(User, SchemaType.Source, "TT_UserProject");
                    cmd.Fill(User, "TT_UserProject");

                    DataTable tblUser;
                    tblUser = User.Tables["TT_UserProject"];

                    for (int i = 0; i < tblUser.Rows.Count; i++)
                    {
                        string query2 = @" SELECT a.UserId ,
                                  SUM(CONVERT(INT, jamDuration)) as Jam ,
                                  SUM(CONVERT(INT, MenitDuration)) as Menit ,
                                  SUM(CONVERT(INT, DetikDuration)) as Detik ,
                                  COUNT(a.UserId) As Attempt,
                                  b.jumlah as Utillization
                                FROM(
                                   SELECT TT_CallHistory.UserId,
                                     TT_CallHistory.ContactId,
                                     --COUNT(ContactId) as Utillization,
                                     SUBSTRING(CallDuration, 1, 2)  AS jamDuration,
                                     SUBSTRING(CallDuration, 4, 2) AS MenitDuration,
                                     SUBSTRING(CallDuration, 7, 2)  AS DetikDuration
                                   FROM TT_CallHistory
                                    inner join TR_CONTACT ct
										on ct.ContactId = TT_CallHistory.ContactId

                                  where ct.CustProId = @CustProId and CallDate between @StartDate  AND @EndDate

                                  ) a join
                                  (
                                    SELECT UserId, sum(1) as jumlah
                                    FROM(
                                        SELECT TT_CallHistory.UserId, TT_CallHistory.contactId
                                        FROM TT_CallHistory
                                            inner join TR_CONTACT ct
										        on ct.ContactId = TT_CallHistory.ContactId
                                        where ct.CustProId = @CustProId and CallDate between @StartDate  AND @EndDate
                                        GROUP BY TT_CallHistory.UserId, TT_CallHistory.ContactId
                                    )a
                                GROUP BY UserId
                                ) b on a.UserId = b.UserId
                                Where a.UserId = @User
                            GROUP BY a.UserId, b.jumlah";

                        using (SqlCommand command = new SqlCommand(query2, dbconnection))
                        {
                            var data = new contactCenterModels.Productivity();

                            command.Parameters.AddWithValue("@User", tblUser.Rows[i][0].ToString());
                            command.Parameters.AddWithValue("@CustProId", CustProId);
                            command.Parameters.AddWithValue("@StartDate", StartDate);
                            command.Parameters.AddWithValue("@EndDate", EndDate);
                            SqlDataReader reader = command.ExecuteReader();

                            if (!reader.HasRows)
                            {
                                //data.Periode = Periode;
                                data.UserName = tblUser.Rows[i][1].ToString();
                                data.callAtemp = 0;
                                data.Utillization = 0;
                                data.talkTime = "00:00:00";
                                productivity.Add(data);
                                //Console.WriteLine((i + 1) + ". | " + tblUser.Rows[i][0].ToString() + " | " + tblUser.Rows[i][1].ToString() + "| 0% | " + target + "| 0 | 0 | 0 | 0 | 0 | 0 | 0");
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    //data.Periode = Periode;
                                    data.UserName = tblUser.Rows[i][1].ToString();
                                    data.callAtemp = Convert.ToInt32(reader["Attempt"].ToString());
                                    data.Utillization = Convert.ToInt32(reader["Utillization"].ToString());
                                   
                                    int Detik = Convert.ToInt32(reader["Detik"].ToString());
                                    int Menit = Convert.ToInt32(reader["Menit"].ToString());
                                    int Jam = Convert.ToInt32(reader["Jam"].ToString());

                                    Menit = Menit + (Detik / 60);
                                    Detik = Detik % 60;
                                    Jam = Jam + (Menit / 60);
                                    Menit = Menit % 60;

                                    data.talkTime = Jam + ":" + Menit + ":" + Detik;
                                    productivity.Add(data);
                                    //Console.WriteLine((i + 1) + ". | " + reader["UserId"].ToString() + " | " + reader["UserName"].ToString() + " | " + reader["Achievement"].ToString() + "% | " + target + " | " + reader["Closing"].ToString() + " | " + reader["Prospect"].ToString() + " | " + reader["Promising"].ToString() + " | " + reader["Contacted"].ToString() + " | " + reader["Connected"].ToString() + " | " + reader["NotConnected"].ToString() + " | " + reader["Total"].ToString());
                                }
                            }

                            reader.Close();
                        }


                        String query1 = @"Select a.UserId ,
                                            b.UserName ,
                                            b.UserSkill,
                                            --Closing*100/Convert(float, @Target) as Achievement ,

                                            Closing/
											CONVERT(FLOAT,
											@Target * (
											(CASE(UserSkill)
											when 'Junior' then ts.Beginner
											when 'Intermediate' then ts.Intermediate
											else ts.Advance
											end)/Convert(Float,100)
											)) * 100 Achievement,

                                            @Target as Target ,
                                            Closing ,
                                            Prospect ,
                                            Promising ,
                                            Contacted ,
                                            Connected ,
                                            NotConnected ,
                                            Total = Closing + Prospect + Promising + Contacted + Connected + NotConnected
                                            from 
                                            (
                                            Select UserId ,CustProId,
                                                Sum(Case when CallStatus='Closing' then 1 else 0 END) as Closing ,
                                                Sum(Case when CallStatus='Prospect' then 1 else 0 END) as Prospect ,
                                                Sum(Case when CallStatus='Promising' then 1 else 0 END) as Promising ,
                                                Sum(Case when CallStatus='Contacted' then 1 else 0 END) as Contacted ,
                                                Sum(Case when CallStatus='Connected' then 1 else 0 END) as Connected ,
                                                Sum(Case when CallStatus='Not Connected' then 1 else 0 END) as NotConnected
                                            from ( 
                                            select callh.UserId ,callh.CallStatus, ct.CustProId
                                            from TT_CallHistory callh join
                                                (
                                                select ContactId, max(CallDate) as CallDate
                                                from TT_CallHistory
                                                group by ContactId
                                                )a
                                            on callh.CallDate=a.CallDate and callh.ContactId=a.ContactId
                                            inner join TR_CONTACT ct
                                            on ct.ContactId = callh.ContactId
                                                        where ct.CustProId = @CustProId and callh.CallDate between @StartDate and @EndDate
                                            )a
                                                group by a.UserId, CustProId
                                            )a join TR_User b on a.UserId=b.UserId
                                            
                                            inner join TR_TargetSetting ts
												on ts.CustProId = a.CustProId
												and ts.TargetFrom < GETDATE() and ts.TargetTo >= GETDATE()
                                            WHERE a.UserId=@User";

                        using (SqlCommand command = new SqlCommand(query1, dbconnection))
                        {
                            var data = new contactCenterModels.Performance();

                            command.Parameters.AddWithValue("@User", tblUser.Rows[i][0].ToString());
                            command.Parameters.AddWithValue("@Target", target);
                            command.Parameters.AddWithValue("@CustProId", CustProId);
                            command.Parameters.AddWithValue("@StartDate", StartDate);
                            command.Parameters.AddWithValue("@EndDate", EndDate);
                            SqlDataReader reader = command.ExecuteReader();

                            if (!reader.HasRows)
                            {
                                //data.Periode = Periode;
                                data.UserName = tblUser.Rows[i][1].ToString();
                                //data.Achievment = 0;
                                data.Target = target;
                                data.Closing = 0;
                                data.Prospect = 0;
                                data.Contacted = 0;
                                data.Connected = 0;
                                data.NotConnected = 0;
                                data.Achievment = "0";
                                data.Image = "../Asset/images/poorly.png";
                                data.UserSkill = tblUser.Rows[i][2].ToString();
                                result.Add(data);
                                //Console.WriteLine((i + 1) + ". | " + tblUser.Rows[i][0].ToString() + " | " + tblUser.Rows[i][1].ToString() + "| 0% | " + target + "| 0 | 0 | 0 | 0 | 0 | 0 | 0");
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    //data.Periode = Periode;
                                    data.UserName = reader["UserName"].ToString();
                                    data.Achievment = String.Format("{0:0.##}", Convert.ToDouble(reader["Achievement"].ToString()));
                                    data.Target = target;
                                    data.Closing = Convert.ToInt32(reader["Closing"].ToString());
                                    data.Prospect = Convert.ToInt32(reader["Prospect"].ToString());
                                    data.Contacted = Convert.ToInt32(reader["Contacted"].ToString());
                                    data.Connected = Convert.ToInt32(reader["Connected"].ToString());
                                    data.NotConnected = Convert.ToInt32(reader["NotConnected"].ToString());
                                    data.UserSkill = reader["UserSkill"].ToString();
                                    data.Image = (Convert.ToDouble(reader["Achievement"].ToString())) > 95 ? "../Asset/images/very-good.png" : "../Asset/images/poorly.png";
                                    result.Add(data);
                                    //Console.WriteLine((i + 1) + ". | " + reader["UserId"].ToString() + " | " + reader["UserName"].ToString() + " | " + reader["Achievement"].ToString() + "% | " + target + " | " + reader["Closing"].ToString() + " | " + reader["Prospect"].ToString() + " | " + reader["Promising"].ToString() + " | " + reader["Contacted"].ToString() + " | " + reader["Connected"].ToString() + " | " + reader["NotConnected"].ToString() + " | " + reader["Total"].ToString());
                                }
                            }

                            reader.Close();
                        }
                    }
                }
            }
            var hasil = new
            {
                Performance = result.ToList(),
                Productivity = productivity.ToList()
            };



            return Json(hasil, JsonRequestBehavior.AllowGet);



        }

        [HttpPost]
        public string GetPerformance(string From, string To, string ProjectId)
        {
            List<Dictionary<String, String>> _list = new List<Dictionary<string, string>>();

            NameValueCollection nvc = Request.Form;
            DateTime dtFrom = Convert.ToDateTime(nvc["DateFrom"]);
            DateTime dtTo = Convert.ToDateTime(nvc["DateTo"]);
            int CustProId = Convert.ToInt32(nvc["CustProId"]);

            int target = 0;

            var result = new List<contactCenterModels.Performance>();
            DateTime dateNow = DateTime.Now;
            String StartDate = dtFrom.ToString("yyyy-MM-dd");
            String EndDate = dtTo.ToString("yyyy-MM-dd");

            //Defined SQL CONNECTION
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;

            using (SqlConnection dbconnection = connection)
            {
                dbconnection.Open();

                //membuka koneksi database

                //mengambil data resource dari tabel TR_User
                String query = @"( select a.TargetData/b.UserId as Jumlah 
                                    from TR_TargetSetting a, 
                            (
                                                select count(UserId) as UserId
                                                from TT_UserProject
                                                where CustProId = @CustProId
                            )b
                                    where a.CustProId = @CustProId )";
                using (SqlCommand cmd = new SqlCommand(query, dbconnection))
                {
                    cmd.Parameters.AddWithValue("@CustProId", CustProId);
                    target = Convert.ToInt32(cmd.ExecuteScalar());
                }

                query = "SELECT a.UserId, b.UserName, b.UserSkill FROM TT_UserProject a join TR_User b on a.UserId = b.UserId WHERE CustProId = " + CustProId + "AND UserManagerId = '" + Session["UserId"] + "'";
                using (SqlDataAdapter cmd = new SqlDataAdapter(query, dbconnection))
                {
                    DataSet User = new DataSet("UserId");
                    cmd.FillSchema(User, SchemaType.Source, "TT_UserProject");
                    cmd.Fill(User, "TT_UserProject");

                    DataTable tblUser;
                    tblUser = User.Tables["TT_UserProject"];

                    for (int i = 0; i < tblUser.Rows.Count; i++)
                    {
                        String query1 = @"Select a.UserId ,
                                            b.UserName ,
                                            b.UserSkill,
                                            --Closing*100/Convert(float, @Target) as Achievement ,

                                            Closing/
											CONVERT(FLOAT,
											@Target * (
											(CASE(UserSkill)
											when 'Junior' then ts.Beginner
											when 'Intermediate' then ts.Intermediate
											else ts.Advance
											end)/Convert(Float,100)
											)) * 100 Achievement,

                                            @Target as Target ,
                                            Closing ,
                                            Prospect ,
                                            Promising ,
                                            Contacted ,
                                            Connected ,
                                            NotConnected ,
                                            Total = Closing + Prospect + Promising + Contacted + Connected + NotConnected
                                            from 
                                            (
                                            Select UserId ,CustProId,
                                                Sum(Case when CallStatus='Closing' then 1 else 0 END) as Closing ,
                                                Sum(Case when CallStatus='Prospect' then 1 else 0 END) as Prospect ,
                                                Sum(Case when CallStatus='Promising' then 1 else 0 END) as Promising ,
                                                Sum(Case when CallStatus='Contacted' then 1 else 0 END) as Contacted ,
                                                Sum(Case when CallStatus='Connected' then 1 else 0 END) as Connected ,
                                                Sum(Case when CallStatus='Not Connected' then 1 else 0 END) as NotConnected
                                            from ( 
                                            select callh.UserId ,callh.CallStatus, ct.CustProId
                                            from TT_CallHistory callh join
                                                (
                                                select ContactId, max(CallDate) as CallDate
                                                from TT_CallHistory
                                                group by ContactId
                                                )a
                                            on callh.CallDate=a.CallDate and callh.ContactId=a.ContactId
                                            inner join TR_CONTACT ct
                                            on ct.ContactId = callh.ContactId
                                                        where ct.CustProId = @CustProId and callh.CallDate between @StartDate and @EndDate
                                            )a
                                                group by a.UserId, CustProId
                                            )a join TR_User b on a.UserId=b.UserId
                                            
                                            inner join TR_TargetSetting ts
												on ts.CustProId = a.CustProId
												and ts.TargetFrom < GETDATE() and ts.TargetTo >= GETDATE()
                                            WHERE a.UserId=@User";

                        using (SqlCommand command = new SqlCommand(query1, dbconnection))
                        {
                            var data = new contactCenterModels.Performance();

                            command.Parameters.AddWithValue("@User", tblUser.Rows[i][0].ToString());
                            command.Parameters.AddWithValue("@Target", target);
                            command.Parameters.AddWithValue("@CustProId", CustProId);
                            command.Parameters.AddWithValue("@StartDate", StartDate);
                            command.Parameters.AddWithValue("@EndDate", EndDate);
                            SqlDataReader reader = command.ExecuteReader();

                            if (!reader.HasRows)
                            {
                                //data.Periode = Periode;
                                data.UserName = tblUser.Rows[i][1].ToString();
                                //data.Achievment = 0;
                                data.Target = target;
                                data.Closing = 0;
                                data.Prospect = 0;
                                data.Contacted = 0;
                                data.Connected = 0;
                                data.NotConnected = 0;
                                data.Achievment = "0";
                                data.Image = "../Asset/images/poorly.png";
                                data.UserSkill = tblUser.Rows[i][2].ToString();
                                result.Add(data);
                                //Console.WriteLine((i + 1) + ". | " + tblUser.Rows[i][0].ToString() + " | " + tblUser.Rows[i][1].ToString() + "| 0% | " + target + "| 0 | 0 | 0 | 0 | 0 | 0 | 0");
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    //data.Periode = Periode;
                                    data.UserName = reader["UserName"].ToString();
                                    data.Achievment = String.Format("{0:0.##}", Convert.ToDouble(reader["Achievement"].ToString()));
                                    data.Target = target;
                                    data.Closing = Convert.ToInt32(reader["Closing"].ToString());
                                    data.Prospect = Convert.ToInt32(reader["Prospect"].ToString());
                                    data.Contacted = Convert.ToInt32(reader["Contacted"].ToString());
                                    data.Connected = Convert.ToInt32(reader["Connected"].ToString());
                                    data.NotConnected = Convert.ToInt32(reader["NotConnected"].ToString());
                                    data.UserSkill = reader["UserSkill"].ToString();
                                    data.Image = (Convert.ToDouble(reader["Achievement"].ToString())) > 95 ? "../Asset/images/very-good.png" : "../Asset/images/poorly.png";
                                    result.Add(data);
                                    //Console.WriteLine((i + 1) + ". | " + reader["UserId"].ToString() + " | " + reader["UserName"].ToString() + " | " + reader["Achievement"].ToString() + "% | " + target + " | " + reader["Closing"].ToString() + " | " + reader["Prospect"].ToString() + " | " + reader["Promising"].ToString() + " | " + reader["Contacted"].ToString() + " | " + reader["Connected"].ToString() + " | " + reader["NotConnected"].ToString() + " | " + reader["Total"].ToString());
                                }
                            }

                            reader.Close();
                        }
                    }
                }
            }

            foreach (var m in result)
            {
                Dictionary<String, String> dict = new Dictionary<string, string>();
                dict.Add("AgentName", m.UserName.ToString());
                dict.Add("Achievement", m.Achievment);
                dict.Add("Target", m.Target.ToString());
                dict.Add("Closing", m.Closing.ToString());
                dict.Add("Prospect", m.Prospect.ToString());
                dict.Add("Contacted", m.Contacted.ToString());
                dict.Add("Connected", m.Connected.ToString());
                dict.Add("NotConnected", m.NotConnected.ToString());
                dict.Add("Image", m.Image.ToString());
                _list.Add(dict);
            }
            
            return JsonConvert.SerializeObject(_list);
        }

        [HttpPost]
        public string GetProductivity(string From, string To, string ProjectId)
        {
            List<Dictionary<String, String>> _list = new List<Dictionary<string, string>>();

            NameValueCollection nvc = Request.Form;
            DateTime dtFrom = Convert.ToDateTime(nvc["DateFrom"]);
            DateTime dtTo = Convert.ToDateTime(nvc["DateTo"]);
            int CustProId = Convert.ToInt32(nvc["CustProId"]);

            int target = 0;

            var productivity = new List<contactCenterModels.Productivity>();
            DateTime dateNow = DateTime.Now;
            String StartDate = dtFrom.ToString("yyyy-MM-dd");
            String EndDate = dtTo.ToString("yyyy-MM-dd");

            //Defined SQL CONNECTION
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;

            using (SqlConnection dbconnection = connection)
            {
                dbconnection.Open();

                //membuka koneksi database

                //mengambil data resource dari tabel TR_User
                String query = @"( select a.TargetData/b.UserId as Jumlah 
                                    from TR_TargetSetting a, 
                            (
                                                select count(UserId) as UserId
                                                from TT_UserProject
                                                where CustProId = @CustProId
                            )b
                                    where a.CustProId = @CustProId )";
                using (SqlCommand cmd = new SqlCommand(query, dbconnection))
                {
                    cmd.Parameters.AddWithValue("@CustProId", CustProId);
                    target = Convert.ToInt32(cmd.ExecuteScalar());
                }

                query = "SELECT a.UserId, b.UserName, b.UserSkill FROM TT_UserProject a join TR_User b on a.UserId = b.UserId WHERE CustProId = " + CustProId + "AND UserManagerId = '" + Session["UserId"] + "'";
                using (SqlDataAdapter cmd = new SqlDataAdapter(query, dbconnection))
                {
                    DataSet User = new DataSet("UserId");
                    cmd.FillSchema(User, SchemaType.Source, "TT_UserProject");
                    cmd.Fill(User, "TT_UserProject");

                    DataTable tblUser;
                    tblUser = User.Tables["TT_UserProject"];

                    for (int i = 0; i < tblUser.Rows.Count; i++)
                    {
                        string query2 = @" SELECT a.UserId ,
                                  SUM(CONVERT(INT, jamDuration)) as Jam ,
                                  SUM(CONVERT(INT, MenitDuration)) as Menit ,
                                  SUM(CONVERT(INT, DetikDuration)) as Detik ,
                                  COUNT(a.UserId) As Attempt,
                                  b.jumlah as Utillization
                                FROM(
                                   SELECT TT_CallHistory.UserId,
                                     TT_CallHistory.ContactId,
                                     --COUNT(ContactId) as Utillization,
                                     SUBSTRING(CallDuration, 1, 2)  AS jamDuration,
                                     SUBSTRING(CallDuration, 4, 2) AS MenitDuration,
                                     SUBSTRING(CallDuration, 7, 2)  AS DetikDuration
                                   FROM TT_CallHistory
                                    inner join TR_CONTACT ct
										on ct.ContactId = TT_CallHistory.ContactId

                                  where ct.CustProId = @CustProId and CallDate between @StartDate  AND @EndDate

                                  ) a join
                                  (
                                    SELECT UserId, sum(1) as jumlah
                                    FROM(
                                        SELECT TT_CallHistory.UserId, TT_CallHistory.contactId
                                        FROM TT_CallHistory
                                            inner join TR_CONTACT ct
										        on ct.ContactId = TT_CallHistory.ContactId
                                        where ct.CustProId = @CustProId and CallDate between @StartDate  AND @EndDate
                                        GROUP BY TT_CallHistory.UserId, TT_CallHistory.ContactId
                                    )a
                                GROUP BY UserId
                                ) b on a.UserId = b.UserId
                                Where a.UserId = @User
                            GROUP BY a.UserId, b.jumlah";

                        using (SqlCommand command = new SqlCommand(query2, dbconnection))
                        {
                            var data = new contactCenterModels.Productivity();

                            command.Parameters.AddWithValue("@User", tblUser.Rows[i][0].ToString());
                            command.Parameters.AddWithValue("@CustProId", CustProId);
                            command.Parameters.AddWithValue("@StartDate", StartDate);
                            command.Parameters.AddWithValue("@EndDate", EndDate);
                            SqlDataReader reader = command.ExecuteReader();

                            if (!reader.HasRows)
                            {
                                //data.Periode = Periode;
                                data.UserName = tblUser.Rows[i][1].ToString();
                                data.callAtemp = 0;
                                data.Utillization = 0;
                                data.talkTime = "00:00:00";
                                productivity.Add(data);
                                //Console.WriteLine((i + 1) + ". | " + tblUser.Rows[i][0].ToString() + " | " + tblUser.Rows[i][1].ToString() + "| 0% | " + target + "| 0 | 0 | 0 | 0 | 0 | 0 | 0");
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    //data.Periode = Periode;
                                    data.UserName = tblUser.Rows[i][1].ToString();
                                    data.callAtemp = Convert.ToInt32(reader["Attempt"].ToString());
                                    data.Utillization = Convert.ToInt32(reader["Utillization"].ToString());

                                    int Detik = Convert.ToInt32(reader["Detik"].ToString());
                                    int Menit = Convert.ToInt32(reader["Menit"].ToString());
                                    int Jam = Convert.ToInt32(reader["Jam"].ToString());

                                    Menit = Menit + (Detik / 60);
                                    Detik = Detik % 60;
                                    Jam = Jam + (Menit / 60);
                                    Menit = Menit % 60;

                                    data.talkTime = Jam + ":" + Menit + ":" + Detik;
                                    productivity.Add(data);
                                    //Console.WriteLine((i + 1) + ". | " + reader["UserId"].ToString() + " | " + reader["UserName"].ToString() + " | " + reader["Achievement"].ToString() + "% | " + target + " | " + reader["Closing"].ToString() + " | " + reader["Prospect"].ToString() + " | " + reader["Promising"].ToString() + " | " + reader["Contacted"].ToString() + " | " + reader["Connected"].ToString() + " | " + reader["NotConnected"].ToString() + " | " + reader["Total"].ToString());
                                }
                            }

                            reader.Close();
                        }
                    }
                }
            }

            foreach (var m in productivity)
            {
                Dictionary<String, String> dict = new Dictionary<string, string>();
                dict.Add("User", m.UserName.ToString());
                dict.Add("CallAttempt", m.callAtemp.ToString());
                dict.Add("Utillization", m.Utillization.ToString());
                dict.Add("TalkTime", m.talkTime.ToString());
                _list.Add(dict);
            }

            return JsonConvert.SerializeObject(_list);
        }


        public ActionResult MasterSettingTarget()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                ViewBag.CustomerProjectName = db.TT_CustomerProject.Where(x => x.status == 1).ToList();
                var result = (from settTarget in db.TR_TargetSetting
                              join custP in db.TT_CustomerProject on settTarget.CustProId equals custP.CustProId
                              select new contactCenterModels.MasterSettingTarget
                              {
                                  TargetId = settTarget.TargetId,
                                  TargetName = settTarget.TargetName,
                                  TargetFrom = settTarget.TargetFrom,
                                  TargetTo = settTarget.TargetTo,
                                  TargetData = settTarget.TargetData,
                                  TargetAmountPaid = settTarget.TargetAmount.ToString(),
                                  CustProName = custP.CustProName,
                                  Advance = settTarget.Advance,
                                  Beginner = settTarget.Beginner,
                                  Intermediate = settTarget.Intermediate
                              }).ToList();
                return View(result);
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        public ActionResult createSettingTarget(int CustProId, String TargetName, string TargetAmountPaid, DateTime? TargetFrom, int TargetData, DateTime? TargetTo, int Advance, int Intermediate, int Junior)
        {
            int TargetAmountPaids = Convert.ToInt32(TargetAmountPaid.Replace(",", ""));

            TR_TargetSetting Target = new TR_TargetSetting();
            Target.Advance = Advance;
            Target.Beginner = Junior;
            Target.CreatedOn = DateTime.Now;
            Target.Created_by = Convert.ToString(Session["UserId"]);
            Target.Intermediate = Intermediate;
            Target.TargetAmount = TargetAmountPaids;
            Target.TargetData = TargetData;
            Target.TargetFrom = TargetFrom;
            Target.TargetTo = TargetTo;
            Target.TargetName = TargetName;
            Target.CustProId = CustProId;
            db.TR_TargetSetting.Add(Target);
            db.SaveChanges();

            return RedirectToAction("MasterSettingTarget");
        }

        [HttpPost]
        public ActionResult editSettingTarget(int TargetId, int CustProId, String TargetName, string TargetAmountPaid, DateTime? TargetFrom, int TargetData, DateTime? TargetTo, int Advance, int Intermediate, int Junior)
        {
            int TargetAmountPaids = Convert.ToInt32(TargetAmountPaid.Replace(",", ""));

            TR_TargetSetting Target = db.TR_TargetSetting.Find(TargetId);
            Target.Advance = Advance;
            Target.Beginner = Junior;
            Target.CreatedOn = DateTime.Now;
            Target.Created_by = Convert.ToString(Session["UserId"]);
            Target.Intermediate = Intermediate;
            Target.TargetAmount = TargetAmountPaids;
            Target.TargetData = TargetData;
            Target.TargetFrom = TargetFrom;
            Target.TargetTo = TargetTo;
            Target.TargetName = TargetName;
            Target.CustProId = CustProId;
            db.SaveChanges();

            return RedirectToAction("MasterSettingTarget");
        }

        public ActionResult deleteSettingTarget(int Id)
        {
            TR_TargetSetting Target = db.TR_TargetSetting.Find(Id);
            db.TR_TargetSetting.Remove(Target);
            db.SaveChanges();

            return RedirectToAction("MasterSettingTarget");
        }
        public ActionResult callView(int id)
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "2")
            {
                IEnumerable<contactCenterModels.HistoryDetail>
                      historydetail = (from callH in db.TT_CallHistory
                                       join contact in db.TR_Contact on callH.ContactId equals contact.ContactId
                                       join agent in db.TR_User on callH.UserId equals agent.UserId
                                       where contact.ContactId.Equals(id)
                                       select new contactCenterModels.HistoryDetail()
                                       {
                                           CallDate = callH.CallDate,
                                           ContactPhone = contact.ContactPhone,
                                           CallStatus = callH.CallStatus,
                                           SubStatus = callH.SubStatus,
                                           Remarks = callH.Remarks,
                                           AgingAgent = SqlFunctions.DateDiff("d", SqlFunctions.DateAdd("d", -3, contact.ExpiredDate), callH.CallDate),//1,//,
                                           UserName = agent.UserName
                                       }).OrderByDescending(x => x.CallDate);
                TR_Contact call = db.TR_Contact.Where(p => p.ContactId.Equals(id)).FirstOrDefault();
                TT_CallHistory callhis = db.TT_CallHistory.Where(p => p.ContactId == (id)).OrderByDescending(x => x.CallDate).FirstOrDefault();

                TT_CustomerProject paramProject = db.TT_CustomerProject.Where(p => p.CustProId == call.CustProId).FirstOrDefault();
                String CustomerName = (from h in db.TR_Customer
                                       where h.CustomerId == paramProject.CustomerId
                                       select h.CustomerName).FirstOrDefault();
                ViewBag.call = call;
                ViewBag.callhis = callhis;
                ViewBag.historydetail = historydetail.ToList();
                ViewBag.param = paramProject;
                ViewBag.CustomerName = CustomerName + " - " + paramProject.CustProName;
                return View();
            } else {   
                return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        // GET: ExportDataContact
        public ActionResult ExportToExcelContact(int? id)
        {

            var data = new List<TR_Contact>();
            if (id == id)
            {
                data = db.TR_Contact.Where(x => x.ContactId == id).ToList();
            }

            foreach (var item in data)
            {
                item.ContactStatus = 1;
                item.UserId = 0;

                var history = db.TT_CallHistory.Where(x => x.ContactId == item.ContactId).ToList();
                foreach (var item1 in history)
                {
                    db.TT_CallHistory.Remove(item1);

                }
            }
            db.SaveChanges();
            String custProName = (from custPro in db.TT_CustomerProject
                                  where custPro.CustProId == custPro.CustProId
                                  select custPro.CustProName).FirstOrDefault();
            
            // set the data source
            GridView gridview = new GridView();
            gridview.DataSource = data;
            gridview.DataBind();

            // Clear all the content from the current response
            Response.ClearContent();
            Response.Buffer = true;

            // set the header
            Response.AddHeader("content-disposition", "attachment;filename = Withdraw " + custProName + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";

            // create HtmlTextWriter object with StringWriter
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    // render the GridView to the HtmlTextWriter
                    gridview.RenderControl(htw);
                    // Output the GridView content saved into StringWriter
                    Response.Output.Write(sw.ToString());
                    Response.Flush();
                    Response.End();
                }
            }

            return RedirectToAction("withdraw");
        }

        [HttpPost]
        // GET: ExportDataContact
        public int UpdateCallStatus(int? ContactId, int Status)
        {
            var data = new List<TR_Contact>();
            if (ContactId != null)
            {
                data = db.TR_Contact.Where(x => x.ContactId == ContactId).ToList();
            }

            if (Status == 1) //Withdraw
            {
                foreach (var item in data)
                {
                    item.ContactStatus = 6;
                    item.CallStatus = "Suspect";
                    var history = db.TT_CallHistory.Where(x => x.ContactId == item.ContactId).ToList();
                    foreach (var item1 in history)
                    {
                        db.TT_CallHistory.Remove(item1);
                    }
                }
                db.SaveChanges();
            }
            else {
                foreach (var item in data)
                {
                    item.ContactStatus = 0;
                    item.CallStatus = "New";
                    item.UserId = null;
                }
                db.SaveChanges();
            }

            return 1;
        }

        [HttpPost]
        public JsonResult selectFromTo()
        {
            //String status = "";
            //get data collumn number dari form upload
            NameValueCollection nvc = Request.Form;
            String CallDuration = Convert.ToString(nvc["CallDuration"]);
            int id = Convert.ToInt32(nvc["CustProId"]);

            String from = Convert.ToDateTime((from contact in db.TR_Contact
                                              where contact.CustProId == id
                                              select contact).Select(x => x.CreatedOn).Min()).ToString("dd-MMM-yyyy");
            String to = Convert.ToDateTime((from custP in db.TT_CustomerProject
                                            where custP.CustProId == id
                                            select custP).Select(x => x.CustProExpired).Single()).ToString("dd-MMM-yyyy");

            System.Diagnostics.Debug.WriteLine(from + ", " + to);

            var hasil = new
            {
                dateFrom = from,
                dateTo = to
            };

            return Json(hasil);
        }
        /* Master Expired*/
        /* Master Expired*/
        public ActionResult masterExpired()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "3")
            {
                ViewBag.CustomerName = db.TR_Customer.ToList();
                var result = (from CustP in db.TT_CustomerProject
                                join cust in db.TR_Customer on CustP.CustomerId equals cust.CustomerId
                                where CustP.status == 0
                                select new contactCenterModels.Customer()
                                {
                                    CustProId = CustP.CustProId,
                                    CustProName = CustP.CustProName,
                                    CustProExpired = CustP.CustProExpired,
                                    CustomerName = cust.CustomerName,
                                    Param1 = CustP.Param1,
                                    Param2 = CustP.Param2,
                                    Param3 = CustP.Param3,
                                    Param4 = CustP.Param4,
                                    Param5 = CustP.Param5
                                }).ToList();
                return View(result);
            }
            else {
                return RedirectToAction("Login", "Auth");
            }
        }

        public ActionResult withdrawExpired()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "3")
            {

                DateTime dtMaxExp = DateTime.Now.AddMonths(-6);
                ViewBag.GroupExp = new SelectList(db.TT_CustomerProject.Where(x => x.status == 0 && x.CustProExpired > dtMaxExp).OrderBy(p=> p.CustProName), "CustProId", "CustProName");
                return View();
            } else {
                return RedirectToAction("Login", "Auth");
            }
        }

        [HttpPost]
        // GET: ExportData
        public ActionResult ExpExportToExcel(int GroupExp, String callstatus, DateTime DateTo, DateTime DateFrom)
        {
            var data = new List<TR_Contact>();
            if (callstatus == "notselected")
            {
                data = db.TR_Contact.Where(x => x.CustProId == GroupExp && x.CreatedOn >= DateFrom && x.CreatedOn <= DateTo && x.ContactStatus != 6).ToList();
            }
            else
            {
                data = db.TR_Contact.Where(x => x.CustProId == GroupExp && x.CallStatus == callstatus && x.CreatedOn >= DateFrom && x.CreatedOn <= DateTo && x.ContactStatus != 6).ToList();
            }
            foreach (var item in data)
            {
                item.ContactStatus = 6;
                var history = db.TT_CallHistory.Where(x => x.ContactId == item.ContactId).ToList();
                foreach (var item1 in history)
                {
                    db.TT_CallHistory.Remove(item1);
                }
            }
            db.SaveChanges();
            String custProName = (from custPro in db.TT_CustomerProject
                                  where custPro.CustProId == GroupExp
                                  select custPro.CustProName).FirstOrDefault();
            // set the data source
            GridView gridview = new GridView();
            gridview.DataSource = data;
            gridview.DataBind();
            // Clear all the content from the current response
            Response.ClearContent();
            Response.Buffer = true;
            // set the header
            Response.AddHeader("content-disposition", "attachment;filename = Withdraw " + custProName + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            // create HtmlTextWriter object with StringWriter
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    // render the GridView to the HtmlTextWriter
                    gridview.RenderControl(htw);
                    // Output the GridView content saved into StringWriter
                    Response.Output.Write(sw.ToString());
                    Response.Flush();
                    Response.End();
                }
                return RedirectToAction("withdrawExpired");
            }

        }


        [HttpPost]
        public string UpdateContactPhone(int ContactId, int PhoneType, string PhoneNumber)
        {
            var dtContact = db.TR_Contact.Where(p => p.ContactId == ContactId).FirstOrDefault();

            if (PhoneType == 1)
            {
                dtContact.ContactPhone = PhoneNumber;
            }
            else if (PhoneType == 2) {
                dtContact.HomePhone = PhoneNumber;
            }
            else if (PhoneType == 3)
            {
                dtContact.OtherPhone = PhoneNumber;
            }

            db.SaveChanges();

            return "ok";
        }
    }
}