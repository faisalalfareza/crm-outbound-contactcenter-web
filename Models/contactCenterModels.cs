using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVC_CRUD.Models
{
    public class contactCenterModels
    {
        public class HistoryCall
        {
            public int ContactId { set; get; }
            public String CallDurtion { set; get; }
            public string CustomerContactId { set; get; }
            public String ContactName { set; get; }
            public String CallStatus { set; get; }
            public String SubStatus { set; get; }
            public String CustomerName { set; get; }
            public String CustProName { set; get; }
            public DateTime? BeginCall { set; get; }
            public DateTime? CallBack { set; get; }
            public int? AgingAgent { set; get; }
            public int? AgingData { set; get; }
            public int Reach { set; get; }
        }
        public class HistoryDetail
        {
            public DateTime? CallDate { set; get; }
            public String ContactPhone { set; get; }
            public String CallDurtion { set; get; }
            public String CallStatus { set; get; }
            public String SubStatus { set; get; }
            public String Remarks { set; get; }
            public int? AgingAgent { set; get; }
            public String UserName { set; get; }
        }
        public class createAgent
        {
            public TR_User user { set; get; }
            public List<TT_UserProject> userProject { set; get; }
        }
        public class Customer
        {
            public int CustProId { get; set; }
            public string CustProName { get; set; }
            public Nullable<System.DateTime> CustProExpired { get; set; }
            public string Param1 { get; set; }
            public string Param2 { get; set; }
            public string Param3 { get; set; }
            public string Param4 { get; set; }
            public string Param5 { get; set; }
            public string CustomerName { get; set; }
        }

        public class User
        {
            public int UserId { get; set; }
            public string UserName { get; set; }
            public string Email { get; set; }
            public string Role { get; set; }
            public string Level { get; set; }
            public string Team { get; set; }
            public string Active { get; set; }
        }

        public class Performance
        {
        public int CustProId { get; set; }
        public int CustProName { get; set; }
        public String Periode { get; set; }
        public int? UserId { get; set; }
        public string UserName { get; set; }
        public string Achievment { get; set; }
        public int Target { get; set; }
        public int Closing { get; set; }
        public int Prospect { get; set; }
        public int Promising { get; set; }
        public int Contacted { get; set; }
        public int Connected { get; set; }
        public int NotConnected { get; set; }
        public string UserSkill { get; set; }
        public string Image { get; set; }
        }

        public class Productivity
        {
            public String Periode { get; set; }
            public string UserName { get; set; }
            public int callAtemp { get; set; }
            public int Utillization { get; set; }
            public String talkTime { get; set; }
         
        }

        public class CustomerProject
        {
            public int CustProId { get; set; }
            public string CustProName { get; set; }
            public Nullable<System.DateTime> CustProExpired { get; set; }
            public string CustomerName { get; set; }
        }

        public class MasterSettingTarget
        {
            public int TargetId { get; set; }
            public string TargetName { get; set; }
            public DateTime? TargetFrom { get; set; }
            public DateTime? TargetTo { get; set; }
            public int? TargetData { get; set; }
            public decimal? TargetAmountPaid { get; set; }
            public String CustProName { get; set; }
            public int? Advance { get; set; }
            public int? Beginner { get; set; }
            public int? Intermediate { get; set; }
            public int CustProId { get; set; }
        }

    }
}