//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace MVC_CRUD.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class TR_User
    {
        public int UserId { get; set; }
        public string UserName { get; set; }
        public string Email { get; set; }
        public Nullable<int> UserStatus { get; set; }
        public Nullable<System.DateTime> CreatedOn { get; set; }
        public string CreatedBy { get; set; }
        public Nullable<int> RoleId { get; set; }
        public string UserPass { get; set; }
        public string UserSkill { get; set; }
        public Nullable<int> UserManagerId { get; set; }
        public Nullable<bool> IsVerified { get; set; }
    }
}
