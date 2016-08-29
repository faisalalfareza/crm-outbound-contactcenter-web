using MVC_CRUD.Models;
using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace crm.Controllers
{
    public class UploadController : Controller
    {
        DB_CRM_CCEntities db = new DB_CRM_CCEntities();

        [HttpPost]
        public JsonResult UploadHandler()
        {
            String status = "";
            int UserId = Convert.ToInt32(Session["UserId"]);

            //get data collumn number dari form upload
            NameValueCollection nvc = Request.Form;
            int collumnNumber = Convert.ToInt32(nvc["collumnNumber"]);
            int CustomerProId = Convert.ToInt32(nvc["GroupId"]);
            String FileLocation = String.Empty;

            int jmlRows = 0;
            DataSet ds = new DataSet();
            if (Request.Files[0].ContentLength > 0)
            {
                string fileExtension = Path.GetExtension(Request.Files[0].FileName);

                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    string fileLocation = Server.MapPath("~/Content/") + Request.Files[0].FileName;
                    FileLocation = fileLocation;
                    if (System.IO.File.Exists(fileLocation))
                    {

                        System.IO.File.Delete(fileLocation);
                    }
                    Request.Files[0].SaveAs(fileLocation);
                    string excelConnectionString = string.Empty;
                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                    fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    //connection String for xls file format.
                    if (fileExtension == ".xls")
                    {
                        excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                        fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    //connection String for xlsx file format.
                    else if (fileExtension == ".xlsx")
                    {
                        excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                        fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }

                    //Create Connection to Excel work book and add oledb namespace
                    OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                    excelConnection.Open();
                    DataTable dt = new DataTable();

                    dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dt == null)
                    {
                        return null;
                    }

                    String[] excelSheets = new String[dt.Rows.Count];
                    int t = 0;

                    //excel data saves in temp file here.
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheets[t] = row["TABLE_NAME"].ToString();
                        t++;
                    }
                    OleDbConnection excelConnection1 = new OleDbConnection(excelConnectionString);


                    string query = string.Format("Select * from [{0}]", excelSheets[0]);
                    using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection1))
                    {
                        dataAdapter.Fill(ds);
                    }
                    excelConnection.Close();
                }

                //mengambil jumlah row dari excel yang di upload
                jmlRows = ds.Tables[0].Rows.Count;

                //Validasi jumlah row upload
                if (collumnNumber == jmlRows)
                {
                    string conn = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
                    SqlConnection con = new SqlConnection(conn);
                    con.Open();

                    String customerName = "";
                    int customerId = 0;
                    String query = "select CustomerName, CustomerId from TR_Customer where CustomerId=(select CustomerId From TT_CustomerProject where CustProId=@CustProId)";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@CustProId", CustomerProId);
                        SqlDataReader dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            customerName = (String)dr["CustomerName"];
                            customerId = (int)dr["CustomerId"];
                        }
                    }

                    query = "select top 0 * into #" + CustomerProId + " from TR_Contact";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    //insert ke tempTable
                    for (int i = 0; i < jmlRows; i++)
                    {
                        DateTime? date1 = ds.Tables[0].Rows[i][5].ToString() == "" ? null : (DateTime?)Convert.ToDateTime(ds.Tables[0].Rows[i][5].ToString());
                        DateTime? date2 = ds.Tables[0].Rows[i][26].ToString() == "" ? null : (DateTime?)Convert.ToDateTime(ds.Tables[0].Rows[i][26].ToString());
                        DateTime? date3 = ds.Tables[0].Rows[i][27].ToString() == "" ? null : (DateTime?)Convert.ToDateTime(ds.Tables[0].Rows[i][27].ToString());
                        DateTime? date4 = ds.Tables[0].Rows[i][30].ToString() == "" ? null : (DateTime?)Convert.ToDateTime(ds.Tables[0].Rows[i][30].ToString());
                        //query = "Insert into #" + CustomerProId + " Values('" + ds.Tables[0].Rows[i][2].ToString() + "','" + ds.Tables[0].Rows[i][4].ToString() + "','0','" + ds.Tables[0].Rows[i][25].ToString() + "','" + ds.Tables[0].Rows[i][0].ToString() + "','" + ds.Tables[0].Rows[i][1].ToString() + "','" + customerId + "','" + ds.Tables[0].Rows[i][3].ToString() + "','" + customerName + "','" + Convert.ToDateTime(ds.Tables[0].Rows[i][5].ToString()).ToString("MM/dd/yyyy") + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][6].ToString().Substring(0, 2)) + "','" + ds.Tables[0].Rows[i][7].ToString() + "','" + ds.Tables[0].Rows[i][8].ToString() + "','" + ds.Tables[0].Rows[i][9].ToString() + "','" + ds.Tables[0].Rows[i][10].ToString() + "','" + ds.Tables[0].Rows[i][11].ToString() + "','" + ds.Tables[0].Rows[i][12].ToString() + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][13].ToString()) + "','" + ds.Tables[0].Rows[i][14].ToString() + "','" + ds.Tables[0].Rows[i][15].ToString() + "','" + ds.Tables[0].Rows[i][16].ToString() + "','" + ds.Tables[0].Rows[i][17].ToString() + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][18].ToString()) + "','" + ds.Tables[0].Rows[i][19].ToString() + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][20].ToString()) + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][21].ToString()) + "','" + ds.Tables[0].Rows[i][22].ToString() + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][23].ToString()) + "','" + ds.Tables[0].Rows[i][24].ToString() + "', '" + Convert.ToDateTime(ds.Tables[0].Rows[i][26].ToString()).ToString("MM/dd/yyyy") + "', '" + Convert.ToDateTime(ds.Tables[0].Rows[i][27].ToString()).ToString("MM/dd/yyyy") + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][28].ToString()) + "','" + Convert.ToInt64(ds.Tables[0].Rows[i][29].ToString()) + "','" + ds.TAbConvert.ToDateTime(ds.Tables[0].Rows[i][30].ToString()).ToString("MM/dd/yyyy") + "','" + ds.Tables[0].Rows[i][31].ToString() + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '" + DateTime.Now.ToString("MM/dd/yyyy") + "', " + CustomerProId + ", '" + DateTime.Now.AddDays(3).ToString("MM/dd/yyyy") + "', 0)";

                        query = "Insert into #" + CustomerProId + " Values('" + ds.Tables[0].Rows[i][2].ToString() + "','" + ds.Tables[0].Rows[i][4].ToString() + "','0','" + ds.Tables[0].Rows[i][25].ToString() + "','" + ds.Tables[0].Rows[i][0].ToString() + "','" + ds.Tables[0].Rows[i][1].ToString() + "','" + customerId + "','" + ds.Tables[0].Rows[i][3].ToString() + "','" + customerName + "','" + (!date1.HasValue ? null : date1.Value.ToString("MM/dd/yyyy")) + "','" + (ds.Tables[0].Rows[i][6].ToString().Length > 1 ? Convert.ToInt64(ds.Tables[0].Rows[i][6].ToString().Substring(0, 2)) : 0)
                            + "','" + ds.Tables[0].Rows[i][7].ToString() + "','" + ds.Tables[0].Rows[i][8].ToString() + "','" + ds.Tables[0].Rows[i][9].ToString() + "','" + ds.Tables[0].Rows[i][10].ToString() + "','" + ds.Tables[0].Rows[i][11].ToString() + "','" + ds.Tables[0].Rows[i][12].ToString() + "','" + (ds.Tables[0].Rows[i][13] == DBNull.Value ? 0 : Convert.ToInt64(ds.Tables[0].Rows[i][13].ToString())) + "','" + ds.Tables[0].Rows[i][14].ToString() + "','" + ds.Tables[0].Rows[i][15].ToString() + "','" + ds.Tables[0].Rows[i][16].ToString() + "','" + ds.Tables[0].Rows[i][17].ToString() + "','" + (ds.Tables[0].Rows[i][18] == DBNull.Value ? 0 : Convert.ToInt64(ds.Tables[0].Rows[i][18].ToString())) + "','" + ds.Tables[0].Rows[i][19].ToString()
                            + "','" + (ds.Tables[0].Rows[i][20] == DBNull.Value ? 0 : Convert.ToInt64(ds.Tables[0].Rows[i][20].ToString())) + "','" + (ds.Tables[0].Rows[i][21] == DBNull.Value ? 0 : Convert.ToInt64(ds.Tables[0].Rows[i][21].ToString())) + "','" + ds.Tables[0].Rows[i][22].ToString() + "','" + (ds.Tables[0].Rows[i][23] == DBNull.Value ? 0 : Convert.ToInt64(ds.Tables[0].Rows[i][23].ToString())) + "','" + ds.Tables[0].Rows[i][24].ToString() + "', '" + (!date2.HasValue ? null : date2.Value.ToString("MM/dd/yyyy")) + "', '" + (!date3.HasValue ? null : date3.Value.ToString("MM/dd/yyyy")) + "','" + (ds.Tables[0].Rows[i][28] == DBNull.Value ? 0 : Convert.ToInt64(ds.Tables[0].Rows[i][28].ToString())) + "','" + (ds.Tables[0].Rows[i][29] == DBNull.Value ? 0 : Convert.ToInt64(ds.Tables[0].Rows[i][29].ToString())) + "','" + (!date4.HasValue ? null : date4.Value.ToString("MM/dd/yyyy"))
                            + "','" + ds.Tables[0].Rows[i][31].ToString() + "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "', " + CustomerProId + ", '" + DateTime.Now.AddDays(3).ToString("MM/dd/yyyy") + "', 0, '" + ds.Tables[0].Rows[i][32].ToString() + "', '-', '" + UserId + "')";

                        using (SqlCommand cmd = new SqlCommand(query, con))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }

                    //Duplicate Check
                    query = "MERGE TR_Contact t USING #" + CustomerProId + " s ON t.ContactPhone = s.ContactPhone AND t.CustProId = s.CustProId " +
                            " WHEN NOT MATCHED THEN Insert Values(s.CustomerContactId,s.ContactName,s.ContactStatus " +

                            " ,s.ContactPhone,s.BranchId,s.BranchFullname,s.CustomerId,s.AgreementNo,s.CustomerName " +
                            " ,(IIF(s.BirthDate = '1900-01-01', null, s.BirthDate)),s.Usia,s.KTP_Adress,s.KTP_RT,s.KTP_RW,s.KTP_Kelurahan,s.KTP_Kecamatan " +
                            " ,s.KTP_City,s.KTP_Zipcode,s.Residence_Adress,s.Residence_Kelurahan,s.Residence_Kecamatan " +
                            " ,s.Residence_City,s.Residence_Zipcode,s.Jenis_Pekerjaan,s.MonthlyFixedIncome,s.InstallmentAmount,s.Status_Rumah, " +
                            " s.JumlahTanggungan,s.DownPayment, (IIF(s.GoLiveDate = '1900-01-01', null, s.GoLiveDate)),(IIF(s.TglSelesaiAngsuran = '1900-01-01', null, s.TglSelesaiAngsuran)),s.Tenor " +
                            " ,s.Odmax_Day_Final,(IIF(s.LastPayment = '1900-01-01', null, s.LastPayment)), " +
                            " s.Payment, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, s.CreatedOn, s.CreatedBy, s.ExpiredDate, 0, s.HomePhone, s.OtherPhone, s.CustProId); ";

                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        jmlRows = cmd.ExecuteNonQuery();
                    }

                    //Menghapus tempTable
                    query = "Drop table #" + CustomerProId;
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();

                    status = jmlRows + " Rows Inserted";
                }
                else
                {
                    status = "Jumlah tidak sesuai";
                    if (System.IO.File.Exists(FileLocation))
                    {

                        System.IO.File.Delete(FileLocation);
                    }
                }
            }
            return Json(new { success = status });
        }

        public ActionResult HistoryUpload()
        {
            if (Session["UserId"] != null && Session["RoleId"].ToString() == "1")
            {
                var dtHistoryContact = (from c in db.TR_Contact
                                        from u in db.TR_User.Where(x => x.UserId == c.CreatedBy).DefaultIfEmpty()
                                        from p in db.TT_CustomerProject.Where(x => x.CustProId == c.CustProId).DefaultIfEmpty()
                                        group new { c, u, p } by new
                                        {
                                            u.UserId,
                                            u.UserName,
                                            c.CreatedOn
                                        } into g
                                        select new contactCenterModels.HistoryUploadContact()
                                        {
                                            CreatedBy = g.Key.UserId,
                                            CreatedByName = g.Key.UserName == null ? "" : g.Key.UserName,
                                            CreatedOn = g.Key.CreatedOn,
                                            TotalContact = g.Count()
                                        }).OrderByDescending(p => p.CreatedOn).Take(20);

                ViewBag.HistoryUpload = dtHistoryContact.ToList();
                return View();
            }
            else
            {
                return RedirectToAction("Login", "Auth");
            }
        }
    }

    }
