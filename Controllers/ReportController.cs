using MVC_CRUD.Models;
using System;
using System.Collections.Generic;
using System.IO;
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
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Web;

namespace MVC_CRUD.Controllers
{
    public class ReportController : Controller
    {
        public ActionResult ReportPerformance(string From, string To, string ProjectId)
        {
            List<Dictionary<String, String>> _list = new List<Dictionary<string, string>>();

            NameValueCollection nvc = Request.Form;
            DateTime dtFrom = Convert.ToDateTime(From);
            DateTime dtTo = Convert.ToDateTime(To);
            int CustProId = Convert.ToInt32(ProjectId);

            int target = 0;

            var result = new List<contactCenterModels.Performance>();
            DateTime dateNow = DateTime.Now;
            String StartDate = dtFrom.ToString("yyyy-MM-dd");
            String EndDate = dtTo.ToString("yyyy-MM-dd");

            ParamReport pr = new ParamReport();
            pr.CustomerName = "";
            pr.ProjectName = "";
            pr.DateFrom = dtFrom.ToString("dd-MMM-yyyy");
            pr.DateTo = dtTo.ToString("dd-MMM-yyyy");

            //Defined SQL CONNECTION
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;

            using (SqlConnection dbconnection = connection)
            {
                //membuka koneksi database
                dbconnection.Open();

                String query = @"( select
	                                    CustProName
	                                    , CustomerName
                                    from TT_CustomerProject p
	                                    inner join TR_Customer c
		                                    on p.CustomerId = c.CustomerId
                                    where CustProId = " + CustProId + " )";
                using (SqlDataAdapter cmd = new SqlDataAdapter(query, dbconnection))
                {
                    DataSet CustPro = new DataSet("CustomerProject");
                    cmd.FillSchema(CustPro, SchemaType.Source, "CustomerProject");
                    cmd.Fill(CustPro, "CustomerProject");

                    DataTable tblCustPro;
                    tblCustPro = CustPro.Tables["CustomerProject"];
                    for (int i = 0; i < tblCustPro.Rows.Count; i++)
                    {
                        pr.CustomerName = tblCustPro.Rows[i][0].ToString();
                        pr.ProjectName = tblCustPro.Rows[i][1].ToString();
                    }
                }

                query = @"( select a.TargetData/b.UserId as Jumlah 
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

            ExportPDFPerformance(result, pr);
            return View();
        }

        public ActionResult ReportProductivity(string From, string To, string ProjectId)
        {
            List<Dictionary<String, String>> _list = new List<Dictionary<string, string>>();

            NameValueCollection nvc = Request.Form;
            DateTime dtFrom = Convert.ToDateTime(From);
            DateTime dtTo = Convert.ToDateTime(To);
            int CustProId = Convert.ToInt32(ProjectId);

            int target = 0;

            var productivity = new List<contactCenterModels.Productivity>();
            DateTime dateNow = DateTime.Now;
            String StartDate = dtFrom.ToString("yyyy-MM-dd");
            String EndDate = dtTo.ToString("yyyy-MM-dd");

            ParamReport pr = new ParamReport();
            pr.CustomerName = "";
            pr.ProjectName = "";
            pr.DateFrom = dtFrom.ToString("dd-MMM-yyyy");
            pr.DateTo = dtTo.ToString("dd-MMM-yyyy");

            //Defined SQL CONNECTION
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;

            using (SqlConnection dbconnection = connection)
            {
                //membuka koneksi database
                dbconnection.Open();

                String query = @"( select
	                                    CustProName
	                                    , CustomerName
                                    from TT_CustomerProject p
	                                    inner join TR_Customer c
		                                    on p.CustomerId = c.CustomerId
                                    where CustProId = " + CustProId + " )";
                using (SqlDataAdapter cmd = new SqlDataAdapter(query, dbconnection))
                {
                    DataSet CustPro = new DataSet("CustomerProject");
                    cmd.FillSchema(CustPro, SchemaType.Source, "CustomerProject");
                    cmd.Fill(CustPro, "CustomerProject");

                    DataTable tblCustPro;
                    tblCustPro = CustPro.Tables["CustomerProject"];
                    for (int i = 0; i < tblCustPro.Rows.Count; i++)
                    {
                        pr.CustomerName = tblCustPro.Rows[i][0].ToString();
                        pr.ProjectName = tblCustPro.Rows[i][1].ToString();
                    }
                }

                query = @"( select a.TargetData/b.UserId as Jumlah 
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

            ExportPDFProductivity(productivity, pr);
            return View();
        }


        void ExportPDFPerformance(List<Performance> pr, ParamReport prp)
        {
            try
            {
                BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
                iTextSharp.text.Font times = new iTextSharp.text.Font(bfTimes, 14, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                Font font8 = FontFactory.GetFont("ARIAL", 12, Font.BOLD, BaseColor.BLACK);

                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
                pdfDoc.SetMargins(8f, 8f, 8f, 8f);
                PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
                pdfDoc.Open();
                //pdfDoc.Add(new Paragraph("Welcome to dotnetfox"));

                Response.ContentType = "application/pdf";
                Response.AddHeader("content-disposition", "attachment;" +
                    "filename=Report Performance.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);

                Font fontstys = FontFactory.GetFont("Times New Roman", 8, BaseColor.BLACK);
                Font fontMap = FontFactory.GetFont("Arial", 10, BaseColor.BLUE);
                Font fontbold = FontFactory.GetFont("Times New Roman", 8, Font.BOLD, BaseColor.BLACK);
                PdfPTable table = new PdfPTable(3);

                float[] widths = new float[] { 0.6f, 0.1f, 3f };
                table.SetWidths(widths);
                table.WidthPercentage = 90f;

                //Detail Merchant
                PdfPCell cell = new PdfPCell(new Phrase("Report Performance", font8));
                cell.Colspan = 3;
                cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                cell.BackgroundColor = BaseColor.CYAN;


                PdfPCell cells = new PdfPCell(new Phrase(" "));
                cells.Colspan = 3;
                cells.Border = 0;
                cells.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right

                PdfPCell cell1 = new PdfPCell(new Phrase("Customer ", fontstys));
                cell1.Border = 0;
                PdfPCell cell1a = new PdfPCell(new Phrase(":", fontstys));
                cell1a.Border = 0;
                PdfPCell cell1s = new PdfPCell(new Phrase(prp.CustomerName, fontstys));
                cell1s.Border = 0;

                PdfPCell cell2 = new PdfPCell(new Phrase("Project ", fontstys));
                cell2.Border = 0;
                PdfPCell cell2a = new PdfPCell(new Phrase(":", fontstys));
                cell2a.Border = 0;
                PdfPCell cell2s = new PdfPCell(new Phrase(prp.ProjectName, fontstys));
                cell2s.Border = 0;

                PdfPCell cell3 = new PdfPCell(new Phrase("From ", fontstys));
                cell3.Border = 0;
                PdfPCell cell3a = new PdfPCell(new Phrase(":", fontstys));
                cell3a.Border = 0;
                PdfPCell cell3s = new PdfPCell(new Phrase(prp.DateFrom, fontstys));
                cell3s.Border = 0;

                PdfPCell cell4 = new PdfPCell(new Phrase("To ", fontstys));
                cell4.Border = 0;
                PdfPCell cell4a = new PdfPCell(new Phrase(":", fontstys));
                cell4a.Border = 0;
                PdfPCell cell4s = new PdfPCell(new Phrase(prp.DateTo, fontstys));
                cell4s.Border = 0;


                table.AddCell(cell);
                table.AddCell(cell1);
                table.AddCell(cell1a);
                table.AddCell(cell1s);
                table.AddCell(cell2);
                table.AddCell(cell2a);
                table.AddCell(cell2s);
                table.AddCell(cell3);
                table.AddCell(cell3a);
                table.AddCell(cell3s);
                table.AddCell(cell4);
                table.AddCell(cell4a);
                table.AddCell(cell4s);

                //Table Report
                PdfPTable tableCust = new PdfPTable(1);
                tableCust.WidthPercentage = 90f;

                PdfPCell MapSpasi = new PdfPCell(new Phrase("\n"));
                MapSpasi.Border = 0;
                tableCust.AddCell(MapSpasi);

                //Data Detail Report
                PdfPTable tbReport = new PdfPTable(9);
                float[] widthReport = new float[] { 1f, 0.2f, 0.3f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.7f};
                tbReport.SetWidths(widthReport);
                tbReport.WidthPercentage = 90f;

                PdfPCell Cell;
                Cell = new PdfPCell(new Phrase("Agent Name", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Achievment", fontbold));
                Cell.HorizontalAlignment = 1;
                Cell.Colspan = 2;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Target", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Closing", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Prospect", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Contacted", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Connected", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Not Connected", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                foreach (var dtReport in pr)
                {
                    PdfPCell CellUser = new PdfPCell(new Phrase(dtReport.UserName, fontstys));
                    CellUser.HorizontalAlignment = 0;
                    tbReport.AddCell(CellUser);

                    PdfPCell CellAchImage = ImageCell(dtReport.Image, 40f, PdfPCell.ALIGN_CENTER);
                    CellAchImage.HorizontalAlignment = 1;
                    CellAchImage.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    CellAchImage.BorderWidthTop = 0;
                    CellAchImage.BorderWidthLeft = 0;
                    CellAchImage.BorderWidthRight = 0;
                    tbReport.AddCell(CellAchImage);

                    PdfPCell CellAch = new PdfPCell(new Phrase("100.00", fontstys));
                    CellAch.HorizontalAlignment = 1;
                    CellAch.BorderWidthLeft = 0;
                    CellAch.BorderWidthTop = 0;
                    tbReport.AddCell(CellAch);

                    PdfPCell CellTarget = new PdfPCell(new Phrase(dtReport.Target.ToString(), fontstys));
                    CellTarget.HorizontalAlignment = 1;
                    tbReport.AddCell(CellTarget);

                    PdfPCell CellClosing = new PdfPCell(new Phrase(dtReport.Closing.ToString(), fontstys));
                    CellClosing.HorizontalAlignment = 1;
                    tbReport.AddCell(CellClosing);

                    PdfPCell CellProspect = new PdfPCell(new Phrase(dtReport.Prospect.ToString(), fontstys));
                    CellProspect.HorizontalAlignment = 1;
                    tbReport.AddCell(CellProspect);

                    PdfPCell CellContacted = new PdfPCell(new Phrase(dtReport.Contacted.ToString(), fontstys));
                    CellContacted.HorizontalAlignment = 1;
                    tbReport.AddCell(CellContacted);

                    PdfPCell CellConnected = new PdfPCell(new Phrase(dtReport.Connected.ToString(), fontstys));
                    CellConnected.HorizontalAlignment = 1;
                    tbReport.AddCell(CellConnected);

                    PdfPCell CellNotConnected = new PdfPCell(new Phrase(dtReport.NotConnected.ToString(), fontstys));
                    CellNotConnected.HorizontalAlignment = 1;
                    tbReport.AddCell(CellNotConnected);
                }
                pdfDoc.Add(new Paragraph("\n"));
                pdfDoc.Add(table);
                //pdfDoc.Add(new Paragraph("\n"));
                pdfDoc.Add(tableCust);
                pdfDoc.Add(tbReport);
                pdfDoc.Add(new Paragraph("\n"));

                pdfDoc.Close();
                Response.Write(pdfDoc);
                Response.End();
            }
            catch (Exception ex)
            {

            }
        }

        void ExportPDFProductivity(List<Productivity> pr, ParamReport prp)
        {
            try
            {
                BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
                iTextSharp.text.Font times = new iTextSharp.text.Font(bfTimes, 14, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                Font font8 = FontFactory.GetFont("ARIAL", 12, Font.BOLD, BaseColor.BLACK);

                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
                pdfDoc.SetMargins(8f, 8f, 8f, 8f);
                PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
                pdfDoc.Open();
                //pdfDoc.Add(new Paragraph("Welcome to dotnetfox"));

                Response.ContentType = "application/pdf";
                Response.AddHeader("content-disposition", "attachment;" +
                    "filename=Report Productivity.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);

                Font fontstys = FontFactory.GetFont("Times New Roman", 8, BaseColor.BLACK);
                Font fontMap = FontFactory.GetFont("Arial", 10, BaseColor.BLUE);
                Font fontbold = FontFactory.GetFont("Times New Roman", 8, Font.BOLD, BaseColor.BLACK);
                PdfPTable table = new PdfPTable(3);

                float[] widths = new float[] { 0.6f, 0.1f, 3f };
                table.SetWidths(widths);
                table.WidthPercentage = 80f;

                //Detail Merchant
                PdfPCell cell = new PdfPCell(new Phrase("Report Productivity", font8));
                cell.Colspan = 3;
                cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                cell.BackgroundColor = BaseColor.CYAN;


                PdfPCell cells = new PdfPCell(new Phrase(" "));
                cells.Colspan = 3;
                cells.Border = 0;
                cells.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right

                PdfPCell cell1 = new PdfPCell(new Phrase("Customer ", fontstys));
                cell1.Border = 0;
                PdfPCell cell1a = new PdfPCell(new Phrase(":", fontstys));
                cell1a.Border = 0;
                PdfPCell cell1s = new PdfPCell(new Phrase(prp.CustomerName, fontstys));
                cell1s.Border = 0;

                PdfPCell cell2 = new PdfPCell(new Phrase("Project ", fontstys));
                cell2.Border = 0;
                PdfPCell cell2a = new PdfPCell(new Phrase(":", fontstys));
                cell2a.Border = 0;
                PdfPCell cell2s = new PdfPCell(new Phrase(prp.ProjectName, fontstys));
                cell2s.Border = 0;

                PdfPCell cell3 = new PdfPCell(new Phrase("From ", fontstys));
                cell3.Border = 0;
                PdfPCell cell3a = new PdfPCell(new Phrase(":", fontstys));
                cell3a.Border = 0;
                PdfPCell cell3s = new PdfPCell(new Phrase(prp.DateFrom, fontstys));
                cell3s.Border = 0;

                PdfPCell cell4 = new PdfPCell(new Phrase("To ", fontstys));
                cell4.Border = 0;
                PdfPCell cell4a = new PdfPCell(new Phrase(":", fontstys));
                cell4a.Border = 0;
                PdfPCell cell4s = new PdfPCell(new Phrase(prp.DateTo, fontstys));
                cell4s.Border = 0;


                table.AddCell(cell);
                table.AddCell(cell1);
                table.AddCell(cell1a);
                table.AddCell(cell1s);
                table.AddCell(cell2);
                table.AddCell(cell2a);
                table.AddCell(cell2s);
                table.AddCell(cell3);
                table.AddCell(cell3a);
                table.AddCell(cell3s);
                table.AddCell(cell4);
                table.AddCell(cell4a);
                table.AddCell(cell4s);

                //Table Report
                PdfPTable tableCust = new PdfPTable(1);
                tableCust.WidthPercentage = 80f;

                PdfPCell MapSpasi = new PdfPCell(new Phrase("\n"));
                MapSpasi.Border = 0;
                tableCust.AddCell(MapSpasi);

                //Data Detail Report
                PdfPTable tbReport = new PdfPTable(4);
                float[] widthReport = new float[] { 1f, 0.5f, 0.5f, 0.5f};
                tbReport.SetWidths(widthReport);
                tbReport.WidthPercentage = 70f;

                PdfPCell Cell;
                Cell = new PdfPCell(new Phrase("Agent Name", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Call Attempt", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Utilization", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                Cell = new PdfPCell(new Phrase("Talk Time", fontbold));
                Cell.HorizontalAlignment = 1;
                tbReport.AddCell(Cell);

                foreach (var dtReport in pr)
                {
                    PdfPCell CellUser = new PdfPCell(new Phrase(dtReport.UserName, fontstys));
                    CellUser.HorizontalAlignment = 0;
                    tbReport.AddCell(CellUser);
                    
                    PdfPCell CellTarget = new PdfPCell(new Phrase(dtReport.callAtemp.ToString(), fontstys));
                    CellTarget.HorizontalAlignment = 1;
                    tbReport.AddCell(CellTarget);

                    PdfPCell CellClosing = new PdfPCell(new Phrase(dtReport.Utillization.ToString(), fontstys));
                    CellClosing.HorizontalAlignment = 1;
                    tbReport.AddCell(CellClosing);

                    PdfPCell CellProspect = new PdfPCell(new Phrase(dtReport.talkTime.ToString(), fontstys));
                    CellProspect.HorizontalAlignment = 1;
                    tbReport.AddCell(CellProspect);
                }
                pdfDoc.Add(new Paragraph("\n"));
                pdfDoc.Add(table);
                //pdfDoc.Add(new Paragraph("\n"));
                pdfDoc.Add(tableCust);
                pdfDoc.Add(tbReport);
                pdfDoc.Add(new Paragraph("\n"));

                pdfDoc.Close();
                Response.Write(pdfDoc);
                Response.End();
            }
            catch (Exception ex)
            {

            }
        }


        private PdfPCell ImageCell(string path, float scale, int align)
        {
            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(Server.MapPath(path));

            image.ScalePercent(scale);
            PdfPCell cell = new PdfPCell(image);
            cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
            cell.HorizontalAlignment = align;
            return cell;
        }
    }

    public class ParamReport
    {
        public string CustomerName { get; set; }
        public string ProjectName { get; set; }
        public string DateFrom { get; set; }
        public string DateTo { get; set; }
    }
}