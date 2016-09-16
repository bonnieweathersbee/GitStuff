using System;

using System.Data.SqlClient;

using System.Web.Mvc;
using Microsoft.SharePoint.Client;

using SP = Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Net;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";

            return View();
        }
        [HttpPost]
        public ActionResult test()
        {

            string mul = @"";

            int ids = 0; string QurAnd = ""; string AVItem = ""; string RoomsName = ""; string Type = "";
            string Table1 = @"<thead><tr><th>Edit</th><th> Delete </th><th> ID </th><th>AV Item </th><th>Rooms</th>  <th>Type</th> </tr> </thead<tbody>";
            string Table3 = "</tbody>"; string Table2 = ""; string returnmessage = ""; string curentdate = DateTime.Now.ToString();
            string connectionString = "Data Source=sully;Initial Catalog=PB_TEST;User ID=pierreb;Password=pierreb"; string querydata = "SELECT* FROM dbo.[AV]";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(querydata, connection))
                {
                    command.Connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader != null)
                        {
                            while (reader.Read())
                            {
                                if (ids != reader.GetInt32(0))
                                {
                                    ids = reader.GetInt32(0);
                                }


                                if (!reader.IsDBNull(reader.GetOrdinal("AVItem")))
                                {
                                    AVItem = reader.GetString(1);

                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("RoomsName")))
                                {
                                    RoomsName = reader.GetString(2);

                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Type")))
                                {
                                    Type = reader.GetString(3);

                                }

                                Table2 += "<tr id=t" + ids + " ><td><button type='button'id=trs" + ids + " class='btn btn-default ' data-toggle='modal' data-target='#myModalEdit' ><span class='glyphicon glyphicon-pencil' aria-hidden='true'></button> </td><td> <button type='button'id=tr" + ids + " class='btn btn-default ' data-toggle='modal' data-target='#myModalDelete' ><span class='glyphicon glyphicon-trash' aria-hidden='true'></button></td><td>" + ids + "</td><td>" +
                                    AVItem +
                                    "</td>  <td>" + RoomsName + "</td><td>" + Type + " </td></tr>";
                                QurAnd += "<script> $(document).on('click', '#trs" + ids + "', function () {  " +
                               " $('.modal-body #IDEdit').val($.trim('" + ids + "'));  $('.modal-body #AVItemEdit').val($.trim('" + AVItem + "')); $('.modal-body #RoomsNameEdit').val($.trim('" + RoomsName + "')); 	$('.modal-body #TypeEdit').val($.trim('" + Type + "')); }); </script >";
                                QurAnd += "<script> $(document).on('click', '#tr" + ids + "', function () {  " +
                                " $('.modal-body #ID').val($.trim('" + ids + "'));  $('.modal-body #AVItemDelete').val($.trim('" + AVItem + "')); }); </script >";

                                ids = 0; AVItem = ""; RoomsName = ""; Type = "";

                            }
                        }
                    }
                }

            }
            returnmessage = Table1 + Table2 + Table3 + QurAnd;

            return Content(returnmessage + mul);
        }
        [HttpPost]
       
        public ActionResult search(string startadte, string Endadte)
        {
            string title = "";
            string Table1 = "";
            string Table2 = " ";
            string Table3 = "";
            string MYTable = "";

            string Datem = "";
            string AM = " ";
            string PM = "";
            string AM9 = " ";
            string PM12 = "";
            string SearchTitle = "";
            string MyDate = "";
            string MyDate2 = "";
            string QurStart = "";
            SP.ClientContext clientContext = new ClientContext("http://intranet/dev/pr/");
            clientContext.Credentials = new NetworkCredential("SptWs", "artcev", "infodepot");
            SP.Web web = clientContext.Web;


            SP.List depList = web.Lists.GetById(new Guid("{E453CF27-3910-4367-A5D8-DFDE02E7B814}"));







            CamlQuery query = new CamlQuery();

            string Qcount = "2016-09-22";
            string Qcount2 = "2016-09-06";

            QurStart = @"<View ><Query>
   <Where>
      <And>
         <Leq>
            <FieldRef Name='Program_x0020_Start_x0020_Date' />
            <Value IncludeTimeValue='TRUE' Type='DateTime'>" + Endadte + @"</Value>
         </Leq>
         <Geq>
            <FieldRef Name='Program_x0020_Start_x0020_Date' />
            <Value IncludeTimeValue='TRUE' Type='DateTime'>" + startadte + @"</Value>
         </Geq>
      </And>
 
   </Where>
<OrderBy>
      <FieldRef Name='Program_x0020_Start_x0020_Date' Ascending='True' />
      <FieldRef Name='Start_x0020_Time' Ascending='True' />
   </OrderBy>
</Query></View>";

            //


            query.ViewXml = QurStart;


            SP.ListItemCollection _productCategories = depList.GetItems(query);




            clientContext.Load(_productCategories);

            clientContext.ExecuteQuery();
            string TIME = "";
            DateTime dt;
            DateTime dtittle;
            DateTime dtittle2;
            bool contains;
            bool contains1;
            bool contains2;
            int counter = 0;
            int check = 0;
            foreach (SP.ListItem item in _productCategories)//                for (int i = 0; i < _productCategories.Count; i++)
            {
                //SP.ListItem item = _productCategories[i];
                contains = Server.HtmlEncode(item["Start_x0020_Time"].ToString()).Contains("AM");
                contains1 = Server.HtmlEncode(item["Start_x0020_Time"].ToString()).Contains("9:");
                contains2 = Server.HtmlEncode(item["Start_x0020_Time"].ToString()).Contains("12:");
                // 
                dt = Convert.ToDateTime(Regex.Replace(Server.HtmlEncode(item["Program_x0020_Start_x0020_Date"].ToString()), @"-", ","));
                MyDate = Server.HtmlEncode(item["Program_x0020_Start_x0020_Date"].ToString());


                if (MyDate != MyDate2)
                {

                    Table2 += AM9 + AM + PM12 + PM;
                    Table2 += "<p> <b>" + dt.DayOfWeek + ", " + dt.ToString("MMMM") + "  " + dt.Day + "</b> </p> ";
                    PM = "";
                    AM = "";
                    PM12 = "";
                    AM9 = "";

                    if (contains == true)
                    {
                        if (contains1 == true)
                        {
                            AM9 = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                         "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";

                        }
                        else
                        {
                            AM = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                       "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                    }
                    if (contains == false)
                    {
                        if (contains2 == true)
                        {
                            PM12 = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                   "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";

                        }
                        else
                        {
                            PM = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                   "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                    }
                    MyDate2 = MyDate;
                    check++;

                }
                else
                {
                    if (contains == true)
                    {
                        if (contains1 == true)
                        {
                            AM9 += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                         "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";

                        }
                        else
                        {
                            AM += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                       "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }

                    }
                    if (contains == false)
                    {
                        if (contains2 == true)
                        {
                            PM12 += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                   "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                        else
                        {
                            PM += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                       "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                    }

                }





                counter++;
            }
            Table2 = Table2 + Datem + AM + PM;



            MYTable = Table1 + Table2 + Table3;

            string correctString = MYTable.Replace("PM", "p.m.");

            string correctString2 = correctString.Replace("AM", "a.m.");

            //Regex.Replace(MYTable, @" PM +", "p.m.");


            dtittle = Convert.ToDateTime(Regex.Replace(startadte, @"-", ","));
            dtittle2 = Convert.ToDateTime(Regex.Replace(Endadte, @"-", ","));

            SearchTitle = "<p><h2><b>You search for </b>:" + startadte + " to " + Endadte + "</h2></p></br> ";
            counter = 0;

            //Table2 += "<p><h1> <b>" + dt.DayOfWeek + ", " + dt.ToString("MMMM") + "  " + dt.Day + "</b> </h2> </p> ";string startadte, string Endadte


            title = "<p><b>Spartanburg County Public Libraries’ Calendar of Events (" + dtittle.ToString("MMMM") + "  " + dtittle.Day + " – " + dtittle2.ToString("MMMM") + "  " + dtittle2.Day + ")</b></p>";
            string endtittle = "<center><p><b> For a complete list of our events visit <a href='http://www.infodepot.org/'>www.infodepot.org </a> or call 864-596-3500.  </b></p></center>";
            // word();
            // myword(title + correctString2 + endtittle);






            // return Content(System.Web.HttpContext.Current.Server.MapPath(@"~/Word/S.docx"));

            return Content(SearchTitle + correctString2);
        }

        public ActionResult WordDoc(string startadte, string Endadte)
        {
            string title = "";
            string Table1 = "";
            string Table2 = " ";
            string Table3 = "";
            string MYTable = "";

            string Datem = "";
            string AM = " ";
            string PM = "";
            string AM9 = " ";
            string PM12 = "";
            string SearchTitle = "";
            string MyDate = "";
            string MyDate2 = "";
            string QurStart = "";
            SP.ClientContext clientContext = new ClientContext("http://intranet/dev/pr/");
            clientContext.Credentials = new NetworkCredential("SptWs", "artcev", "infodepot");
            SP.Web web = clientContext.Web;


            SP.List depList = web.Lists.GetById(new Guid("{E453CF27-3910-4367-A5D8-DFDE02E7B814}"));







            CamlQuery query = new CamlQuery();

            string Qcount = "2016-09-22";
            string Qcount2 = "2016-09-06";

            QurStart = @"<View ><Query>
   <Where>
      <And>
         <Leq>
            <FieldRef Name='Program_x0020_Start_x0020_Date' />
            <Value IncludeTimeValue='TRUE' Type='DateTime'>" + Endadte + @"</Value>
         </Leq>
         <Geq>
            <FieldRef Name='Program_x0020_Start_x0020_Date' />
            <Value IncludeTimeValue='TRUE' Type='DateTime'>" + startadte + @"</Value>
         </Geq>
      </And>
 
   </Where>
<OrderBy>
      <FieldRef Name='Program_x0020_Start_x0020_Date' Ascending='True' />
      <FieldRef Name='Start_x0020_Time' Ascending='True' />
   </OrderBy>
</Query></View>";

            //


            query.ViewXml = QurStart;


            SP.ListItemCollection _productCategories = depList.GetItems(query);




            clientContext.Load(_productCategories);

            clientContext.ExecuteQuery();
            string TIME = "";
            DateTime dt;
            DateTime dtittle;
            DateTime dtittle2;
            bool contains;
            bool contains1;
            bool contains2;
            int counter = 0;
            int check = 0;
            foreach (SP.ListItem item in _productCategories)//                for (int i = 0; i < _productCategories.Count; i++)
            {
                //SP.ListItem item = _productCategories[i];
                contains = Server.HtmlEncode(item["Start_x0020_Time"].ToString()).Contains("AM");
                contains1 = Server.HtmlEncode(item["Start_x0020_Time"].ToString()).Contains("9:");
                contains2 = Server.HtmlEncode(item["Start_x0020_Time"].ToString()).Contains("12:");
                // 
                dt = Convert.ToDateTime(Regex.Replace(Server.HtmlEncode(item["Program_x0020_Start_x0020_Date"].ToString()), @"-", ","));
                MyDate = Server.HtmlEncode(item["Program_x0020_Start_x0020_Date"].ToString());


                if (MyDate != MyDate2)
                {

                    Table2 += AM9 + AM + PM12 + PM;
                    Table2 += "<p> <b>" + dt.DayOfWeek + ", " + dt.ToString("MMMM") + "  " + dt.Day + "</b> </p> ";
                    PM = "";
                    AM = "";
                    PM12 = "";
                    AM9 = "";

                    if (contains == true)
                    {
                        if (contains1 == true)
                        {
                            AM9 = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                         "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";

                        }
                        else
                        {
                            AM = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                       "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                    }
                    if (contains == false)
                    {
                        if (contains2 == true)
                        {
                            PM12 = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                   "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";

                        }
                        else
                        {
                            PM = "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                   "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                    }
                    MyDate2 = MyDate;
                    check++;

                }
                else
                {
                    if (contains == true)
                    {
                        if (contains1 == true)
                        {
                            AM9 += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                         "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";

                        }
                        else
                        {
                            AM += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                       "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }

                    }
                    if (contains == false)
                    {
                        if (contains2 == true)
                        {
                            PM12 += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                                   "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                        else
                        {
                            PM += "<p><b>" + Server.HtmlEncode(item["Start_x0020_Time"].ToString()) +
                                       "</b> " + Server.HtmlEncode(item["Program_x0020_Location"].ToString()) + ", " + Server.HtmlEncode(item["Title"].ToString()) + ". " + Server.HtmlEncode(item["Program_x0020_Description"].ToString()) + " Call " + Server.HtmlEncode(item["Contact_x0020_Phone"].ToString()) + ".</p> ";
                        }
                    }

                }





                counter++;
            }
            Table2 = Table2 + Datem + AM + PM;



            MYTable = Table1 + Table2 + Table3;

            string correctString = MYTable.Replace("PM", "p.m.");

            string correctString2 = correctString.Replace("AM", "a.m.");

            //Regex.Replace(MYTable, @" PM +", "p.m.");


            dtittle = Convert.ToDateTime(Regex.Replace(startadte, @"-", ","));
            dtittle2 = Convert.ToDateTime(Regex.Replace(Endadte, @"-", ","));

  SearchTitle = "<p><h2><b>You search for </b>:" + startadte + " to " + Endadte + "</h2></p></br> ";
            counter = 0;

            //Table2 += "<p><h1> <b>" + dt.DayOfWeek + ", " + dt.ToString("MMMM") + "  " + dt.Day + "</b> </h2> </p> ";string startadte, string Endadte


            title = "<p><b>Spartanburg County Public Libraries’ Calendar of Events ("+ dtittle.ToString("MMMM") + "  "+ dtittle.Day + " – "+ dtittle2.ToString("MMMM") + "  " + dtittle2.Day + ")</b></p>";
            string endtittle = "<center><p><b> For a complete list of our events visit <a href='http://www.infodepot.org/'>www.infodepot.org </a> or call 864-596-3500.  </b></p></center>";
            // word();
            // myword(title + correctString2 + endtittle);


            // CreateWordprocessingDocument(System.Web.HttpContext.Current.Server.MapPath(@"~/Word/Invoice.docx"));
            //CreateWordprocessingDocument(title + correctString2 + endtittle);

            // return Content(System.Web.HttpContext.Current.Server.MapPath(@"~/Word/S.docx"));

            return Content(SearchTitle + "<div id='page-content'>"+ title + correctString2 + endtittle +"</div>");
        }
        public void word()
        {
            Microsoft.Office.Interop.Word._Application oWord;

            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;



            //oWord.Documents.Open("f:\\test.docx");
            //oWord.Selection.TypeText("Write your text here");


           // Microsoft.Office.Interop.Word.Document oDocument = new Microsoft.Office.Interop.Word.Document();
            //oDocument = oWord.Documents.Add();




            //oDocument.Content.Text = "Write your text here";
            //Create a missing variable for missing value
            object missing = System.Reflection.Missing.Value;

            //Create a new document
            Microsoft.Office.Interop.Word.Document document = oWord.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            //Add header into the document
            foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
            {
                //Get the header range and add the header details.
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                headerRange.Text = "Header text goes here";
            }

            //Add the footers into the document
            foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
            {
                //Get the footer range and add the footer details.
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 10;
               
                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = "Footer text goes here";
            }

            //adding text to document .Name = "Times New Roman";
            document.Content.SetRange(0, 0);
            document.Content.Font.Name = "Times New Roman";
            document.Content.Text = "This is test document " + Environment.NewLine;

            //Add paragraph with Heading 1 style
            Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            object styleHeading1 = "Heading 1";
            para1.Range.set_Style(ref styleHeading1);
            para1.Range.Text = "Para 1 text";
            para1.Range.InsertParagraphAfter();

            //Add paragraph with Heading 2 style
            Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
            object styleHeading2 = "Heading 2";
            para2.Range.set_Style(ref styleHeading2);
            para2.Range.Text = "Para 2 text";
            para2.Range.InsertParagraphAfter();
            DateTime localDate = DateTime.Now;
            // oDocument.SaveAs();

            
            
            oWord.Quit();

        }
        public void myword(string html)
        {

            // startInfo.Arguments = System.Web.HttpContext.Current.Server.MapPath(@"~/Word/S.docx");


            //string html = @"<div>Test 1</div>
            //   <div>Test 2,</div>
            //  <div>Test 3</div>";
            //  MemoryStream stream = new MemoryStream();
            // Write the string to a file.
          

            System.IO.StreamWriter file = new System.IO.StreamWriter(System.Web.HttpContext.Current.Server.MapPath(@"~/Word/S.docx"));
            file.WriteLine("mik");

            file.Close();
            // Apply the Heading 3 style to a paragraph. SaveToTemporaryFile(html)
           

            Application wordApp = new Application();
            wordApp.Visible = true;

            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
            Range rng = wordApp.ActiveDocument.Range(0, 0);
            rng.Text = "";

            object missing = Type.Missing;
            ContentControl contentControl = doc.ContentControls.Add(WdContentControlType.wdContentControlRichText, ref missing);
            
            contentControl.Title = "";
            contentControl.Range.InsertFile(SaveToTemporaryFile(html), ref missing, ref missing, ref missing, ref missing);


           


        }

        public static string SaveToTemporaryFile(string html)
        {
            // string htmlTempFilePath = Path.Combine(Path.GetTempPath(), string.Format("{0}.html", Path.GetRandomFileName()));C:\Users\Public\GetHub\WebApplication1\WebApplication1\Word\
            string htmlTempFilePath = System.IO.Path.Combine(System.Web.HttpContext.Current.Server.MapPath(@"~/Word"), System.IO.Path.GetRandomFileName()); //Path.Combine("WebApplication1/Word/", string.Format("{ 0}.html", Path.GetRandomFileName()));
            string htmlTempFilePath1 = htmlTempFilePath;
            using (StreamWriter writer = System.IO.File.CreateText(htmlTempFilePath))
            {
                html = string.Format("<html><meta http-equiv='Content-Type' content='text/html; charset=UTF-8'/>{0}</html>", html);

                writer.WriteLine(html);
            }

            return htmlTempFilePath;
        }
        
        public FileResult Download(string file)
        {

            //// filepath is a string which contains the path where the new document has to be createdSystem.Web.HttpContext.Current.Server.MapPath()
           
          




            string fullPath = System.Web.HttpContext.Current.Server.MapPath(@"~/Word/D1.docx");// application/vnd.openxmlformats-officedocument.wordprocessingml.document
            return File(fullPath, "application/msword", file);
        }
        public static void CreateWordprocessingDocument(string html)
        {
            
            // Open a doc file.
            Application application = new Application();
            Microsoft.Office.Interop.Word.Document document = application.Documents.Open(System.Web.HttpContext.Current.Server.MapPath(@"~/Word/D1.docx"));
            Range rng = application.ActiveDocument.Range(0, 0);
            rng.Text = "";
            object missing = Type.Missing;
            ContentControl contentControl = document.ContentControls.Add(WdContentControlType.wdContentControlRichText, ref missing);

            contentControl.Title = "";
            contentControl.Range.InsertFile(SaveToTemporaryFile(html), ref missing, ref missing, ref missing, ref missing);
            // Close word.
            document.SaveAs("MyFile.docx", ref missing, ref missing, ref missing, ref missing);
            application.Quit();
          

        }

    }

}
