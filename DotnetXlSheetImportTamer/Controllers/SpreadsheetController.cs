//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.Linq;

//using Microsoft.AspNetCore.Hosting;
//using Microsoft.AspNetCore.Mvc;

//namespace DotnetXlSheetImportTamer.Controllers
//{
//    public class SpreadsheetController : Controller
//    {
//        public IActionResult Index()
//        {
// 1- Import From online ExcelFile Url
//URL => https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fwww.cisco.com%2Fc%2Fdam%2Fen_us%2Fsolutions%2Findustries%2Fgovernment%2Fnvpce%2Fdocs%2Fpricelists%2F2021%2F20210106_Cisco_NVP_CE_Pricelist_Final.xlsb&wdOrigin=BROWSELINK
//URL => https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fwww.cisco.com%2Fc%2Fdam%2Fen_us%2Fsolutions%2Findustries%2Fgovernment%2Fnvpce%2Fdocs%2Fpricelists%2F2021%2F20210106_Cisco_NVP_CE_Pricelist_Final.xlsb&wdOrigin=BROWSELINK

//var myStringUrl = await _httpClient.GetStringAsync("https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fwww.cisco.com%2Fc%2Fdam%2Fen_us%2Fsolutions%2Findustries%2Fgovernment%2Fnvpce%2Fdocs%2Fpricelists%2F2021%2F20210106_Cisco_NVP_CE_Pricelist_Final.xlsb&wdOrigin=BROWSELINK");



#region 
//    try
//    {
//        #region Variable Declaration  
//        string message = "";
//        HttpResponseMessage ResponseMessage = null;
//        var httpRequest = HttpContext.Current.Request;
//        DataSet dsexcelRecords = new DataSet();
//        IExcelDataReader reader = null;
//        HttpPostedFile Inputfile = null;
//        Stream FileStream = null;
//        #endregion

//        #region Save  Detail From Excel  To Db
//        using (dbCodingvilaEntities objEntity = new dbCodingvilaEntities())
//        {
//            if (httpRequest.Files.Count > 0)
//            {
//                Inputfile = httpRequest.Files[0];
//                FileStream = Inputfile.InputStream;

//                if (Inputfile != null && FileStream != null)
//                {
//                    if (Inputfile.FileName.EndsWith(".xls"))
//                        reader = ExcelReaderFactory.CreateBinaryReader(FileStream);
//                    else if (Inputfile.FileName.EndsWith(".xlsx"))
//                        reader = ExcelReaderFactory.CreateOpenXmlReader(FileStream);
//                    else
//                        message = "The file format is not supported.";

//                    dsexcelRecords = reader.AsDataSet();
//                    reader.Close();

//                    if (dsexcelRecords != null && dsexcelRecords.Tables.Count > 0)
//                    {
//                        DataTable dtStudentRecords = dsexcelRecords.Tables[0];
//                        for (int i = 0; i < dtStudentRecords.Rows.Count; i++)
//                        {
//                            Student objStudent = new Student();
//                            objStudent.RollNo = Convert.ToInt32(dtStudentRecords.Rows[i][0]);
//                            objStudent.EnrollmentNo = Convert.ToString(dtStudentRecords.Rows[i][1]);
//                            objStudent.Name = Convert.ToString(dtStudentRecords.Rows[i][2]);
//                            objStudent.Branch = Convert.ToString(dtStudentRecords.Rows[i][3]);
//                            objStudent.University = Convert.ToString(dtStudentRecords.Rows[i][4]);
//                            objEntity.Students.Add(objStudent);
//                        }

//                        int output = objEntity.SaveChanges();
//                        if (output > 0)
//                            message = "The Excel file has been successfully uploaded.";
//                        else
//                            message = "Something Went Wrong!, The Excel file uploaded has fiald.";
//                    }
//                    else
//                        message = "Selected file is empty.";
//                }
//                else
//                    message = "Invalid File.";
//            }
//            else
//                ResponseMessage = Request.CreateResponse(HttpStatusCode.BadRequest);
//        }
//        return message;
//        #endregion
//    }
//    catch (Exception)
//    {
//        throw;
//    }
//}
#endregion


// convert to a stream


/* HttpClient http = new HttpClient();
 http.DefaultRequestHeaders.Add(schemename, header);
 var data = http.GetAsync(https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fwww.cisco.com%2Fc%2Fdam%2Fen_us%2Fsolutions%2Findustries%2Fgovernment%2Fnvpce%2Fdocs%2Fpricelists%2F2021%2F20210106_Cisco_NVP_CE_Pricelist_Final.xlsb&wdOrigin=BROWSELINK).Result.Content.ReadAsStringAsync().Result;
 //return data;*/




//            var urlS = "https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fwww.cisco.com%2Fc%2Fdam%2Fen_us%2Fsolutions%2Findustries%2Fgovernment%2Fnvpce%2Fdocs%2Fpricelists%2F2021%2F20210106_Cisco_NVP_CE_Pricelist_Final.xlsb&wdOrigin=BROWSELINK";

//            var ds = new DataTable();
//            return View();
//        }

//    }
//}
