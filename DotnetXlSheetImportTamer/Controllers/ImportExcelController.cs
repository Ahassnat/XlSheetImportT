using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using DotnetXlSheetImportTamer.Data;
using DotnetXlSheetImportTamer.Models;
using DotnetXlSheetImportTamer.Models.VM;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace DotnetXlSheetImportTamer.Controllers
{
    public class ImportExcelController : Controller
    {
        private readonly DataContext _context;
        //private readonly HttpClient _httpClient;

        public ImportExcelController(DataContext context)
        {
            _context = context;
            // _httpClient = httpClient;
        }
        public async Task<IActionResult> Index(int requiredPage = 1)
        {
           
            const int pageSize = 10;
            decimal rowsCount = await _context.NVPCiscos.CountAsync();
            var pagesCount = Math.Ceiling(rowsCount / pageSize);
            if (requiredPage > pagesCount)
            {
                requiredPage = 1;
            }
            if (requiredPage <= 0)
            {
                requiredPage = 1;
            }
            requiredPage = requiredPage <= 0 ? 1 : requiredPage;
            int skipCount = (requiredPage - 1) * pageSize;
           
           
            var ImportList = await _context.NVPCiscos.Skip(skipCount).Take(pageSize).ToListAsync();
            //var model =  GetPagedData(ImportList, requiredPage);
            ViewBag.CurrentPage = requiredPage;
            ViewBag.PagesCount = pagesCount;
            return View(ImportList);
        }
        public async Task<IActionResult> GetExel(NVPCiscoViewModel model)
        {
            //var findSKu = await _context.NVPCiscos.FindAsync(model.PartSKU);
            // IEnumerable<NVPCisco> nVPCiscos = new IEnumerable<NVPCisco>();
            var fileName = "./wwwroot/Excel/ExcelTest2.xlsx";
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            //using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            //{
            using (var package = new ExcelPackage())
            {
                using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    package.Load(stream);
                    //var worksheets = package.Workbook.Worksheets.ToList();
                    var worksheet = package.Workbook.Worksheets.First();
                    var rowCount = worksheet.Dimension.Rows;

                    for (var row = 3; row <= rowCount; row++)
                    {
                        try
                        {
                            var PartSKU = worksheet.Cells[row, 4].Value?.ToString();
                            if (string.IsNullOrWhiteSpace(PartSKU)) continue;

                            if (_context.NVPCiscos.Any(x => x.PartSKU == PartSKU)) continue;

                            model.PartSKU = PartSKU;
                            model.Brand = worksheet.Cells[row, 1].Value?.ToString();
                            model.CategoryCode = worksheet.Cells[row, 2].Value?.ToString();
                            model.Manufacturer = worksheet.Cells[row, 3].Value?.ToString();
                            
                            model.ItemDescription = worksheet.Cells[row, 5].Value?.ToString();
                            model.PriceList = worksheet.Cells[row, 6].Value?.ToString();
                            model.MinDiscount = worksheet.Cells[row, 7].Value?.ToString();
                            model.DiscountPrice = worksheet.Cells[row, 8].Value?.ToString();

                            //var nvp = new NVPCisco()
                            //{
                            //    Brand = model.Brand,
                            //    CategoryCode = model.CategoryCode,
                            //    DiscountPrice = model.DiscountPrice,
                            //    ItemDescription = model.ItemDescription,
                            //    Manufacturer = model.Manufacturer,
                            //    MinDiscount = model.MinDiscount,
                            //    PartSKU = model.PartSKU,
                            //    PriceList = model.PriceList
                            //};

                            //nVPCiscos.Add(nvp);

                            var obj = new NVPCisco
                            {
                                Brand = model.Brand,
                                CategoryCode = model.CategoryCode,
                                DiscountPrice = model.DiscountPrice,
                                ItemDescription = model.ItemDescription,
                                Manufacturer = model.Manufacturer,
                                MinDiscount = model.MinDiscount,
                                PartSKU = model.PartSKU,
                                PriceList = model.PriceList
                            };

                            await _context.NVPCiscos.AddAsync(obj);
                            
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }

                }

            }


            //2- read the xl from url 
            //3- map data with local db 
            //4- search and pagenation 
            //5-  show the grid table
            await _context.SaveChangesAsync();
            return RedirectToAction("Index");
        }



        /* Pagination Function  */
        private async Task<List<NVPCiscoViewModel>> GetPagedData(IQueryable<NVPCiscoViewModel> result, int requiredPage = 1)
        {
            const int pageSize = 10;
            decimal rowsCount = await _context.NVPCiscos.CountAsync();

            var pagesCount = Math.Ceiling(rowsCount / pageSize);

            if (requiredPage > pagesCount)
            {
                requiredPage = 1;
            }
            if (requiredPage <= 0)
            {
                requiredPage = 1;
            }

            int skipCount = (requiredPage - 1) * pageSize;

            //select count(*) from NVPCisco ;
            var pagedData = await result
                .Skip(skipCount)
                .Take(pageSize)
                .ToListAsync();
            ViewBag.CurrentPage = requiredPage;
            ViewBag.PagesCount = pagesCount;

            return pagedData;
        }
        [HttpGet]
        public async Task<IActionResult> Download()
        {
            var selectedUpload = await _context.NVPCiscos.ToListAsync();
           // List<NVPCiscoViewModel> selcectedUpload = new List<NVPCiscoViewModel>();                                                             
            if (selectedUpload != null)
                if (selectedUpload == null) return NotFound();


            // Export Excel File 
            var res = selectedUpload;

            // Start exporting to Excel
            var stream = new MemoryStream();

            using (var xlPackage = new ExcelPackage(stream))
            {
                // Define a worksheet
                var worksheet = xlPackage.Workbook.Worksheets.Add("res");

                // Styling
                var customStyle = xlPackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                customStyle.Style.Font.UnderLine = true;
                customStyle.Style.Font.Color.SetColor(Color.Red);

                // First row
                var startRow = 8;
                var row = startRow;

                worksheet.Cells["A1"].Value = "Search Result";
                using (var r = worksheet.Cells["A1:H1"])
                {
                    r.Merge = true;
                    r.Style.Font.Color.SetColor(Color.Green);
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                }

                worksheet.Cells["A4"].Value = "Brand";
                worksheet.Cells["B4"].Value = "Category Code";
                worksheet.Cells["C4"].Value = "Manufacturer";

                worksheet.Cells["D4"].Value = "Part SKU";
                worksheet.Cells["E4"].Value = "Item Description";
                worksheet.Cells["F4"].Value = "PriceList";
                worksheet.Cells["G4"].Value = "Min Discount";
                worksheet.Cells["H4"].Value = "Discount Price";


                worksheet.Cells["A4:H4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells["A4:H4"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                row = 8;
                foreach (var rec in res)
                {
                    worksheet.Cells[row, 1].Value = rec.Brand;
                    worksheet.Cells[row, 2].Value = rec.CategoryCode;
                    worksheet.Cells[row, 3].Value = rec.Manufacturer;
                    worksheet.Cells[row, 4].Value = rec.PartSKU;
                    worksheet.Cells[row, 5].Value = rec.ItemDescription;
                    worksheet.Cells[row, 6].Value = rec.PriceList;
                    worksheet.Cells[row, 7].Value = rec.MinDiscount;
                    worksheet.Cells[row, 8].Value = rec.DiscountPrice;

                    row++; // row = row + 1;
                }

                xlPackage.Workbook.Properties.Title = "Result List";
                xlPackage.Workbook.Properties.Author = "Abdallah Hassnat";

                xlPackage.Save();
            }

            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SearchResult.xlsx");
            
        }



        //[HttpGet]
        //public async Task<IActionResult> Browse(int requiredPage = 1)
        //{
        //    var result = _context.NVPCiscos
        //    .Select(x => new NVPCiscoViewModel
        //    {

        //        Brand = x.Brand,
        //        CategoryCode = x.CategoryCode,
        //        DiscountPrice = x.DiscountPrice,
        //        ItemDescription = x.ItemDescription,
        //        Manufacturer = x.Manufacturer,
        //        MinDiscount = x.MinDiscount,
        //        PartSKU = x.PartSKU,
        //        PriceList = x.PriceList

        //    });
        //    var model = await GetPagedData(result, requiredPage);
        //    return View(model);
        //}

        [HttpGet]
        public async Task< IActionResult> SearchResults(int requiredPage = 1)
        {
            var result = _context.NVPCiscos
            .Select(x => new NVPCiscoViewModel
            {

                Brand = x.Brand,
                CategoryCode = x.CategoryCode,
                DiscountPrice = x.DiscountPrice,
                ItemDescription = x.ItemDescription,
                Manufacturer = x.Manufacturer,
                MinDiscount = x.MinDiscount,
                PartSKU = x.PartSKU,
                PriceList = x.PriceList

            });
            var model = await GetPagedData(result, requiredPage);
            return View(model);
        }


        [HttpPost]
        // Search
        public async Task<IActionResult> SearchResults(string term, int requiredPage = 1)
        {
            var result = _context.NVPCiscos.Where(x => x.ItemDescription.Contains(term) ||
                                                   x.Manufacturer.Contains(term) ||
                                                   x.PartSKU.Contains(term) ||
                                                   x.Brand.Contains(term) ||
                                                   x.CategoryCode.Contains(term))

            .Select(x => new NVPCiscoViewModel
            {

                Brand = x.Brand,
                CategoryCode = x.CategoryCode,
                DiscountPrice = x.DiscountPrice,
                ItemDescription = x.ItemDescription,
                Manufacturer = x.Manufacturer,
                MinDiscount = x.MinDiscount,
                PartSKU = x.PartSKU,
                PriceList = x.PriceList

            });
            var model = await GetPagedData(result, requiredPage);
            return View(model);
        }


    }
}
