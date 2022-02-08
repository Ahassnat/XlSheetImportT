using System;
using System.Collections.Generic;
using System.Data;
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
        public async Task<IActionResult> Index()
        {
            var ImportList = await _context.NVPCiscos.ToListAsync();
            //var model =  GetPagedData(ImportList, requiredPage);
            return View(ImportList);
        }
        public async Task<IActionResult> GetExel(NVPCiscoViewModel model)
        {
            var findSKu = await _context.NVPCiscos.FindAsync(model.PartSKU);
            // IEnumerable<NVPCisco> nVPCiscos = new IEnumerable<NVPCisco>();
            var fileName = "./wwwroot/Excel/20210106_Cisco_NVP_CE_Pricelist_Final.xlsb";
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
            const int pageSize = 3;
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
        public async Task<IActionResult> Download(string id)
        {
            var selectdResult = await _context.NVPCiscos.FindAsync(id);
            if (selectdResult == null) return NotFound();

            var path = "~/Excel/" + selectdResult.PartSKU;
            return File(path, "");
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
