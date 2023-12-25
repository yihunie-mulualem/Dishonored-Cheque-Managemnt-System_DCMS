using DCMS.Data;
using DCMS.Help;
using DCMS.Models;
using DCMS.Models.Setup;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Composition;
using System.Drawing;
using System.IO;
using static DCMS.Help.Helper;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DCMS.Controllers
{
    [CheckSessionIsAvailable]
    [NoDirectAccess]
    public class ReportController : Controller
    {
        private readonly DCMSDbContext _context;
        private readonly IHttpContextAccessor _httpContext;
        public ReportController(DCMSDbContext context, IHttpContextAccessor httpContext)
        {
            _context = context;
            _httpContext = httpContext;

        }
        /// /////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public IActionResult BusinessUnitRecords()
        {
            var branchId = _httpContext.HttpContext.Session.GetInt32("UserBranch");
            var instances = _context.InstanceTypes.ToList();
            ViewBag.InstanceTypeList = new SelectList(instances, "Id", "Name");
            return View(_context.DishonoredCheques.Where(x => x.Branch == branchId).ToList());
        }
        [HttpPost]
        public IActionResult BusinessUnitRecords(int InstanceId, DateTime? from, DateTime? to)
        {
            DateTime dateTime = DateTime.Now;
            var branchId = _httpContext.HttpContext.Session.GetInt32("UserBranch");
            var instances = _context.InstanceTypes.ToList();
            ViewBag.InstanceTypeList = new SelectList(instances, "Id", "Name");

            if (InstanceId != 0 && from != null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.Branch == branchId && x.IssueDate >= from && x.IssueDate <= to).ToList());
            }
            else if (InstanceId == 0 && from != null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.IssueDate >= from && x.Branch == branchId && x.IssueDate <= to).ToList());
            }
            else if (InstanceId != 0 && from == null && to == null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.Branch == branchId).ToList());
            }
            else if (InstanceId == 0 && from == null && to == null)
            {
                return View(_context.DishonoredCheques.Where(x => x.Branch == branchId).ToList());
            }
            else if (InstanceId == 0 && from == null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.IssueDate <= to && x.Branch == branchId).ToList());
            }
            else if (InstanceId == 0 && from != null && to == null)
            {
                return View(_context.DishonoredCheques.Where(x => x.IssueDate >= from && x.Branch == branchId).ToList());
            }
            else if (InstanceId != 0 && from == null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id >= InstanceId && x.IssueDate <= to && x.Branch == branchId).ToList());
            }
            else if (InstanceId != 0 && from != null && to == null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id >= InstanceId && x.IssueDate >= from && x.Branch == branchId).ToList());
            }
            else
            {
                TempData["msg_tbl"] = "Searching Key Error";
                return View(_context.DishonoredCheques.Where(x => x.Branch == branchId).ToList());
            }
        }
        public IActionResult AllRegisteredRecords()
        {
            var instances = _context.InstanceTypes.ToList();
            ViewBag.InstanceTypeList = new SelectList(instances, "Id", "Name");
            return View(_context.DishonoredCheques.ToList());
        }
        [HttpPost]
        public IActionResult AllRegisteredRecords(int InstanceId, DateTime? from, DateTime? to)
        {
            var instances = _context.InstanceTypes.ToList();
            ViewBag.InstanceTypeList = new SelectList(instances, "Id", "Name");

            if (InstanceId != 0 && from != null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.IssueDate >= from && x.IssueDate <= to).ToList());
            }
            else if (InstanceId == 0 && from != null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.IssueDate >= from && x.IssueDate <= to).ToList());
            }
            else if (InstanceId != 0 && from == null && to == null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId).ToList());
            }
            else if (InstanceId == 0 && from == null && to == null)
            {
                return View(_context.DishonoredCheques.ToList());
            }
            else if (InstanceId == 0 && from == null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.IssueDate <= to).ToList());
            }
            else if (InstanceId == 0 && from != null && to == null)
            {
                return View(_context.DishonoredCheques.Where(x => x.IssueDate >= from).ToList());
            }
            else if (InstanceId != 0 && from == null && to != null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.IssueDate <= to).ToList());
            }
            else if (InstanceId != 0 && from != null && to == null)
            {
                return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.IssueDate >= from).ToList());
            }
            else
            {

                TempData["msg_tbl"] = "Searching Key Error";
                return View(_context.DishonoredCheques.ToList());
            }
        }
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public async Task<IActionResult> ViewDetail(int id = 0)
        {
            var userDistrict = _context.Branches.ToList();

            ViewBag.branchList = new SelectList(userDistrict, "Id", "Name");
            DishonoredCheque instanceType = new DishonoredCheque();
            if (id != 0)
            {
                instanceType = await _context.DishonoredCheques.Where(x => x.Id == id).FirstOrDefaultAsync();
            }

            return View(instanceType);
        }


        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


        [HttpPost]
        public IActionResult BusinessUnitReport(int InstanceId, DateTime? from, DateTime? to)
        {
            DateTime dateTime = DateTime.Now;
            var branchId = _httpContext.HttpContext.Session.GetInt32("UserBranch");
            var instances = _context.InstanceTypes.ToList();
            ViewBag.InstanceTypeList = new SelectList(instances, "Id", "Name");

            var report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == 0 && x.Branch == 0000).ToList();
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var stream = new MemoryStream();
            try
            {

                if (InstanceId != 0 && from != null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.Branch == branchId && x.IssueDate >= from && x.IssueDate <= to).ToList();
                }
                else if (InstanceId == 0 && from != null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.IssueDate >= from && x.Branch == branchId && x.IssueDate <= to).ToList();
                }
                else if (InstanceId != 0 && from == null && to == null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.Branch == branchId).ToList();
                }
                else if (InstanceId == 0 && from == null && to == null)
                {
                    report = _context.DishonoredCheques.Where(x => x.Branch == branchId).ToList();
                }
                else if (InstanceId == 0 && from == null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.IssueDate <= to && x.Branch == branchId).ToList();
                }
                else if (InstanceId == 0 && from != null && to == null)
                {
                    report = _context.DishonoredCheques.Where(x => x.IssueDate >= from && x.Branch == branchId).ToList();
                }
                else if (InstanceId != 0 && from == null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id >= InstanceId && x.IssueDate <= to && x.Branch == branchId).ToList();
                }
                else if (InstanceId != 0 && from != null && to == null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id >= InstanceId && x.IssueDate >= from && x.Branch == branchId).ToList();
                }
                else
                {
                    TempData["msg_tbl"] = "Searching Key Error";
                    return View(_context.DishonoredCheques.Where(x => x.Branch == branchId).ToList());
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {


                using (var xlpackage = new ExcelPackage(stream))
                {
                    // define workSheet for export
                    var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                    // defining Some styles if it is nessary
                    var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                    customStyle.Style.Font.UnderLine = true;
                    customStyle.Style.Font.Color.SetColor(Color.Gold);
                    // First row from Database
                    var startRow = 2;
                    var row = startRow;
                    /*
                    worksheet.Cells["A1"].Value = "ShareHolders Information";
                    using (var r = worksheet.Cells["A1:E1"])
                    {
                        r.Merge = true;
                        r.Style.Font.Color.SetColor(Color.Gray);
                        r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                    }*/
                    worksheet.Cells["A1"].Value = "Id";
                    worksheet.Cells["B1"].Value = "Branch";
                    worksheet.Cells["C1"].Value = "Home Branch";
                    worksheet.Cells["D1"].Value = "Sub City";
                    worksheet.Cells["E1"].Value = "Telephone";
                    worksheet.Cells["F1"].Value = "Instance Date";
                    worksheet.Cells["G1"].Value = "Registeration Date";
                    worksheet.Cells["H1"].Value = "Account Number";
                    worksheet.Cells["I1"].Value = "Full Name";
                    worksheet.Cells["J1"].Value = "IssueDate";
                    worksheet.Cells["K1"].Value = "Cheque Number";
                    worksheet.Cells["E1"].Value = "Beneficiary Name";
                    worksheet.Cells["E1"].Value = "Amount";
                    worksheet.Cells["E1"].Value = "Current Balance";
                    worksheet.Cells["E1"].Value = "Instance Type";
                    worksheet.Cells["E1"].Value = "Registered By";
                    worksheet.Cells["E1"].Value = "Authorized By";
                    worksheet.Cells["E1"].Value = "Remark";
                    worksheet.Cells["A1:E1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                    row = 2;
                    foreach (var share in report)
                    {
                        worksheet.Cells[row, 1].Value = share.Id;
                        worksheet.Cells[row, 2].Value = share.Branch;
                        worksheet.Cells[row, 3].Value = share.HomeBranch;
                        // worksheet.Cells[row, 4].Value = share.dateOfSubscription.ToString("dd-MM-yyyy");
                        worksheet.Cells[row, 4].Value = share.SubCity;
                        worksheet.Cells[row, 5].Value = share.Telephone;
                        worksheet.Cells[row, 6].Value = share.InstanceDate?.ToString("dd-MM-yyyy");
                        worksheet.Cells[row, 7].Value = share.RegisterationDate;
                        worksheet.Cells[row, 8].Value = share.AccountNumber;
                        worksheet.Cells[row, 9].Value = share.FullName;
                        worksheet.Cells[row, 10].Value = share.IssueDate;
                        worksheet.Cells[row, 11].Value = share.ChequeNumber;
                        worksheet.Cells[row, 12].Value = share.BeneficiaryName;
                        worksheet.Cells[row, 13].Value = share.AmountOfMoney;

                        worksheet.Cells[row, 14].Value = share.CurrentBalance;
                        worksheet.Cells[row, 15].Value = share.InstanceType;
                        worksheet.Cells[row, 16].Value = share.RegisteredBy;
                        worksheet.Cells[row, 17].Value = share.AuthorizedBy;
                        worksheet.Cells[row, 18].Value = share.Remark;
                        row++;
                    }
                    xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                    xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                    xlpackage.Save();

                }
                stream.Position = 0;


            }
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

        }
        [HttpPost]
        public IActionResult AllRegisteredReport(int InstanceId, DateTime? from, DateTime? to)
        {
            var instances = _context.InstanceTypes.ToList();
            ViewBag.InstanceTypeList = new SelectList(instances, "Id", "Name");

            var report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == 0 && x.Branch == 0000).ToList();
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var stream = new MemoryStream();
            try
            {


                if (InstanceId != 0 && from != null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.IssueDate >= from && x.IssueDate <= to).ToList();
                }
                else if (InstanceId == 0 && from != null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.IssueDate >= from && x.IssueDate <= to).ToList();
                }
                else if (InstanceId != 0 && from == null && to == null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId).ToList();
                }
                else if (InstanceId == 0 && from == null && to == null)
                {
                    report = _context.DishonoredCheques.ToList();
                }
                else if (InstanceId == 0 && from == null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.IssueDate <= to).ToList();
                }
                else if (InstanceId == 0 && from != null && to == null)
                {
                    report = _context.DishonoredCheques.Where(x => x.IssueDate >= from).ToList();
                }
                else if (InstanceId != 0 && from == null && to != null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.IssueDate <= to).ToList();
                }
                else if (InstanceId != 0 && from != null && to == null)
                {
                    report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId && x.IssueDate >= from).ToList();
                }
                else
                {

                    TempData["msg_tbl"] = "Searching Key Error";
                    return View(_context.DishonoredCheques.ToList());
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {


                using (var xlpackage = new ExcelPackage(stream))
                {
                    // define workSheet for export
                    var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                    // defining Some styles if it is nessary
                    var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                    customStyle.Style.Font.UnderLine = true;
                    customStyle.Style.Font.Color.SetColor(Color.Gold);
                    // First row from Database
                    var startRow = 2;
                    var row = startRow;
                    /*
                    worksheet.Cells["A1"].Value = "ShareHolders Information";
                    using (var r = worksheet.Cells["A1:E1"])
                    {
                        r.Merge = true;
                        r.Style.Font.Color.SetColor(Color.Gray);
                        r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                    }*/
                    worksheet.Cells["A1"].Value = "Id";
                    worksheet.Cells["B1"].Value = "Branch";
                    worksheet.Cells["C1"].Value = "Home Branch";
                    worksheet.Cells["D1"].Value = "Sub City";
                    worksheet.Cells["E1"].Value = "Telephone";
                    worksheet.Cells["F1"].Value = "Instance Date";
                    worksheet.Cells["G1"].Value = "Registeration Date";
                    worksheet.Cells["H1"].Value = "Account Number";
                    worksheet.Cells["I1"].Value = "Full Name";
                    worksheet.Cells["J1"].Value = "IssueDate";
                    worksheet.Cells["K1"].Value = "Cheque Number";
                    worksheet.Cells["L1"].Value = "Beneficiary Name";
                    worksheet.Cells["M1"].Value = "Amount";
                    worksheet.Cells["N1"].Value = "Current Balance";
                    worksheet.Cells["O1"].Value = "Registered By";
                    worksheet.Cells["P1"].Value = "Authorized By";
                    worksheet.Cells["Q1"].Value = "Remark";
                    worksheet.Cells["A1:E1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                    row = 2;
                    foreach (var share in report)
                    {
                        worksheet.Cells[row, 1].Value = share.Id;
                        worksheet.Cells[row, 2].Value = share.Branch;
                        worksheet.Cells[row, 3].Value = share.HomeBranch;
                        // worksheet.Cells[row, 4].Value = share.dateOfSubscription.ToString("dd-MM-yyyy");
                        worksheet.Cells[row, 4].Value = share.SubCity;
                        worksheet.Cells[row, 5].Value = share.Telephone;
                        worksheet.Cells[row, 6].Value = share.InstanceDate.ToString();
                        worksheet.Cells[row, 7].Value = share.RegisterationDate.ToString();
                        worksheet.Cells[row, 8].Value = share.AccountNumber;
                        worksheet.Cells[row, 9].Value = share.FullName;
                        worksheet.Cells[row, 10].Value = share.IssueDate.ToString();
                        worksheet.Cells[row, 11].Value = share.ChequeNumber;
                        worksheet.Cells[row, 12].Value = share.BeneficiaryName;
                        worksheet.Cells[row, 13].Value = share.AmountOfMoney;

                        worksheet.Cells[row, 14].Value = share.CurrentBalance;
                        worksheet.Cells[row, 16].Value = share.RegisteredBy;
                        worksheet.Cells[row, 17].Value = share.AuthorizedBy;
                        worksheet.Cells[row, 18].Value = share.Remark;
                        row++;
                    }
                    xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                    xlpackage.Save();

                }
                stream.Position = 0;


            }
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

        }
        [HttpPost]
        public IActionResult GenerateReport(int InstanceId, DateTime? from = null, DateTime? to = null)
        {
            TempData["msg"] = "rpt";

            TempData["InstanceId"] = InstanceId;
            TempData["from"] = from;
            TempData["to"] = to;
            return RedirectToAction("AllRegisteredRecords");
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        [HttpGet]
        public IActionResult DishonoredChequeRecords()
        {
            ViewBag.rptStatus = "All Records";

            var getBranchName = _context.Branches.ToList();
            ViewBag.Branchlist = new SelectList(getBranchName, "Id", "Name");

            var InstanceName = _context.InstanceTypes.ToList();
            ViewBag.Instancelist = new SelectList(InstanceName, "Id", "Name");
            return View(_context.DishonoredCheques.ToList());
        }
        [HttpPost]
        public IActionResult ActiveRecords(int BranchId, int Instance_Type, DateTime? dateFrom, DateTime? dateTo)
        {
            ViewBag.rptStatus = "active";
            var getBranchName = _context.Branches.ToList();
            ViewBag.Branchlist = new SelectList(getBranchName, "Id", "Name");

            var InstanceName = _context.InstanceTypes.ToList();
            ViewBag.Instancelist = new SelectList(InstanceName, "Id", "Name");

            if (dateTo != null && dateFrom != null && dateTo < dateFrom)
            {
                ViewBag.ErrorMessage = "* Date From can not exceeded Date To";
                return View("DishonoredChequeRecords", _context.DishonoredCheques.ToList());
            }
            else
            {
                if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && Instance_Type != 3 && c.InstanceType.Id != 3 && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && Instance_Type != 3 && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
                else
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                }
            }

        }
        [HttpPost]
        public IActionResult ExpiredRecords(int BranchId, int Instance_Type, DateTime? dateFrom, DateTime? dateTo)
        {
            ViewBag.rptStatus = "expire";
            var getBranchName = _context.Branches.ToList();
            ViewBag.Branchlist = new SelectList(getBranchName, "Id", "Name");

            var InstanceName = _context.InstanceTypes.ToList();
            ViewBag.Instancelist = new SelectList(InstanceName, "Id", "Name");

            if (dateTo != null && dateFrom != null && dateTo < dateFrom)
            {
                ViewBag.ErrorMessage = "* Date From can not exceeded Date To";
                return View("DishonoredChequeRecords", _context.DishonoredCheques.ToList());
            }
            else
            {
                if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.IsExpired == true).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.IsExpired == true).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.IsExpired == true).ToList());
                }
                else
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.IsExpired == true).ToList());
                }
            }

        }

        [HttpPost]
        public IActionResult BlockedRecords(int BranchId, int Instance_Type, DateTime? dateFrom, DateTime? dateTo)
        {
            ViewBag.rptStatus = "blocked";
            var getBranchName = _context.Branches.ToList();
            ViewBag.Branchlist = new SelectList(getBranchName, "Id", "Name");

            var InstanceName = _context.InstanceTypes.ToList();
            ViewBag.Instancelist = new SelectList(InstanceName, "Id", "Name");

            if (dateTo != null && dateFrom != null && dateTo < dateFrom)
            {
                ViewBag.ErrorMessage = "* Date From can not exceeded Date To";
                return View("DishonoredChequeRecords", _context.DishonoredCheques.ToList());
            }
            else
            {
                if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceTypeId == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                }
                else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.InstanceType.Id == 3).ToList());
                }
                else
                {
                    return View("DishonoredChequeRecords", _context.DishonoredCheques.Where(c => c.InstanceType.Id == 3).ToList());
                }
            }

        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////
        [HttpPost]
        public IActionResult ExportAllRecoreds(int btnActiveRecords, int btnExpiredRecords, int btnBlockedRecords, string rptStatus, int BranchId, int Instance_Type, DateTime? dateFrom, DateTime? dateTo)
        {

            var getBranchName = _context.Branches.ToList();
            ViewBag.Branchlist = new SelectList(getBranchName, "Id", "Name");

            var InstanceName = _context.InstanceTypes.ToList();
            ViewBag.Instancelist = new SelectList(InstanceName, "Id", "Name");
            var data = _context.DishonoredCheques.ToList();
            try
            {
                if (btnActiveRecords > 0)
                {
                    if (dateTo != null && dateFrom != null && dateTo < dateFrom)
                    {
                        ViewBag.ErrorMessage = "* Date From can not exceeded Date To";

                    }
                    else
                    {
                        if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && Instance_Type != 3 && c.InstanceType.Id != 3 && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && Instance_Type != 3 && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList());
                        }
                    }
                }
                else if (btnExpiredRecords > 0)
                {
                    if (dateTo != null && dateFrom != null && dateTo < dateFrom)
                    {
                        ViewBag.ErrorMessage = "* Date From can not exceeded Date To";
                        return View();
                    }
                    else
                    {
                        if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.IsExpired == true).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.IsExpired == true).ToList());
                        }
                    }
                }
                else if (btnBlockedRecords > 0)
                {
                    if (dateTo != null && dateFrom != null && dateTo < dateFrom)
                    {
                        ViewBag.ErrorMessage = "* Date From can not exceeded Date To";
                        return View();
                    }
                    else
                    {
                        if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                        {
                            //.InstanceType.Id == Instance_Type
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceTypeId == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.InstanceType.Id == 3).ToList());
                        }
                        else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                        {
                            return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == 3).ToList());
                        }
                    }
                }

            }
            catch (Exception ex)
            {

            }
            finally
            {

            }
            return View(_context.DishonoredCheques.ToList());
        }
        //


        [HttpPost]
        public IActionResult ExportAllRecoredsExcell(int BranchId, string rptStatus, int Instance_Type, DateTime? dateFrom, DateTime? dateTo, int btnActiveRecords = 0, int btnExpiredRecords = 0, int btnBlockedRecords = 0)
        {

            var getBranchName = _context.Branches.ToList();
            ViewBag.Branchlist = new SelectList(getBranchName, "Id", "Name");

            var InstanceName = _context.InstanceTypes.ToList();
            ViewBag.Instancelist = new SelectList(InstanceName, "Id", "Name");
            //
            var report = _context.DishonoredCheques.Where(x => x.InstanceType.Id == 0 && x.Branch == 0000).ToList();
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var stream = new MemoryStream();
            //
            if (rptStatus == "active")
            {
                if (dateTo != null && dateFrom != null && dateTo < dateFrom)
                {
                    ViewBag.ErrorMessage = "* Date From can not exceeded Date To";

                }
                else
                {
                    if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                    {

                        // var reportAll=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceTypeId == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false).ToList();
                        /*
                         * && Instance_Type != 3 && c.InstanceType.Id != 3 **/
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceTypeId == Instance_Type && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == false)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                        // 
                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                    {

                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceType.Id != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");


                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                    {

                        report=_context.DishonoredCheques.Where(c => c.InstanceTypeId == Instance_Type && Instance_Type != 3 && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && c.InstanceType.Id != 3).ToList();


                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceTypeId == Instance_Type && Instance_Type != 3 && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == false && DC.InstanceType.Id != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                        //////
                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceTypeId == Instance_Type && DC.RegisterationDate <= dateTo && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome
                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceTypeId == Instance_Type && DC.RegisterationDate >= dateFrom && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome
                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome
                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceTypeId == Instance_Type && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate <= dateTo && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate >= dateFrom &&DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.RegisterationDate <= dateTo && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Instance Date";
                            worksheet.Cells["G1"].Value = "Registeration Date";
                            worksheet.Cells["H1"].Value = "Account Number";
                            worksheet.Cells["I1"].Value = "Full Name";
                            worksheet.Cells["J1"].Value = "IssueDate";
                            worksheet.Cells["K1"].Value = "Cheque Number";
                            worksheet.Cells["E1"].Value = "Beneficiary Name";
                            worksheet.Cells["E1"].Value = "Amount";
                            worksheet.Cells["E1"].Value = "Current Balance";
                            worksheet.Cells["E1"].Value = "Instance Type";
                            worksheet.Cells["E1"].Value = "Registered By";
                            worksheet.Cells["E1"].Value = "Authorized By";
                            worksheet.Cells["E1"].Value = "Remark";
                            worksheet.Cells["A1:E1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in report)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                // worksheet.Cells[row, 4].Value = share.dateOfSubscription.ToString("dd-MM-yyyy");
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.InstanceDate?.ToString("dd-MM-yyyy");
                                worksheet.Cells[row, 7].Value = share.RegisterationDate;
                                worksheet.Cells[row, 8].Value = share.AccountNumber;
                                worksheet.Cells[row, 9].Value = share.FullName;
                                worksheet.Cells[row, 10].Value = share.IssueDate;
                                worksheet.Cells[row, 11].Value = share.ChequeNumber;
                                worksheet.Cells[row, 12].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 13].Value = share.AmountOfMoney;

                                worksheet.Cells[row, 14].Value = share.CurrentBalance;
                                worksheet.Cells[row, 15].Value = share.InstanceType;
                                worksheet.Cells[row, 16].Value = share.RegisteredBy;
                                worksheet.Cells[row, 17].Value = share.AuthorizedBy;
                                worksheet.Cells[row, 18].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                            // }
                            stream.Position = 0;
                        }
                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate <= dateTo && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceTypeId == Instance_Type && DC.RegisterationDate >= dateFrom && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceTypeId != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate >= dateFrom && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceType.Id != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceTypeId == Instance_Type && DC.IsExpired == false && Instance_Type != 3 && DC.InstanceType.Id != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.IsExpired == false && Instance_Type != 3 && c.InstanceType.Id != 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.IsExpired == false && Instance_Type != 3 && DC.InstanceType.Id != 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");
                        //wellcome finished here 
                    }
                }
            }
            /////
            else if (rptStatus== "expire")
            {
                if (dateTo != null && dateFrom != null && dateTo < dateFrom)
                {
                    ViewBag.ErrorMessage = "* Date From can not exceeded Date To";
                    // return View();
                }
                else
                {
                    if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                    {

                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();


                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == Instance_Type && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == true) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                        // 
                        //

                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == true) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == true) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where DC.Branch == BranchId && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.IsExpired == true
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceTypeId == Instance_Type && DC.RegisterationDate <= dateTo && DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == Instance_Type && DC.RegisterationDate >= dateFrom && DC.IsExpired == true) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate >= dateFrom && DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.RegisterationDate <= dateTo && DC.IsExpired == true) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.RegisterationDate <= dateTo && DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate <= dateTo && DC.IsExpired == true) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate >= dateFrom && DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == Instance_Type && DC.IsExpired == true) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.IsExpired == true).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.IsExpired == true)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                }
            }

            //wellcome
            else if (rptStatus=="blocked")
            {
                if (dateTo != null && dateFrom != null && dateTo < dateFrom)
                {
                    ViewBag.ErrorMessage = "* Date From can not exceeded Date To";
                    //return View();
                }
                else
                {
                    if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                    {

                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == Instance_Type && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.InstanceType.Id == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.InstanceType.Id == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.InstanceType.Id == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate >= dateFrom && DC.RegisterationDate <= dateTo && DC.InstanceType.Id == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == Instance_Type && DC.RegisterationDate <= dateTo && DC.InstanceType.Id == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == Instance_Type && DC.RegisterationDate >= dateFrom && DC.InstanceType.Id == 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.InstanceType.Id == 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");


                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate <= dateTo && DC.InstanceType.Id == 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.RegisterationDate >= dateFrom && DC.InstanceType.Id == 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == Instance_Type && DC.RegisterationDate <= dateTo && DC.InstanceTypeId == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type != 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceTypeId == Instance_Type && DC.RegisterationDate >= dateFrom && DC.InstanceTypeId == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom == null && dateTo != null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate <= dateTo && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate <= dateTo && DC.InstanceTypeId == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type == 0 && dateFrom != null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.RegisterationDate >= dateFrom && c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.RegisterationDate >= dateFrom && DC.InstanceTypeId == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId != 0 && Instance_Type != 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type && c.InstanceType.Id == 3).ToList();

                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.Branch == BranchId && DC.InstanceType.Id == Instance_Type && DC.InstanceTypeId == 3) join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id

                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                    else if (BranchId == 0 && Instance_Type == 0 && dateFrom == null && dateTo == null)
                    {
                        report=_context.DishonoredCheques.Where(c => c.InstanceType.Id == 3).ToList();
                        var reportAll = (from DC in _context.DishonoredCheques
                                         join Br in _context.Branches on DC.Branch equals Br.Id
                                         //where DC.Branch == branchId
                                         where (DC.InstanceType.Id == 3)
                                         join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id


                                         select new DishonoredChequeView
                                         {
                                             Id = DC.Id,
                                             Branch = Br.Name,
                                             HomeBranch = DC.HomeBranch,
                                             AccountNumber = DC.AccountNumber,
                                             FullName = DC.FullName,
                                             SubCity = DC.SubCity,
                                             Telephone = DC.Telephone,
                                             ChequeNumber = DC.ChequeNumber,
                                             BeneficiaryName = DC.BeneficiaryName,
                                             AmountOfMoney = DC.AmountOfMoney,
                                             CurrentBalance = DC.CurrentBalance,
                                             IssueDate = DC.IssueDate,
                                             RegisteredBy = DC.RegisteredBy,
                                             RegisterationDate = DC.RegisterationDate,
                                             Remark=DC.Remark
                                         }).ToList();
                        using (var xlpackage = new ExcelPackage(stream))
                        {
                            // define workSheet for export
                            var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                            // defining Some styles if it is nessary
                            var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                            customStyle.Style.Font.UnderLine = true;
                            customStyle.Style.Font.Color.SetColor(Color.Gold);
                            // First row from Database
                            var startRow = 2;
                            var row = startRow;
                            /*
                            worksheet.Cells["A1"].Value = "ShareHolders Information";
                            using (var r = worksheet.Cells["A1:E1"])
                            {
                                r.Merge = true;
                                r.Style.Font.Color.SetColor(Color.Gray);
                                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            }*/
                            worksheet.Cells["A1"].Value = "Id";
                            worksheet.Cells["B1"].Value = "Branch";
                            worksheet.Cells["C1"].Value = "Home Branch";
                            worksheet.Cells["D1"].Value = "Sub City";
                            worksheet.Cells["E1"].Value = "Telephone";
                            worksheet.Cells["F1"].Value = "Registeration Date";
                            worksheet.Cells["G1"].Value = "Account Number";
                            worksheet.Cells["H1"].Value = "Full Name";
                            worksheet.Cells["I1"].Value = "IssueDate";
                            worksheet.Cells["J1"].Value = "Cheque Number";
                            worksheet.Cells["K1"].Value = "Beneficiary Name";
                            worksheet.Cells["L1"].Value = "Amount";
                            worksheet.Cells["M1"].Value = "Current Balance";
                            worksheet.Cells["N1"].Value = "Remark";
                            worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                            row = 2;
                            foreach (var share in reportAll)
                            {
                                worksheet.Cells[row, 1].Value = share.Id;
                                worksheet.Cells[row, 2].Value = share.Branch;
                                worksheet.Cells[row, 3].Value = share.HomeBranch;
                                worksheet.Cells[row, 4].Value = share.SubCity;
                                worksheet.Cells[row, 5].Value = share.Telephone;
                                worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 7].Value = share.AccountNumber;
                                worksheet.Cells[row, 8].Value = share.FullName;
                                worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                                worksheet.Cells[row, 10].Value = share.ChequeNumber;
                                worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                                worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                                worksheet.Cells[row, 13].Value = share.CurrentBalance;
                                worksheet.Cells[row, 14].Value = share.Remark;
                                row++;
                            }
                            xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                            xlpackage.Workbook.Properties.Author = "solomon.sefiw";
                            xlpackage.Save();

                        }

                        stream.Position = 0;

                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

                    }
                }
            }
            //
            var reportNo = (from DC in _context.DishonoredCheques
                            join Br in _context.Branches on DC.Branch equals Br.Id
                            //where DC.Branch == branchId
                            //where (DC.InstanceType.Id == 3)
                            join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id
                            select new DishonoredChequeView
                            {
                                Id = DC.Id,
                                Branch = Br.Name,
                                HomeBranch = DC.HomeBranch,
                                AccountNumber = DC.AccountNumber,
                                FullName = DC.FullName,
                                SubCity = DC.SubCity,
                                Telephone = DC.Telephone,
                                ChequeNumber = DC.ChequeNumber,
                                BeneficiaryName = DC.BeneficiaryName,
                                AmountOfMoney = DC.AmountOfMoney,
                                CurrentBalance = DC.CurrentBalance,
                                IssueDate = DC.IssueDate,
                                RegisteredBy = DC.RegisteredBy,
                                RegisterationDate = DC.RegisterationDate,
                                Remark=DC.Remark
                            }).ToList();
            using (var xlpackage = new ExcelPackage(stream))
            {
                // define workSheet for export
                var worksheet = xlpackage.Workbook.Worksheets.Add("DishonoredCheques");
                // defining Some styles if it is nessary
                var customStyle = xlpackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                customStyle.Style.Font.UnderLine = true;
                customStyle.Style.Font.Color.SetColor(Color.Gold);
                // First row from Database
                var startRow = 2;
                var row = startRow;
                /*
                worksheet.Cells["A1"].Value = "ShareHolders Information";
                using (var r = worksheet.Cells["A1:E1"])
                {
                    r.Merge = true;
                    r.Style.Font.Color.SetColor(Color.Gray);
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                }*/
                worksheet.Cells["A1"].Value = "Id";
                worksheet.Cells["B1"].Value = "Branch";
                worksheet.Cells["C1"].Value = "Home Branch";
                worksheet.Cells["D1"].Value = "Sub City";
                worksheet.Cells["E1"].Value = "Telephone";
                worksheet.Cells["F1"].Value = "Registeration Date";
                worksheet.Cells["G1"].Value = "Account Number";
                worksheet.Cells["H1"].Value = "Full Name";
                worksheet.Cells["I1"].Value = "IssueDate";
                worksheet.Cells["J1"].Value = "Cheque Number";
                worksheet.Cells["K1"].Value = "Beneficiary Name";
                worksheet.Cells["L1"].Value = "Amount";
                worksheet.Cells["M1"].Value = "Current Balance";
                worksheet.Cells["N1"].Value = "Remark";
                worksheet.Cells["A1:N1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(224, 224, 224));
                row = 2;
                foreach (var share in reportNo)
                {
                    worksheet.Cells[row, 1].Value = share.Id;
                    worksheet.Cells[row, 2].Value = share.Branch;
                    worksheet.Cells[row, 3].Value = share.HomeBranch;
                    worksheet.Cells[row, 4].Value = share.SubCity;
                    worksheet.Cells[row, 5].Value = share.Telephone;
                    worksheet.Cells[row, 6].Value = share.RegisterationDate.ToString("yyyy-MM-dd");
                    worksheet.Cells[row, 7].Value = share.AccountNumber;
                    worksheet.Cells[row, 8].Value = share.FullName;
                    worksheet.Cells[row, 9].Value = share.IssueDate.ToString("yyyy-MM-dd");
                    worksheet.Cells[row, 10].Value = share.ChequeNumber;
                    worksheet.Cells[row, 11].Value = share.BeneficiaryName;
                    worksheet.Cells[row, 12].Value = share.AmountOfMoney;
                    worksheet.Cells[row, 13].Value = share.CurrentBalance;
                    worksheet.Cells[row, 14].Value = share.Remark;
                    row++;
                }
                xlpackage.Workbook.Properties.Title = "Dishonored Cheque List";
                xlpackage.Save();

            }

            stream.Position = 0;

            //
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DCMS_Report.xlsx");

        }
        //

        // return View(_context.DishonoredCheques.ToList());

        //}
        //
    }
    //

}
