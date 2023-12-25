using DCMS.Models;
using Microsoft.AspNetCore.Mvc;
using DCMS.Data;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System;
using System.Security.Policy;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using Microsoft.DotNet.Scaffolding.Shared.Messaging;
using DCMS.Models.Setup;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace DCMS.Controllers
{
    public class DishonoredChequesController : Controller
    {
        private DCMSDbContext _context;
        private readonly IHttpContextAccessor _httpContext;
        public string alert = "";

        public DishonoredChequesController(DCMSDbContext context, IHttpContextAccessor session)
        {
            _context = context;
            _httpContext = session;

        }
        public async Task<IActionResult> Index()
        {

      /*     using (HttpClient client = new HttpClient())
            {
                string url = "http://localhost:5113/DishonoredCheques/Isexpired";
               HttpResponseMessage response = await client.GetAsync(url);
               if (response.IsSuccessStatusCode)
               {
                   List<User> user = _context.Users.ToList();

                   return View();
               }
                else
                {
                   // Handle error
               }
            }**/
            //  return RedirectToAction("Users", "AccountController");
            //var user = _context.Users.ToList();
            return View();
        }
        //public string Isexpired()
        //{
        //    List<DishonoredCheque> cheque = _context.DishonoredCheques.Where(x => x.RegisterationDate 
        //    <= DateTime.Today).ToList();
        //    foreach (var d in cheque)
        //    {
        //        d.IsExpired = true;
        //        _context.DishonoredCheques.Update(d);
        //        _context.SaveChanges(true);

        //    }
        //    return "";
        //}
        public IActionResult RegisterCheque(DishonoredCheque cheques)
        {
            var userName = _httpContext.HttpContext.Session.GetString("UserName");
            if (cheques.CurrentBalance >= cheques.AmountOfMoney)
            {
                TempData["Warning"] = "Current balance must be less than Amount of money";
                return RedirectToAction("InsertCheques");
            }
            cheques.Branch = (int)_httpContext.HttpContext.Session.GetInt32("UserBranch");
            cheques.IsExpired = false;
            cheques.ExpiryDate = DateTime.Today.AddDays(45);
            cheques.statusId = 1;
            cheques.Status = (Status)1;
            cheques.InstanceDate = DateTime.Now;
            cheques.RegisteredBy = userName;
            _context.DishonoredCheques.Add(cheques);
            _context.SaveChanges();
            TempData["AlertMessage"] = "Cheque Registered successfully";
            return RedirectToAction("InsertCheques");
        }

        [HttpGet]
        public IActionResult Instancecheckvariable(string id)
        {
            string redirectUrl = Url.Action("InsertCheques", "DishonoredCheques");
            ViewBag.Account = id;

            int Count = _context.DishonoredCheques.Where(e => e.AccountNumber == id).Count();
            if (Count == 0)
            {
                ViewBag.Text = "first Instance";
                ViewBag.Value = 1;
            }
            if (Count == 1)
            {
                ViewBag.Text = "second Instance";
                ViewBag.Value = 2;
            }
            if (Count == 2)
            {
                ViewBag.Text = "third Instance";
                ViewBag.Value = 3;
            }
            redirectUrl += "?ViewBag.Text=" + ViewBag.Text;
            redirectUrl += "&ViewBag.Value=" + ViewBag.Value;
            redirectUrl += "&ViewBag.Account=" + ViewBag.Account;
            return Json(new { redirectUrl });
        }
        public IActionResult Instancechecking()
        {
            return View();
        }
        //
        [HttpGet]
        public async Task<IActionResult> InsertCheques()
        {

            return View();
        }
        //instnce check

        public IActionResult DishonoredCheques()
        {
            var branchId = _httpContext.HttpContext.Session.GetInt32("UserBranch");
           // var BranchCode = _context.Branches.Find(branchIdCode);
            //var branchId = BranchCode.Code;
            if (branchId !=3)
            {
                var Cheque = new List<DishonoredChequeView>();
                Cheque = (from DC in _context.DishonoredCheques
                          join Br in _context.Branches on DC.Branch equals Br.Id
                          where DC.Branch == branchId
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
                              InstanceType = Ins.Name,
                              IssueDate = DC.IssueDate,
                              RegisteredBy = DC.RegisteredBy,
                              RegisterationDate = DC.RegisterationDate
                          }).ToList();

              /*  int pageSize = 4; // Number of items per page

                // Your logic to retrieve the full data set
                var fullDataSet = Cheque; // Your logic to get data;

                // Calculate the total number of pages
                int totalItems = Cheque.Count();
                int totalPages = (int)Math.Ceiling((double)totalItems / pageSize);

                // Determine the current page
                int currentPage = page ?? 1; // If page is null, default to page 1

                // Skip and take to get the current page data
                var currentPageData = Cheque.Skip((currentPage - 1) * pageSize).Take(pageSize);

                // Pass the current page data, current page, and total pages to the view
                ViewBag.CurrentPageData = currentPageData;
                ViewBag.CurrentPage = currentPage;
                ViewBag.TotalPages = totalPages;*/
                return View(Cheque);
            }
            else
            {

                var Cheque = new List<DishonoredChequeView>();
                Cheque = (from DC in _context.DishonoredCheques
                          join Br in _context.Branches on DC.Branch equals Br.Id
                         //where DC.Branch == 1//branchId
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
                              InstanceType = Ins.Name,
                              IssueDate = DC.IssueDate,
                              RegisteredBy = DC.RegisteredBy,
                              RegisterationDate = DC.RegisterationDate
                          }).ToList();

               /* int pageSize = 4; // Number of items per page

                // Your logic to retrieve the full data set
                var fullDataSet = Cheque; // Your logic to get data;

                // Calculate the total number of pages
                int totalItems = Cheque.Count();
                int totalPages = (int)Math.Ceiling((double)totalItems / pageSize);

                // Determine the current page
                int currentPage = page ?? 1; // If page is null, default to page 1

                // Skip and take to get the current page data
                var currentPageData = Cheque.Skip((currentPage - 1) * pageSize).Take(pageSize);

                // Pass the current page data, current page, and total pages to the view
                ViewBag.CurrentPageData = currentPageData;
                ViewBag.CurrentPage = currentPage;
                ViewBag.TotalPages = totalPages;*/
                return View(Cheque);
            }
        }

        //for rejected
        public IActionResult Rejected()
        {

            var branchId = _httpContext.HttpContext.Session.GetInt32("UserBranch");
            if (branchId == 3)
            {
                var Cheque = new List<DishonoredChequeView>();
                Cheque = (from DC in _context.DishonoredCheques
                          join Br in _context.Branches on DC.Branch equals Br.Id
                          // where DC.statusId == 2
                          where DC.Status == (Status)2
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
                              InstanceType = Ins.Name,
                              IssueDate = DC.IssueDate,
                              RegisteredBy = DC.RegisteredBy,
                              RegisterationDate = DC.RegisterationDate
                          }).ToList();
                return View(Cheque);
            }
            else
            {

                var Cheque = new List<DishonoredChequeView>();
                Cheque = (from DC in _context.DishonoredCheques
                          join Br in _context.Branches on DC.Branch equals Br.Id
                          where DC.Status == (Status)2 && DC.Branch == branchId
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
                              InstanceType = Ins.Name,
                              IssueDate = DC.IssueDate,
                              RegisteredBy = DC.RegisteredBy,
                              RegisterationDate = DC.RegisterationDate
                          }).ToList();
                return View(Cheque);
            }

        }
        //Rejected update
        public IActionResult DishonoredRejected(DishonoredCheque cheques)
        {
            cheques.statusId = 1;
            cheques.Status = (Status)1;
            cheques.InstanceDate = DateTime.Today;
            _context.DishonoredCheques.Update(cheques);
            _context.SaveChanges(true);
            TempData["AlertMessage"] = "Rejected Cheque Reviewed successfully";
            return RedirectToAction("UpdateDishonored");
        }
        //
        //update Dishonored
        public IActionResult DishonoreUpdation(DishonoredCheque cheques)
        {
            cheques.statusId = 1;
            cheques.Status = (Status)1;
            cheques.InstanceDate = DateTime.Today;
            _context.DishonoredCheques.Update(cheques);
            _context.SaveChanges(true);
            TempData["AlertMessage"] = "successfully update";
            return RedirectToAction("UpdateDishonoredCheque");
        }
        //return the Dishonered cheque which is onprogress
        public IActionResult Authorize()
        {
       
            var branchId = _httpContext.HttpContext.Session.GetInt32("UserBranch");
            if (branchId == 3)
            {
                var Cheque = new List<DishonoredChequeView>();
                Cheque = (from DC in _context.DishonoredCheques
                          join Br in _context.Branches on DC.Branch equals Br.Id
                          join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id
                          where DC.Status == (Status)1
                          select new DishonoredChequeView
                          {
                              Id = DC.Id,
                              Branch = Br.Name,
                              HomeBranch = DC.HomeBranch,
                              AccountNumber = DC.AccountNumber,
                              FullName = DC.FullName,
                              SubCity=DC.SubCity,
                              Telephone=DC.Telephone,
                              ChequeNumber = DC.ChequeNumber,
                              BeneficiaryName = DC.BeneficiaryName,
                              AmountOfMoney = DC.AmountOfMoney,
                              CurrentBalance = DC.CurrentBalance,
                              InstanceType=Ins.Name,
                              IssueDate=DC.IssueDate,
                              RegisteredBy=DC.RegisteredBy,
                              RegisterationDate = DC.RegisterationDate
                          }).ToList();
                return View(Cheque);

            }
            else
            {
                var Cheque = new List<DishonoredChequeView>();
                Cheque = (from DC in _context.DishonoredCheques
                          join Br in _context.Branches on DC.Branch equals Br.Id
                          join Ins in _context.InstanceTypes on DC.InstanceTypeId equals Ins.Id
                          where DC.Status == (Status)1 && DC.Branch == branchId && DC.InstanceTypeId != 3
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
                              InstanceType = Ins.Name,
                              IssueDate = DC.IssueDate,
                              RegisteredBy = DC.RegisteredBy,
                              RegisterationDate = DC.RegisterationDate
                          }).ToList();
                return View(Cheque);

            }
            //return View(Cheque);
        }
        //Render the the Authorization page
        public IActionResult AuthorizeDishonored(int id)
        {
            DishonoredCheque cheque = _context.DishonoredCheques.Find(id);

            return View(cheque);
        }
        //post back after authorized
        public IActionResult AuthorizeDishonoredBack()
        {
            @TempData["AlertMessage"] = "Authorized Successfully";

            return View();
        }

   
        //postback after rejecteed
        public IActionResult AuthorizeDishonoredRejectedBack()
        {
            @TempData["AlertMessage"] = "Rejected Successfully";
            return View();
        }
        //Render Dishonored Updation page
        public IActionResult UpdateDishonoredCheque(int id)
        {
            DishonoredCheque cheque = _context.DishonoredCheques.Find(id);

            return View(cheque);
        }

        [HttpPost]
        public IActionResult AuthorizeAccept(int id)
        {
            var userName = _httpContext.HttpContext.Session.GetString("UserName");

            if (id == 0)
            {
                string referrerUrl = Request.Headers["Referer"].ToString();
                string redirectUrl = $"{referrerUrl}?Empty=NoIdselected";
                return Redirect(redirectUrl);
            }
            DishonoredCheque cheques = _context.DishonoredCheques.Find(id);
            cheques.statusId = 3;
            cheques.Status = (Status)3;
            cheques.AuthorizationDate = DateTime.Now;
            cheques.AuthorizedBy = userName;
            _context.DishonoredCheques.Update(cheques);
             _context.SaveChanges();
              return RedirectToAction("AuthorizeDishonoredBack");
         
        }
        //toAuthorize
        [HttpPost]
        public async Task<IActionResult> AuthorizeReject(int id)
        {
            var userName = _httpContext.HttpContext.Session.GetString("UserName");

            if (id == 0)
            {
                string referrerUrl = Request.Headers["Referer"].ToString();
                string redirectUrl = $"{referrerUrl}?Empty=HelloFromGoBack";
                return Redirect(redirectUrl);
            }
            DishonoredCheque cheques = _context.DishonoredCheques.Find(id);
            // cheques.Status = (Models.Setup.Status)2;
            cheques.statusId = 2;
            cheques.Status = (Status)2;
            cheques.AuthorizationDate = DateTime.Now;
            cheques.AuthorizedBy = userName;
           _context.DishonoredCheques.Update(cheques);
           _context.SaveChanges();
          return RedirectToAction("AuthorizeDishonoredRejectedBack");
         
        }

        //
        public IActionResult UpdateDishonored(int id)
        {
            DishonoredCheque cheque = _context.DishonoredCheques.Find(id);
            return View(cheque);
        }
        public IActionResult DishonoredUpdate(DishonoredCheque cheques)
        {
                _context.DishonoredCheques.Update(cheques);
                _context.SaveChanges(true);
                TempData["AlertMessage"] = "Cheque Updated successfully";
                return RedirectToAction("Rejected");
    
        }

        //////////////////
        public IActionResult SearchDishonoredCheque(string Account_Number, int BranchId, int Instance_Type, int btnSearch, int btnClear)
        {
            var getBranchName = _context.Branches.ToList();
            ViewBag.Branchlist = new SelectList(getBranchName, "Id", "Name");

            var InstanceName = _context.InstanceTypes.ToList();
            ViewBag.Instancelist = new SelectList(InstanceName, "Id", "Name");
            if (btnSearch > 0)
            {
                if (Account_Number != null && BranchId != 0 && Instance_Type != 0)
                {
                    return View(_context.DishonoredCheques.Where(c => c.AccountNumber == Account_Number && c.Branch == BranchId && c.InstanceType.Id == Instance_Type).ToList());
                }
                else if (Account_Number == null && BranchId == 0 && Instance_Type == 0)
                {
                    return View(_context.DishonoredCheques.ToList());
                }
                else if (Account_Number == null && BranchId == 0 && Instance_Type != 0)
                {
                    return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type).ToList());
                }
                else if (Account_Number == null && BranchId != 0 && Instance_Type == 0)
                {
                    return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId).ToList());
                }
                else if (Account_Number == null && BranchId != 0 && Instance_Type != 0)
                {
                    return View(_context.DishonoredCheques.Where(c => c.Branch == BranchId && c.InstanceType.Id == Instance_Type).ToList());
                }
                else if (Account_Number != null && BranchId == 0 && Instance_Type == 0)
                {
                    return View(_context.DishonoredCheques.Where(c => c.AccountNumber == Account_Number).ToList());
                }
                else if (Account_Number != null && BranchId == 0 && Instance_Type != 0)
                {
                    return View(_context.DishonoredCheques.Where(c => c.InstanceType.Id == Instance_Type && c.AccountNumber == Account_Number).ToList());
                }
                else if (Account_Number != null && BranchId != 0 && Instance_Type == 0)
                {
                    return View(_context.DishonoredCheques.Where(c => c.AccountNumber == Account_Number && c.Branch == BranchId).ToList());
                }
            }
            return View(_context.DishonoredCheques.ToList());
        }



    }
}
