
using DCMS.Data;
using DCMS.Help;
using DCMS.Models.Setup;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using static DCMS.Help.Helper;

namespace DCMS.Controllers
{

    [CheckSessionIsAvailable]
    [NoDirectAccess]
    public class SetupController : Controller
    {
        private readonly DCMSDbContext _context;
        private readonly IHttpContextAccessor _httpContext;
        public SetupController(DCMSDbContext context, IHttpContextAccessor httpContext)
        {
            _context = context;
            _httpContext = httpContext;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public async Task<IActionResult> TypeofInstance()
        {
            return View(await _context.InstanceTypes.ToListAsync());
        }

        public async Task<IActionResult> AddOrEditInstanceType(int id = 0)
        {
            InstanceType instanceType  = new InstanceType();
            if (id != 0)
            {
                instanceType = await _context.InstanceTypes.Where(x => x.Id == id).FirstOrDefaultAsync();
            }

            return View(instanceType);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> AddOrEditInstanceType(int id, InstanceType instanceType)
        {
            if (id == 0)
            {
                await _context.InstanceTypes.AddAsync(instanceType);
                await _context.SaveChangesAsync();
                TempData["msg"] = "Added Successfully";
            }
            else
            {
                try
                {
                    _context.Update(instanceType);
                    await _context.SaveChangesAsync();
                    TempData["msg"] = "Updated Successfully";
                }
                catch (DbUpdateConcurrencyException)
                {
                    throw;
                }
            }
            return RedirectToAction("TypeofInstance", "Setup");
        }
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public async Task<IActionResult> Branch()
        {
            var userDistrict = _context.Districts.ToList();

            ViewBag.districtList = new SelectList(userDistrict, "Id", "Name");
            return View(await _context.Branches.ToListAsync());
        }

        public async Task<IActionResult> AddOrEditBranch(int id = 0)
        {
            var userDistrict = _context.Districts.ToList();

            ViewBag.districtList = new SelectList(userDistrict, "Id", "Name");
            Branch branch = new Branch();
            if (id != 0)
            {
                branch = await _context.Branches.FirstOrDefaultAsync();
            }

            return View(branch);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> AddOrEditBranch(int id, Branch branch)
        {

            var userDistrict = _context.Districts.ToList();

            ViewBag.districtList = new SelectList(userDistrict , "Id", "Name");

            if (id == 0)
            {
                await _context.Branches.AddAsync(branch);
                await _context.SaveChangesAsync();
                TempData["msg"] = "Added Successfully";
            }
            else
            {
                try
                {
                    _context.Update(branch);
                    await _context.SaveChangesAsync();
                    TempData["msg"] = "Updated Successfully";
                }
                catch (DbUpdateConcurrencyException)
                {
                    throw;
                }
            }
            return RedirectToAction("Branch", "Setup");
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////// 
        public async Task<IActionResult> District()
        {
            return View(await _context.Districts.ToListAsync());
        }

        public async Task<IActionResult> AddOrEditDistrict(int id = 0)
        {
            District district = new District();
            if (id != 0)
            {
                district = await _context.Districts.FirstOrDefaultAsync();
            }

            return View(district);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> AddOrEditDistrict(int id, District district)
        {
            if (id == 0)
            {
                await _context.Districts.AddAsync(district);
                await _context.SaveChangesAsync();
                TempData["msg"] = "Added Successfully";
            }
            else
            {
                try
                {
                    _context.Update(district);
                    await _context.SaveChangesAsync();
                    TempData["msg"] = "Updated Successfully";
                }
                catch (DbUpdateConcurrencyException)
                {
                    throw;
                }
            }
            return RedirectToAction("District", "Setup");
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////// 


    }
}
