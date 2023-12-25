
using DCMS.Data;
using DCMS.Models.Setup;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;


namespace DCMS.Controllers
{

    public class AccountController : Microsoft.AspNetCore.Mvc.Controller
    {
        private readonly DCMSDbContext _context;
        private readonly IHttpContextAccessor _httpContext;

        public AccountController(DCMSDbContext context, IHttpContextAccessor httpContext)
        {
            _context = context;
            _httpContext = httpContext;
        }




        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public async Task<IActionResult> Users()
        {
            return View(await _context.Users.ToListAsync());
        }

        public async Task<IActionResult> AddOrEditUser(int id = 0)
        {
            var userRole = _context.Roles.ToList();
            var userBranch = _context.Branches.ToList();

            ViewBag.roleList = new SelectList(userRole, "Id", "Name");
            ViewBag.branchList = new SelectList(userBranch, "Id", "Name");
            User user = new User();
            if (id != 0)
            {
                user = await _context.Users.Where(x => x.Id == id).FirstOrDefaultAsync();
                user.Password = user.Password.ToString();
            }

            return View(user);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> AddOrEditUser(int id, User user)
        {
            if (id == 0)
            {
                var user2 = new User
                {
                    FullName = user.FullName,
                    UserName = user.UserName,
                    Email = user.Email,
                    Password = user.Password,
                    BranchId = user.BranchId,
                    RoleId = user.RoleId
                };

                string pass = user.Password;
                string encryptPass = EncodePasswordToBase64(pass);
                user2.Password = encryptPass;
               await _context.Users.AddAsync(user2);
               await _context.SaveChangesAsync();
                TempData["msg"] = "Added Successfully";
            }
            else
            {
                try
                {
                    string pass = user.Password;
                    string encryptPass = EncodePasswordToBase64(pass);
                    user.Password = encryptPass;
                    _context.Users.Update(user);
                   await _context.SaveChangesAsync();
                    TempData["msg"] = "Updated Successfully";
                }
                catch (DbUpdateConcurrencyException)
                {
                    throw;
                }
            }
            return RedirectToAction("Users", "Account");
        }
        private User GetUserByID(int id)
        {
            var user = _context.Users.FirstOrDefault(x => x.Id == id);
            return user;
        }
        public async Task<IActionResult> ChangePassword()
        {
            var user1 = _httpContext.HttpContext.Session.GetString("UserName");
            var userdetail =  _context.Users.Where(x => x.UserName == user1).FirstOrDefault();
            return View(userdetail);
        }
        [HttpPost]
        public async Task<IActionResult> ChangePassword(User user)
        {
          //if(ModelState.IsValid)
          //  {
                string passwd = EncodePasswordToBase64(user.Password);
                string newpasswd = EncodePasswordToBase64(user.NewPassword);
                string confiermpasswd = EncodePasswordToBase64(user.ConfiermPassword);

                var userupdate = _context.Users.Where(x => x.Id == user.Id).FirstOrDefault();
                if(userupdate.Password  == passwd)
                {
                    if(newpasswd == confiermpasswd) {
                    userupdate.Password = newpasswd;
                        _context.Users.Update(userupdate);
                         _context.SaveChanges();
                        return RedirectToAction("Login", "Account");
                    }
                    else
                    {
                        TempData["msg"] = " Confirm Password Not Match!";
                        return View();
                    }

                }
                else
                {
                    TempData["msg"] = " Current Password Not Correct!";
                    return View();
                }



            //}
     
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////






        public IActionResult Login()
        {
            _httpContext.HttpContext.Session.Clear();
            return View();
        }

        [HttpPost]
        public IActionResult LoginAction(Login login)
        {

            if (ModelState.IsValid)
            {
                string password = EncodePasswordToBase64(login.Password);
                var users = _context.Users.Where(a => a.UserName.Equals(login.UserName) && a.Password.Equals(password) && a.viewStatus == true).FirstOrDefault();
                if (users != null)
                {
                    var user = _context.Users.Where(x => x.UserName == login.UserName).FirstOrDefault();
                    _httpContext.HttpContext.Session.SetString("UserName", user.UserName);
                    _httpContext.HttpContext.Session.SetInt32("UserRole", user.RoleId);
                    _httpContext.HttpContext.Session.SetInt32("UserBranch", user.BranchId);
                    if (user.RoleId ==3 )
                    {
                        TempData["msg"] = "Access Guaranteed !! ";
                        return RedirectToAction("Index", "DishonoredCheques");
                    }
                    else if (user.RoleId == 1 || user.RoleId == 2)
                    {
                        TempData["msg"] = "Access Guaranteed !! ";
                        return RedirectToAction("Index", "DishonoredCheques");
                    }
                    else
                    {
                       // TempData["msg"] = "Access Denied !! ";
                        //return RedirectToAction(nameof(Login));
                    }
                }
                //
                else
                {
                    TempData["Error"] = "UserName or Password Is Incorrect";
                    return RedirectToAction(nameof(Login));
                }
                //

            }
            TempData["msg"] = "Access Denied !! ";
            return RedirectToAction(nameof(Login));
        }

        public IActionResult Logout()
        {
            string username = HttpContext.Session.GetString("UserName");
            if (!string.IsNullOrEmpty(username))
            {
                _httpContext.HttpContext.Session.Clear();
                return RedirectToAction("Login", "Account");
            }
            return RedirectToAction("Login", "Account");

        }
        public IActionResult LoginView()
        {
            return RedirectToAction("Login", "Account");
        }
        //this function Convert to Encord your Password
        public static string EncodePasswordToBase64(string password)
        {
            try
            {
                byte[] encData_byte = new byte[password.Length];
                encData_byte = System.Text.Encoding.UTF8.GetBytes(password);
                string encodedData = Convert.ToBase64String(encData_byte);
                return encodedData;
            }
            catch (Exception ex)
            {
                throw new Exception("Error in base64Encode" + ex.Message);
            }
        }
        public string DecodeFrom64(string encodedData)
        {
            System.Text.UTF8Encoding encoder = new System.Text.UTF8Encoding();
            System.Text.Decoder utf8Decode = encoder.GetDecoder();
            byte[] todecode_byte = Convert.FromBase64String(encodedData);
            int charCount = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length);
            char[] decoded_char = new char[charCount];
            utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0);
            string result = new String(decoded_char);
            return result;
        }
    }
}
