using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;

namespace RabbitMQ.ExcelCreate.Controllers
{
    public class AccountController : Controller
    {
        private readonly UserManager<IdentityUser> _userManager;
        private readonly SignInManager<IdentityUser> _signInManager;

        public AccountController(SignInManager<IdentityUser> signInManager, UserManager<IdentityUser> userManager)
        {
            _signInManager = signInManager;
            _userManager = userManager;
        }

        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Login(string Email, string Password)
        {
            var user = await _userManager.FindByEmailAsync(Email);

            if (user == null)
            {
                return View();
            }

            var signInResult = await _signInManager.PasswordSignInAsync(user, Password, true, false);

            if (!signInResult.Succeeded)
            {
                return View();
            }

            return RedirectToAction(nameof(HomeController.Index), "Home");
        }
    }
}
