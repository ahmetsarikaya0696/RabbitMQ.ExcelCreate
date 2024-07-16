using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RabbitMQ.ExcelCreate.Models;
using RabbitMQ.ExcelCreate.Services;
using Shared;

namespace RabbitMQ.ExcelCreate.Controllers
{
    [Authorize]
    public class ProductsController : Controller
    {
        private readonly AppDbContext _appDbContext;
        private readonly UserManager<IdentityUser> _userManager;
        private readonly RabbitMQPublisher _rabbitMQPublisher;

        public ProductsController(AppDbContext appDbContext, UserManager<IdentityUser> userManager, RabbitMQPublisher rabbitMQPublisher)
        {
            _appDbContext = appDbContext;
            _userManager = userManager;
            _rabbitMQPublisher = rabbitMQPublisher;
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<IActionResult> CreateProductExcel()
        {
            var user = await _userManager.FindByNameAsync(User.Identity.Name);

            var fileName = $"product-excel-{Guid.NewGuid()}";

            UserFile userFile = new()
            {
                UserId = user.Id,
                FileName = fileName,
                FileStatus = FileStatus.Creating,
            };

            await _appDbContext.UserFiles.AddAsync(userFile);
            await _appDbContext.SaveChangesAsync();

            _rabbitMQPublisher.Publish(new CreateExcelMessage()
            {
                FileId = userFile.Id
            });

            TempData["StartCreatingExcel"] = true;
            return RedirectToAction(nameof(Files));
        }

        public async Task<IActionResult> Files()
        {
            var user = await _userManager.FindByNameAsync(User.Identity.Name);

            var userFiles = await _appDbContext.UserFiles.Where(x => x.UserId == user.Id).ToListAsync();

            return View(userFiles);
        }
    }
}
