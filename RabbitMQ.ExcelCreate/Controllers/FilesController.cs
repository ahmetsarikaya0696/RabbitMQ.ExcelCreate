using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.EntityFrameworkCore;
using RabbitMQ.ExcelCreate.Hubs;
using RabbitMQ.ExcelCreate.Models;

namespace RabbitMQ.ExcelCreate.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FilesController : ControllerBase
    {
        private readonly AppDbContext _appDbContext;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IHubContext<MyHub> _hubContext;

        public FilesController(AppDbContext appDbContext, IWebHostEnvironment webHostEnvironment, IHubContext<MyHub> hubContext)
        {
            _appDbContext = appDbContext;
            _webHostEnvironment = webHostEnvironment;
            _hubContext = hubContext;
        }

        public async Task<IActionResult> Upload(IFormFile file, int fileId)
        {
            if (file is not { Length: > 0 }) return BadRequest();

            var userFile = await _appDbContext.UserFiles.FirstOrDefaultAsync(x => x.Id == fileId);

            if (userFile is null) return NotFound();

            var filePath = userFile.FileName + Path.GetExtension(file.FileName);

            var path = Path.Combine(_webHostEnvironment.WebRootPath, "files", filePath);

            using FileStream stream = new(path, FileMode.Create);

            await file.CopyToAsync(stream);

            userFile.CreatedDate = DateTime.Now;
            userFile.FilePath = filePath;
            userFile.FileStatus = FileStatus.Completed;

            await _appDbContext.SaveChangesAsync();

            // SignalR notification
            await _hubContext.Clients.User(userFile.UserId).SendAsync("CompletedFile");

            return Ok();
        }
    }
}
