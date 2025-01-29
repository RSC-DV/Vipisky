using System.Data;
using System.Diagnostics;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Sebtum.Models;
using YourProjectName.Models;

namespace YourProjectName.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }
    private readonly IWebHostEnvironment _webHostEnvironment;
    public IActionResult Index()
    {
        // Инициализация настроек по умолчанию
        Настройки настройки = new Настройки();

        // Передача настроек в представление
        ViewData["Настройки"] = настройки;

        return View(настройки);
    }
    [HttpPost]
    public async Task<IActionResult> Index(IFormFile file, Настройки настройки)
    {
        try
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("Файл не был загружен.");
            }

            // Проверка типа файла
            if (!file.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("Неверный тип файла. Допустим только Excel (.xlsx).");
            }

            string uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "Files");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }
            string fileName = Path.GetFileName(file.FileName);
            string fileSavePath = Path.Combine(uploadsFolder, "Обработываемый_файл.xlsx");

            using (FileStream stream = new FileStream(fileSavePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }


            // Обработка файла
            Excel excel = new Excel();
            DataTable dataTable = excel.ReadExcelFile(@fileSavePath);



            Анализ_выписки анализ_Выписки = new Анализ_выписки();
            DataTable dtResult = анализ_Выписки.Сделать_анализ(dataTable, настройки);
            string processedFilePath = Path.Combine(_webHostEnvironment.WebRootPath, "processed");
            if (!Directory.Exists(processedFilePath))
            {
                Directory.CreateDirectory(processedFilePath);
            }
            // Сохранение обработанных данных в новый файл
            processedFilePath = Path.Combine(processedFilePath, "Обработанный_файл.xlsx");
            excel.SaveExel(dtResult, processedFilePath);

            // Возвращение файла пользователю
            return File(System.IO.File.ReadAllBytes(processedFilePath), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Обработанный_файл.xlsx");



        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message, "Ошибка при обработке файла Excel");
            return BadRequest("Произошла ошибка при обработке файла. Попробуйте снова.");
        }}
    }
