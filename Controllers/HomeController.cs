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
        // ������������� �������� �� ���������
        ��������� ��������� = new ���������();

        // �������� �������� � �������������
        ViewData["���������"] = ���������;

        return View(���������);
    }
    [HttpPost]
    public async Task<IActionResult> Index(IFormFile file, ��������� ���������)
    {
        try
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("���� �� ��� ��������.");
            }

            // �������� ���� �����
            if (!file.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("�������� ��� �����. �������� ������ Excel (.xlsx).");
            }

            string uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "Files");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }
            string fileName = Path.GetFileName(file.FileName);
            string fileSavePath = Path.Combine(uploadsFolder, "��������������_����.xlsx");

            using (FileStream stream = new FileStream(fileSavePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }


            // ��������� �����
            Excel excel = new Excel();
            DataTable dataTable = excel.ReadExcelFile(@fileSavePath);



            ������_������� ������_������� = new ������_�������();
            DataTable dtResult = ������_�������.�������_������(dataTable, ���������);
            string processedFilePath = Path.Combine(_webHostEnvironment.WebRootPath, "processed");
            if (!Directory.Exists(processedFilePath))
            {
                Directory.CreateDirectory(processedFilePath);
            }
            // ���������� ������������ ������ � ����� ����
            processedFilePath = Path.Combine(processedFilePath, "������������_����.xlsx");
            excel.SaveExel(dtResult, processedFilePath);

            // ����������� ����� ������������
            return File(System.IO.File.ReadAllBytes(processedFilePath), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "������������_����.xlsx");



        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message, "������ ��� ��������� ����� Excel");
            return BadRequest("��������� ������ ��� ��������� �����. ���������� �����.");
        }}
    }
