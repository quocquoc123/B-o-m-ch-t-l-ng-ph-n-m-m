using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;
using ExcelDataReader;
using System.Data;
using System.Text;
namespace Testdangnhapadmin
{
    public class Tests
    {
        private ChromeDriver driver;
        private string baseUrl = "https://localhost:7053/";
        private List<Dictionary<string, string>> loginData;
        private List<Dictionary<string, string>> productData;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(baseUrl);

            // Đọc dữ liệu từ file Excel
            string filePath = @"D:\TestData.xlsx";
            loginData = ReadExcel(filePath, "Login");
            productData = ReadExcel(filePath, "Products");
        }
        [Test]
        public void TestDangNhapAdmin()
        {
            var data = loginData[0]; // Chọn dòng thứ 2 (index 1)

            driver.Navigate().GoToUrl(baseUrl);
            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("bx-user")).Click(); // Click icon đăng nhập
            Thread.Sleep(2000);

            driver.FindElement(By.Id("Username")).SendKeys(data["Username"]);
            driver.FindElement(By.Id("Matkhau")).SendKeys(data["Password"] + Keys.Enter);
            Thread.Sleep(3000);
        }



        [Test]
        public void TestTaoSanPham()
        {
            foreach (var product in productData)
            {
                DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
                string formattedDate = hanSuDung.ToString("dd/MM/yyyy");
                driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");
                driver.FindElement(By.Id("TenSp")).SendKeys(product["TenSp"]);
                driver.FindElement(By.Id("ThanhPhan")).SendKeys(product["ThanhPhan"]);
                driver.FindElement(By.Id("GiaTien")).SendKeys(product["GiaTien"]);
                driver.FindElement(By.Id("DonVi")).SendKeys(product["DonVi"]);
                driver.FindElement(By.Id("HansuDung")).SendKeys(formattedDate);
                driver.FindElement(By.Id("ChitietSp")).SendKeys(product["ChiTietSp"]);
                driver.FindElement(By.Id("SoLuong")).SendKeys(product["SoLuong"]);
                driver.FindElement(By.Id("SoLuongMua")).SendKeys(product["SoLuongMua"]);
                driver.FindElement(By.Id("SoBinhLuan")).SendKeys(product["SoBinhLuan"]);
                driver.FindElement(By.Id("Congdung")).SendKeys(product["CongDung"]);
                driver.FindElement(By.Id("Cachdung")).SendKeys(product["CachDung"]);
                driver.FindElement(By.Id("Doituongsudung")).SendKeys(product["DoiTuongSuDung"]);
                driver.FindElement(By.Id("Tacdungphu")).SendKeys(product["TacDungPhu"]);
                driver.FindElement(By.Id("Ngaysanxuat")).SendKeys(product["NgaySanXuat"]);
                driver.FindElement(By.Id("Noisanxuat")).SendKeys(product["NoiSanXuat"]);
                Thread.Sleep(2000);

                // Chọn danh mục

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("window.scrollBy(0,6000);");
                // Upload ảnh
                // Lấy danh sách ảnh từ nhiều cột trong Excel
                string[] imagePaths = new string[]
                {
    product.ContainsKey("ImagePaths1") ? product["ImagePaths1"] : "",
    product.ContainsKey("ImagePaths2") ? product["ImagePaths2"] : "",
    product.ContainsKey("ImagePaths3") ? product["ImagePaths3"] : "",
    product.ContainsKey("ImagePaths4") ? product["ImagePaths4"] : ""
                };

                string[] inputIds = { "file1", "file2", "file3", "file4" };

                var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
                .Where(x => !string.IsNullOrEmpty(x.path))
                .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
                .Where(x => x.element != null);

                validImages.ToList().ForEach(image =>
                {
                    image.element.SendKeys(image.path);
                    Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
                });
                Thread.Sleep(2000);


                // Bấm nút tạo sản phẩm
                var createButton = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
                createButton?.Click();
                Thread.Sleep(2000);
                Console.WriteLine(createButton != null ? "✅ Nút Create đã được nhấn" : "❌ Không tìm thấy nút Create!");

                Thread.Sleep(2000);

                // Kiểm tra sản phẩm đã được tạo
                driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams");

                var table = driver.FindElements(By.ClassName("table")).FirstOrDefault();
                var tableText = table?.Text ?? string.Empty;
                var missingFields = new List<string>();

                new[] { "TenSp", "ThanhPhan", "GiaTien", "DonVi", "ChiTietSp", "SoLuong", "SoLuongMua" }
                    .Where(field => !tableText.Contains(product[field]))
                    .ToList()
                    .ForEach(missingFields.Add);

                Console.WriteLine(missingFields.Any()
                    ? $"Test FAILED: Thiếu {string.Join(", ", missingFields)}"
                    : "Tạo sản phẩm thành công");

                (missingFields.Any() ? new Action(() => Assert.Fail($"Thiếu {string.Join(", ", missingFields)}")) : Assert.Pass)();
                Thread.Sleep(2000);

            }
        }

        [Test]
        public void TestTaoSanPham_DeTrongThongTin()
        {
            using IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            string[] fieldIds = {
        "TenSp", "ThanhPhan", "GiaTien", "DonVi", "HansuDung", "ChitietSp",
        "SoLuong", "SoLuongMua", "SoBinhLuan", "Congdung", "Cachdung", "Doituongsudung",
        "Tacdungphu", "Ngaysanxuat", "Noisanxuat"
    };

            fieldIds.ToList().ForEach(fieldId =>
                driver.FindElements(By.Id(fieldId)).FirstOrDefault()?.Clear()
            );

            Thread.Sleep(1000);

            driver.FindElements(By.Id("MaDm"))
                .Select(d => new SelectElement(d))
                .Where(select => select.Options.Count > 1)
                .ToList()
                .ForEach(select => select.SelectByIndex(0));

            Thread.Sleep(1000);

            var createButton = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
            ((IJavaScriptExecutor)driver)?.ExecuteScript("arguments[0].scrollIntoView(true);", createButton);
            Thread.Sleep(500);

            createButton?.Click();

            Thread.Sleep(2000);

            var errorMessages = driver.FindElements(By.ClassName("text-danger"));
            errorMessages.ToList().ForEach(error => Console.WriteLine($"✅ Test Passed! Báo lỗi: {error.Text}"));

            Assert.That(errorMessages.Count, Is.GreaterThan(0), "❌ Test Failed! Không có thông báo lỗi.");
        }

        [Test]
        public void TestTaoSanPham_TrungTen()
        {
            string filePath = @"D:\TestData.xlsx";

            // Đọc dữ liệu từ SHEET 3
            var productData = ReadExcel(filePath, "Products");



            // ✅ Lấy dữ liệu từ dòng đầu tiên (index 0)
            var product = productData[0];

            DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
            string formattedDate = hanSuDung.ToString("dd/MM/yyyy");

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");

            // 🔹 Nhập thông tin sản phẩm
            driver.FindElement(By.Id("TenSp")).SendKeys(product["TenSp"]);
            driver.FindElement(By.Id("ThanhPhan")).SendKeys(product["ThanhPhan"]);
            driver.FindElement(By.Id("GiaTien")).SendKeys(product["GiaTien"]);
            driver.FindElement(By.Id("DonVi")).SendKeys(product["DonVi"]);
            driver.FindElement(By.Id("HansuDung")).SendKeys(formattedDate);
            driver.FindElement(By.Id("ChitietSp")).SendKeys(product["ChiTietSp"]);
            driver.FindElement(By.Id("SoLuong")).SendKeys(product["SoLuong"]);
            driver.FindElement(By.Id("SoLuongMua")).SendKeys(product["SoLuongMua"]);
            driver.FindElement(By.Id("SoBinhLuan")).SendKeys(product["SoBinhLuan"]);
            driver.FindElement(By.Id("Congdung")).SendKeys(product["CongDung"]);
            driver.FindElement(By.Id("Cachdung")).SendKeys(product["CachDung"]);
            driver.FindElement(By.Id("Doituongsudung")).SendKeys(product["DoiTuongSuDung"]);
            driver.FindElement(By.Id("Tacdungphu")).SendKeys(product["TacDungPhu"]);
            driver.FindElement(By.Id("Ngaysanxuat")).SendKeys(product["NgaySanXuat"]);
            driver.FindElement(By.Id("Noisanxuat")).SendKeys(product["NoiSanXuat"]);
            Thread.Sleep(2000);

            // 🔹 Chọn danh mục
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,6000);");
            string[] imagePaths =
         {
        product.GetValueOrDefault("ImagePaths1", ""),
        product.GetValueOrDefault("ImagePaths2", ""),
        product.GetValueOrDefault("ImagePaths3", ""),
        product.GetValueOrDefault("ImagePaths4", "")
    };

            string[] inputIds = { "file1", "file2", "file3", "file4" };

            var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
                 .Where(x => !string.IsNullOrEmpty(x.path))
                 .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
                 .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);
            // 🔹 Bấm nút tạo sản phẩm
            // Bấm nút tạo sản phẩm
            var createButton = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
            createButton?.Click();
            Thread.Sleep(2000);
            Console.WriteLine(createButton != null ? "✅ Nút Create đã được nhấn" : "❌ Không tìm thấy nút Create!");

            Thread.Sleep(2000);

            // 🔹 Nhập lại cùng tên sản phẩm và thử tạo lần 2
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");

            driver.FindElement(By.Id("TenSp")).SendKeys(product["TenSp"]); // 🔴 Tên trùng
            driver.FindElement(By.Id("ThanhPhan")).SendKeys(product["ThanhPhan"]);
            driver.FindElement(By.Id("GiaTien")).SendKeys(product["GiaTien"]);
            driver.FindElement(By.Id("DonVi")).SendKeys(product["DonVi"]);
            driver.FindElement(By.Id("HansuDung")).SendKeys(formattedDate);
            driver.FindElement(By.Id("ChitietSp")).SendKeys(product["ChiTietSp"]);
            driver.FindElement(By.Id("SoLuong")).SendKeys(product["SoLuong"]);
            driver.FindElement(By.Id("SoLuongMua")).SendKeys(product["SoLuongMua"]);
            driver.FindElement(By.Id("SoBinhLuan")).SendKeys(product["SoBinhLuan"]);
            driver.FindElement(By.Id("Congdung")).SendKeys(product["CongDung"]);
            driver.FindElement(By.Id("Cachdung")).SendKeys(product["CachDung"]);
            driver.FindElement(By.Id("Doituongsudung")).SendKeys(product["DoiTuongSuDung"]);
            driver.FindElement(By.Id("Tacdungphu")).SendKeys(product["TacDungPhu"]);
            driver.FindElement(By.Id("Ngaysanxuat")).SendKeys(product["NgaySanXuat"]);
            driver.FindElement(By.Id("Noisanxuat")).SendKeys(product["NoiSanXuat"]);

            Thread.Sleep(2000);

            js.ExecuteScript("window.scrollBy(0,3000);");
            string[] imagePaths1 =
         {
        product.GetValueOrDefault("ImagePaths1", ""),
        product.GetValueOrDefault("ImagePaths2", ""),
        product.GetValueOrDefault("ImagePaths3", ""),
        product.GetValueOrDefault("ImagePaths4", "")
    };

            string[] inputIds1 = { "file1", "file2", "file3", "file4" };

            var validImages1 = inputIds.Zip(imagePaths, (id, path) => new { id, path })
                .Where(x => !string.IsNullOrEmpty(x.path))
                .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
                .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);
            // Bấm nút tạo sản phẩm
            var createButton1 = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
            createButton1?.Click();
            Thread.Sleep(2000);
            Console.WriteLine(createButton != null ? "✅ Nút Create đã được nhấn" : "❌ Không tìm thấy nút Create!");

            Thread.Sleep(2000);

            // 🔹 Kiểm tra thông báo lỗi "Tên sản phẩm đã tồn tại"
            var errorMessages = driver.FindElements(By.ClassName("text-danger"));
            errorMessages.ToList().ForEach(error => Console.WriteLine($"✅ Test Passed! Báo lỗi: {error.Text}"));

            Assert.That(errorMessages.Count, Is.GreaterThan(0), "❌ Test Failed! Không có thông báo sản phẩm tạo đã bị trùng tên.");
        }


        [Test]
        public void TestTaoSanPham_GiaAm()
        {

            var product = productData[1]; // Chỉ lấy dòng thứ 2

            DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
            string formattedDate = hanSuDung.ToString("dd/MM/yyyy");

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");

            // Nhập thông tin sản phẩm
            driver.FindElement(By.Id("TenSp")).SendKeys(product["TenSp"]);
            driver.FindElement(By.Id("ThanhPhan")).SendKeys(product["ThanhPhan"]);
            driver.FindElement(By.Id("GiaTien")).SendKeys(product["GiaTien"]);
            driver.FindElement(By.Id("DonVi")).SendKeys(product["DonVi"]);
            driver.FindElement(By.Id("HansuDung")).SendKeys(formattedDate);
            driver.FindElement(By.Id("ChitietSp")).SendKeys(product["ChiTietSp"]);
            driver.FindElement(By.Id("SoLuong")).SendKeys(product["SoLuong"]);
            driver.FindElement(By.Id("SoLuongMua")).SendKeys(product["SoLuongMua"]);
            driver.FindElement(By.Id("SoBinhLuan")).SendKeys(product["SoBinhLuan"]);
            driver.FindElement(By.Id("Congdung")).SendKeys(product["CongDung"]);
            driver.FindElement(By.Id("Cachdung")).SendKeys(product["CachDung"]);
            driver.FindElement(By.Id("Doituongsudung")).SendKeys(product["DoiTuongSuDung"]);
            driver.FindElement(By.Id("Tacdungphu")).SendKeys(product["TacDungPhu"]);
            driver.FindElement(By.Id("Ngaysanxuat")).SendKeys(product["NgaySanXuat"]);
            driver.FindElement(By.Id("Noisanxuat")).SendKeys(product["NoiSanXuat"]);
            Thread.Sleep(2000);
            // Chọn danh mục
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,6000);");
            // Upload ảnh
            string[] imagePaths = new string[]
           {
    product.ContainsKey("ImagePaths1") ? product["ImagePaths1"] : "",
    product.ContainsKey("ImagePaths2") ? product["ImagePaths2"] : "",
    product.ContainsKey("ImagePaths3") ? product["ImagePaths3"] : "",
    product.ContainsKey("ImagePaths4") ? product["ImagePaths4"] : ""
           };

            string[] inputIds = { "file1", "file2", "file3", "file4" };

            var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
            .Where(x => !string.IsNullOrEmpty(x.path))
            .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
            .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);


            // Bấm nút tạo sản phẩm
            var createButton = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
            createButton?.Click();
            Thread.Sleep(2000);
            Console.WriteLine(createButton != null ? "✅ Nút Create đã được nhấn" : "❌ Không tìm thấy nút Create!");

            Thread.Sleep(2000);


            var errorMessages = driver.FindElements(By.ClassName("text-danger"));
            errorMessages.ToList().ForEach(error => Console.WriteLine($"✅ Test Passed! Báo lỗi: {error.Text}"));

            Assert.That(errorMessages.Count, Is.GreaterThan(0), "❌ Test Failed! Không có thông báo lỗi giá tiền không được âm.");
        }



        [Test]
        public void TestTaoSanPham_ChiNhapTen()
        {
            string filePath = @"D:\TestData.xlsx";
            string sheetName = "Products";        // Tên sheet
            List<Dictionary<string, string>> productData = ReadExcel(filePath, sheetName);



            // Chỉ lấy 1 sản phẩm đầu tiên để test
            var product = productData[0];

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            var inputTenSp = wait.Until(d => d.FindElements(By.Id("TenSp")).FirstOrDefault());
            inputTenSp?.Clear();
            inputTenSp?.SendKeys(product["TenSp"]);
            Console.WriteLine(inputTenSp != null ? "✅ Nhập tên sản phẩm thành công" : "❌ Không tìm thấy input tên sản phẩm!");
            (inputTenSp == null ? new Action(() => { driver.Quit();
                Assert.Fail("Không tìm thấy input tên sản phẩm!"); }) : () => { })();

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,3000);");

            Thread.Sleep(2000);

            var createButton = wait.Until(d => d.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault());

            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", createButton);
            Thread.Sleep(500);

            (createButton?.Displayed == true && createButton.Enabled
                ? new Action(() => createButton.Click())
                : new Action(() => ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", createButton)))();
            Thread.Sleep(2000);

            var errorMessages = driver.FindElements(By.ClassName("text-danger"));
            Assert.That(errorMessages.Count, Is.GreaterThan(0), "Không có thông báo lỗi.");
            Console.WriteLine("✅ Test Passed! Hệ thống báo lỗi.");

            driver.Quit();
        }


        //Test HanSDTronQK
        [Test]
        public void TestHanSDTronQK()
        {

            string filePath = @"D:\TestData.xlsx";
            string sheetName = "Products";


            var product = productData[6];

            DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
            string formattedDate = hanSuDung.ToString("dd/MM/yyyy");

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");

            // Nhập thông tin sản phẩm
            driver.FindElement(By.Id("TenSp")).SendKeys(product["TenSp"]);
            driver.FindElement(By.Id("ThanhPhan")).SendKeys(product["ThanhPhan"]);
            driver.FindElement(By.Id("GiaTien")).SendKeys(product["GiaTien"]); // Giá không phải số
            driver.FindElement(By.Id("DonVi")).SendKeys(product["DonVi"]);
            driver.FindElement(By.Id("HansuDung")).SendKeys(formattedDate);
            driver.FindElement(By.Id("ChitietSp")).SendKeys(product["ChiTietSp"]);
            driver.FindElement(By.Id("SoLuong")).SendKeys(product["SoLuong"]);
            driver.FindElement(By.Id("SoLuongMua")).SendKeys(product["SoLuongMua"]);
            driver.FindElement(By.Id("SoBinhLuan")).SendKeys(product["SoBinhLuan"]);
            driver.FindElement(By.Id("Congdung")).SendKeys(product["CongDung"]);
            driver.FindElement(By.Id("Cachdung")).SendKeys(product["CachDung"]);
            driver.FindElement(By.Id("Doituongsudung")).SendKeys(product["DoiTuongSuDung"]);
            driver.FindElement(By.Id("Tacdungphu")).SendKeys(product["TacDungPhu"]);
            driver.FindElement(By.Id("Ngaysanxuat")).SendKeys(product["NgaySanXuat"]);
            driver.FindElement(By.Id("Noisanxuat")).SendKeys(product["NoiSanXuat"]);
            Thread.Sleep(2000);

            // Chọn danh mục
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,6000);");
            string[] imagePaths =
             {
        product.GetValueOrDefault("ImagePaths1", ""),
        product.GetValueOrDefault("ImagePaths2", ""),
        product.GetValueOrDefault("ImagePaths3", ""),
        product.GetValueOrDefault("ImagePaths4", "")
    };

            string[] inputIds = { "file1", "file2", "file3", "file4" };

            var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
                .Where(x => !string.IsNullOrEmpty(x.path))
                .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
                .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);
            // Bấm nút tạo sản phẩm
            var createButton = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
            createButton?.Click();
            Thread.Sleep(2000);
            Console.WriteLine(createButton != null ? "✅ Nút Create đã được nhấn" : "❌ Không tìm thấy nút Create!");


            Thread.Sleep(2000);

            var errorMessages = driver.FindElements(By.ClassName("text-danger"));
            errorMessages.ToList().ForEach(error => Console.WriteLine($"✅ Test Passed! Báo lỗi: {error.Text}"));

            Assert.That(errorMessages.Count, Is.GreaterThan(0), "❌ Test Failed! Không có thông báo lỗi sai hạn sử dụng đã trong quá khứ.");
        }

        [Test]
        public void TestTaoSanPham_GiaTienKhongPhaiSo()
        {
            string filePath = @"D:\TestData.xlsx";
            string sheetName = "Products";


            var product = productData[5];

            DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
            string formattedDate = hanSuDung.ToString("dd/MM/yyyy");

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");

            // Nhập thông tin sản phẩm
            driver.FindElement(By.Id("TenSp")).SendKeys(product["TenSp"]);
            driver.FindElement(By.Id("ThanhPhan")).SendKeys(product["ThanhPhan"]);
            driver.FindElement(By.Id("GiaTien")).SendKeys(product["GiaTien"]); // Giá không phải số
            driver.FindElement(By.Id("DonVi")).SendKeys(product["DonVi"]);
            driver.FindElement(By.Id("HansuDung")).SendKeys(formattedDate);
            driver.FindElement(By.Id("ChitietSp")).SendKeys(product["ChiTietSp"]);
            driver.FindElement(By.Id("SoLuong")).SendKeys(product["SoLuong"]);
            driver.FindElement(By.Id("SoLuongMua")).SendKeys(product["SoLuongMua"]);
            driver.FindElement(By.Id("SoBinhLuan")).SendKeys(product["SoBinhLuan"]);
            driver.FindElement(By.Id("Congdung")).SendKeys(product["CongDung"]);
            driver.FindElement(By.Id("Cachdung")).SendKeys(product["CachDung"]);
            driver.FindElement(By.Id("Doituongsudung")).SendKeys(product["DoiTuongSuDung"]);
            driver.FindElement(By.Id("Tacdungphu")).SendKeys(product["TacDungPhu"]);
            driver.FindElement(By.Id("Ngaysanxuat")).SendKeys(product["NgaySanXuat"]);
            driver.FindElement(By.Id("Noisanxuat")).SendKeys(product["NoiSanXuat"]);
            Thread.Sleep(2000);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,6000);");
            // Chọn danh mục

            string[] imagePaths =
         {
        product.GetValueOrDefault("ImagePaths1", ""),
        product.GetValueOrDefault("ImagePaths2", ""),
        product.GetValueOrDefault("ImagePaths3", ""),
        product.GetValueOrDefault("ImagePaths4", "")
    };

            string[] inputIds = { "file1", "file2", "file3", "file4" };

            var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
                .Where(x => !string.IsNullOrEmpty(x.path))
                .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
                .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);
            // Bấm nút tạo sản phẩm
            var createButton = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
            createButton?.Click();
            Thread.Sleep(2000);
            Console.WriteLine(createButton != null ? "✅ Nút Create đã được nhấn" : "❌ Không tìm thấy nút Create!");


            Thread.Sleep(2000);

            var errorMessages = driver.FindElements(By.ClassName("text-danger"));
            errorMessages.ToList().ForEach(error => Console.WriteLine($"✅ Test Passed! Báo lỗi: {error.Text}"));

            Assert.That(errorMessages.Count, Is.GreaterThan(0), "❌ Test Failed! Không có thông báo lỗi sai định dạng tiền.");
        }


        [Test]
        public void TestTaoSanPhamVoiSoLuong0()
        {
            var product = productData[4];

            DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
            string formattedDate = hanSuDung.ToString("dd/MM/yyyy");
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Create");
            driver.FindElement(By.Id("TenSp")).SendKeys(product["TenSp"]);
            driver.FindElement(By.Id("ThanhPhan")).SendKeys(product["ThanhPhan"]);
            driver.FindElement(By.Id("GiaTien")).SendKeys(product["GiaTien"]);
            driver.FindElement(By.Id("DonVi")).SendKeys(product["DonVi"]);
            driver.FindElement(By.Id("HansuDung")).SendKeys(formattedDate);
            driver.FindElement(By.Id("ChitietSp")).SendKeys(product["ChiTietSp"]);
            driver.FindElement(By.Id("SoLuong")).SendKeys(product["SoLuong"]);
            driver.FindElement(By.Id("SoLuongMua")).SendKeys(product["SoLuongMua"]);
            driver.FindElement(By.Id("SoBinhLuan")).SendKeys(product["SoBinhLuan"]);
            driver.FindElement(By.Id("Congdung")).SendKeys(product["CongDung"]);
            driver.FindElement(By.Id("Cachdung")).SendKeys(product["CachDung"]);
            driver.FindElement(By.Id("Doituongsudung")).SendKeys(product["DoiTuongSuDung"]);
            driver.FindElement(By.Id("Tacdungphu")).SendKeys(product["TacDungPhu"]);
            driver.FindElement(By.Id("Ngaysanxuat")).SendKeys(product["NgaySanXuat"]);
            driver.FindElement(By.Id("Noisanxuat")).SendKeys(product["NoiSanXuat"]);
            Thread.Sleep(2000);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,6000);");
            // Chọn danh mục


            // Upload ảnh
            // Lấy danh sách ảnh từ nhiều cột trong Excel
            string[] imagePaths = new string[]
                {
    product.ContainsKey("ImagePaths1") ? product["ImagePaths1"] : "",
    product.ContainsKey("ImagePaths2") ? product["ImagePaths2"] : "",
    product.ContainsKey("ImagePaths3") ? product["ImagePaths3"] : "",
    product.ContainsKey("ImagePaths4") ? product["ImagePaths4"] : ""
                };

            string[] inputIds = { "file1", "file2", "file3", "file4" };

            var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
            .Where(x => !string.IsNullOrEmpty(x.path))
            .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
            .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);


            // Bấm nút tạo sản phẩm
            var createButton = driver.FindElements(By.XPath("//input[@type='submit' and @value='Create']")).FirstOrDefault();
            createButton?.Click();
            Thread.Sleep(2000);
            Console.WriteLine(createButton != null ? "✅ Nút Create đã được nhấn" : "❌ Không tìm thấy nút Create!");

            Thread.Sleep(2000);

            Thread.Sleep(3000);

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams");

            var errorMessages = driver.FindElements(By.ClassName("text-danger"));

            Assert.IsTrue(errorMessages.Any(), "❌ Test Failed! Không có thông báo lỗi khi nhập số lượng bằng 0.");

            Console.WriteLine("✅ Test Passed! Hệ thống báo lỗi khi nhập số lượng bằng 0.");
            errorMessages.ToList().ForEach(error => Console.WriteLine($"🔹 Lỗi: {error.Text}"));

            driver.Quit();
        }
        //Chỉnh sửa
        [Test]
        public void TestChinhSuaTenSanPham()
        {
            string filePath = @"D:\TestData.xlsx";
            string sheetName = "Products";
            var productData = ReadExcel(filePath, sheetName);
            Thread.Sleep(2000);


            var product = productData[3];

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(2000);

            IWebElement editButton = driver.FindElement(By.XPath("//a[contains(text(), 'Sửa')]"));
            Actions actions = new Actions(driver);
            actions.MoveToElement(editButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn vào 'Sửa'");

            Thread.Sleep(2000);
            Assert.That(driver.Url, Does.Contain("Edit"), "❌ Không vào được trang chỉnh sửa.");

            // ✅ Chỉnh sửa tên sản phẩm
            IWebElement inputField = driver.FindElement(By.Id("TenSp"));
            inputField.Clear();
            inputField.SendKeys(product["TenSp"]);
            Console.WriteLine($"✏️ Đã sửa tên sản phẩm thành: {product["TenSp"]}");

            Thread.Sleep(2000);

            // ✅ Nhấn lưu sản phẩm
            IWebElement saveButton = driver.FindElement(By.ClassName("btn-primary"));
            actions.MoveToElement(saveButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn 'Lưu'.");

            Thread.Sleep(2000);
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(2000);

            // ✅ Kiểm tra xem tên sản phẩm đã cập nhật chưa
            IWebElement table = driver.FindElement(By.ClassName("table"));
            Assert.That(table.Text, Does.Contain(product["TenSp"]), $"❌ Test Failed! Không tìm thấy sản phẩm với tên mới: {product["TenSp"]}");

            Console.WriteLine($"✅ Test Passed! Sản phẩm đã được cập nhật thành: {product["TenSp"]}");
            Console.WriteLine("Test hoàn thành.");
            driver.Quit();
        }

        [Test]
        public void TestChinhSuaSanPham()
        {
            string filePath = @"D:\\TestData.xlsx";
            string sheetName = "Products";
            var productData = ReadExcel(filePath, sheetName);



            var product = productData[2];
            DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
            string formattedDate = hanSuDung.ToString("dd/MM/yyyy");
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(2000);

            IWebElement editButton = driver.FindElement(By.XPath("//a[contains(text(), 'Sửa')]"));
            Actions actions = new Actions(driver);
            actions.MoveToElement(editButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn vào 'Sửa'");

            Thread.Sleep(2000);
            Assert.That(driver.Url, Does.Contain("Edit"), "❌ Không vào được trang chỉnh sửa.");


            var fields = new (string Id, string Key)[] {
        ("TenSp", "TenSp"),
        ("ThanhPhan", "ThanhPhan"),
        ("GiaTien", "GiaTien"),
        ("DonVi", "DonVi"),
        ("ChitietSp", "ChiTietSp"),
        ("SoLuong", "SoLuong"),
        ("SoLuongMua", "SoLuongMua"),
        ("SoBinhLuan", "SoBinhLuan"),
        ("Congdung", "CongDung"),
        ("Cachdung", "CachDung"),
        ("Doituongsudung", "DoiTuongSuDung"),
        ("Tacdungphu", "TacDungPhu"),
        ("Ngaysanxuat", "NgaySanXuat"),
        ("Noisanxuat", "NoiSanXuat")
    };

            foreach (var field in fields.Where(field => product.ContainsKey(field.Key)))
            {
                IWebElement inputField = driver.FindElement(By.Id(field.Id));
                inputField.Clear();
                inputField.SendKeys(product[field.Key]);
                Console.WriteLine($"✏️ Đã sửa {field.Key}.");
            }

            IWebElement hansuDungField = driver.FindElement(By.Id("HansuDung"));
            hansuDungField.Clear();
            hansuDungField.SendKeys(formattedDate);
            Console.WriteLine("✏️ Đã sửa HansuDung với định dạng ngày chuẩn.");

            string[] imagePaths = new string[]
        {
    product.ContainsKey("ImagePaths1") ? product["ImagePaths1"] : "",
    product.ContainsKey("ImagePaths2") ? product["ImagePaths2"] : "",
    product.ContainsKey("ImagePaths3") ? product["ImagePaths3"] : "",
    product.ContainsKey("ImagePaths4") ? product["ImagePaths4"] : ""
        };

            string[] inputIds = { "file1", "file2", "file3", "file4" };

            var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
                .Where(x => !string.IsNullOrEmpty(x.path))
                .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
                .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);


            IWebElement saveButton = driver.FindElement(By.ClassName("btn-primary"));

            actions.MoveToElement(saveButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn 'Lưu'.");

            Thread.Sleep(2000);
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(2000);

            IWebElement table = driver.FindElement(By.ClassName("table"));
            Assert.That(table.Text, Does.Contain(product["TenSp"]), $"❌ Test Failed! Không tìm thấy sản phẩm với tên mới: {product["TenSp"]}");

            Console.WriteLine($"✅ Test Passed! Sản phẩm đã được cập nhật thành: {product["TenSp"]}");
            Console.WriteLine("Test hoàn thành.");
            driver.Quit();
        }
        [Test]
        public void TestChinhSuaSanPhamMaKoDoiTT()
        {
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(3000);

            // Lấy danh sách các dòng sản phẩm trong bảng
            var rows = driver.FindElements(By.XPath("//table[contains(@class, 'table')]//tr"));

            Assert.That(rows.Count, Is.GreaterThan(1), "❌ Không có sản phẩm nào để chỉnh sửa!");

            // Lấy hàng đầu tiên chứa sản phẩm (bỏ qua tiêu đề)
            var firstRow = rows[1];  // Index 1 vì index 0 là header

            // Lấy tên sản phẩm đầu tiên để kiểm tra sau này
            string firstProductName = firstRow.FindElement(By.XPath(".//td[1]")).Text.Trim();
            Console.WriteLine($"🔍 Đang chỉnh sửa sản phẩm đầu tiên: {firstProductName}");

            // Click vào nút "Sửa" trong hàng đầu tiên
            IWebElement editButton = firstRow.FindElement(By.XPath(".//a[contains(text(), 'Sửa')]"));
            new Actions(driver).MoveToElement(editButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn vào 'Sửa'.");

            Thread.Sleep(2000);
            Assert.That(driver.Url, Does.Contain("Edit"), "❌ Không vào được trang chỉnh sửa.");

            // Click "Lưu" mà không chỉnh sửa
            IWebElement saveButton = driver.FindElement(By.ClassName("btn-primary"));
            new Actions(driver).MoveToElement(saveButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn 'Lưu'.");

            Thread.Sleep(2000);
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(2000);

            // Kiểm tra lại xem sản phẩm đầu tiên vẫn còn
            IWebElement table = driver.FindElement(By.ClassName("table"));
            Assert.That(table.Text, Does.Contain(firstProductName), $"❌ Test Failed! Không tìm thấy sản phẩm: {firstProductName}");

            Console.WriteLine($"✅ Test Passed! Sản phẩm vẫn giữ nguyên tên: {firstProductName}");
            driver.Quit();
        }


        [Test]
        public void TestChinhSuaSanPhamTrungTen()
        {

            string filePath = @"D:\\TestData.xlsx";
            string sheetName = "Products";
            var productData = ReadExcel(filePath, sheetName);



            var product = productData[0];
            DateTime hanSuDung = DateTime.Parse(product["HanSuDung"]);
            string formattedDate = hanSuDung.ToString("dd/MM/yyyy");
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(2000);

            IWebElement editButton = driver.FindElement(By.XPath("//a[contains(text(), 'Sửa')]"));
            Actions actions = new Actions(driver);
            actions.MoveToElement(editButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn vào 'Sửa'");

            Thread.Sleep(2000);
            Assert.That(driver.Url, Does.Contain("Edit"), "❌ Không vào được trang chỉnh sửa.");


            var fields = new (string Id, string Key)[] {
        ("TenSp", "TenSp"),
        ("ThanhPhan", "ThanhPhan"),
        ("GiaTien", "GiaTien"),
        ("DonVi", "DonVi"),
        ("ChitietSp", "ChiTietSp"),
        ("SoLuong", "SoLuong"),
        ("SoLuongMua", "SoLuongMua"),
        ("SoBinhLuan", "SoBinhLuan"),
        ("Congdung", "CongDung"),
        ("Cachdung", "CachDung"),
        ("Doituongsudung", "DoiTuongSuDung"),
        ("Tacdungphu", "TacDungPhu"),
        ("Ngaysanxuat", "NgaySanXuat"),
        ("Noisanxuat", "NoiSanXuat")
    };

            foreach (var field in fields.Where(field => product.ContainsKey(field.Key)))
            {
                IWebElement inputField = driver.FindElement(By.Id(field.Id));
                inputField.Clear();
                inputField.SendKeys(product[field.Key]);
                Console.WriteLine($"✏️ Đã sửa {field.Key}.");
            }

            IWebElement hansuDungField = driver.FindElement(By.Id("HansuDung"));
            hansuDungField.Clear();
            hansuDungField.SendKeys(formattedDate);
            Console.WriteLine("✏️ Đã sửa HansuDung với định dạng ngày chuẩn.");

            string[] imagePaths = new string[]
        {
    product.ContainsKey("ImagePaths1") ? product["ImagePaths1"] : "",
    product.ContainsKey("ImagePaths2") ? product["ImagePaths2"] : "",
    product.ContainsKey("ImagePaths3") ? product["ImagePaths3"] : "",
    product.ContainsKey("ImagePaths4") ? product["ImagePaths4"] : ""
        };

            string[] inputIds = { "file1", "file2", "file3", "file4" };

            var validImages = inputIds.Zip(imagePaths, (id, path) => new { id, path })
                .Where(x => !string.IsNullOrEmpty(x.path))
                .Select(x => new { x.id, x.path, element = driver.FindElements(By.Name(x.id)).FirstOrDefault() })
                .Where(x => x.element != null);

            validImages.ToList().ForEach(image =>
            {
                image.element.SendKeys(image.path);
                Console.WriteLine($"✅ Ảnh đã được tải lên: {image.path}");
            });
            Thread.Sleep(2000);


            IWebElement saveButton = driver.FindElement(By.ClassName("btn-primary"));

            actions.MoveToElement(saveButton).Click().Perform();
            Console.WriteLine("✅ Đã nhấn 'Lưu'.");

            Thread.Sleep(2000);
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams/Index");
            Thread.Sleep(2000);

            var errorMessages = driver.FindElements(By.ClassName("text-danger"));
            errorMessages.ToList().ForEach(error => Console.WriteLine($"✅ Test Passed! Báo lỗi: {error.Text}"));

            Assert.That(errorMessages.Count, Is.GreaterThan(0), "❌ Test Failed! Không có thông báo lỗi tên sản phẩm bị trùng.");
            driver.Quit();
        }

        [Test]
        public void TestXoaSanPham()
        {
            IWebDriver driver = new ChromeDriver();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);

            IWebElement productRow = driver.FindElement(By.XPath("//td[contains(text(), 'Thuốc ho')]/.."));
            IWebElement deleteButton = wait.Until(driver => productRow.FindElement(By.XPath(".//a[contains(text(), 'Xóa')]")));

            Actions actions = new Actions(driver);
            actions.MoveToElement(deleteButton).Click().Perform();


            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,6000);");
            Thread.Sleep(2000);
            IWebElement confirmDeleteButton = wait.Until(driver => driver.FindElement(By.XPath("//input[@type='submit' and @value='Delete']")));
            confirmDeleteButton.Click();
            Thread.Sleep(2000);

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams");
            Thread.Sleep(2000);

            bool isProductExist = driver.FindElements(By.XPath("//td[contains(text(), 'Thuốc ho')]")).Count > 0;
            Console.WriteLine(isProductExist ? "Sản phẩm 'Thuốc ho' vẫn còn tồn tại." : "Sản phẩm 'Thuốc ho' đã được xóa thành công.");

            driver.Quit();
        }

        [Test]
        public void TestXoaSanPhamCoTrongDonHang()
        {
            IWebDriver driver = new ChromeDriver();

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);

            IWebElement productRow = driver.FindElement(By.XPath("//td[contains(text(), 'Thuốc ho')]/.."));//Sản Phẩm đã có trong đơn hàng
            IWebElement deleteButton = productRow.FindElement(By.XPath(".//a[contains(text(), 'Xóa')]"));
            deleteButton.Click();
          

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,6000);");
            Thread.Sleep(2000);
            IWebElement confirmDeleteButton = driver.FindElement(By.XPath("//input[@type='submit' and @value='Delete']"));
            confirmDeleteButton.Click();
            Thread.Sleep(2000);

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/SanPhams");
            Thread.Sleep(2000);

            // Directly output the result without if-else
            Console.WriteLine(
                driver.FindElements(By.XPath("//td[contains(text(), 'Thuốc ho')]")).Count > 0
                ? "Sản phẩm 'Thuốc ho' vẫn còn tồn tại."
                : "Sản phẩm 'Thuốc ho' đã được xóa thành công."
            );

            driver.Quit();
        }


        //public List<Dictionary<string, string>> ReadExcelData(string filePath, string sheetName)
        //{
        //    var result = new List<Dictionary<string, string>>();

        //    // Mở file Excel
        //    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //    {
        //        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        //        using (var reader = ExcelReaderFactory.CreateReader(stream))
        //        {
        //            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        //            {
        //                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
        //                {
        //                    UseHeaderRow = true
        //                }
        //            });

        //            // Kiểm tra nếu sheet tồn tại
        //            var dataTable = dataSet.Tables[sheetName];
        //            if (dataTable == null)
        //            {
        //                throw new Exception($"Không tìm thấy sheet: {sheetName}");
        //            }

        //            // Đọc dữ liệu từ sheet
        //            foreach (DataRow row in dataTable.Rows)
        //            {
        //                var dict = new Dictionary<string, string>();
        //                foreach (DataColumn col in dataTable.Columns)
        //                {
        //                    dict[col.ColumnName] = row[col]?.ToString() ?? "";
        //                }
        //                result.Add(dict);
        //            }
        //        }
        //    }
        //    return result;
        //}


        [TearDown]
        public void TearDown()
        {
            driver.Dispose();
        }
        public List<Dictionary<string, string>> ReadExcel(string filePath, string sheetName, bool readSheet3 = false)
        {
            List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
            List<Dictionary<string, string>> dataSheet3 = new List<Dictionary<string, string>>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    // ✅ Đọc dữ liệu từ sheet được chỉ định (mặc định)
                    var table = result.Tables[sheetName];
                    if (table != null)
                    {
                        data = ExtractDataFromSheet(table);
                    }

                    // ✅ Đọc dữ liệu từ Sheet 3 nếu được yêu cầu
                    if (readSheet3 && result.Tables.Count >= 3)
                    {
                        var sheet3 = result.Tables[2]; // Sheet 3 có index là 2
                        if (sheet3 != null)
                        {
                            dataSheet3 = ExtractDataFromSheet(sheet3);
                        }
                    }
                }
            }

            // Nếu chỉ cần đọc sheet 3, trả về data của sheet 3
            return readSheet3 ? dataSheet3 : data;
        }

        // ✅ Tách logic đọc sheet thành hàm riêng để tái sử dụng
        private List<Dictionary<string, string>> ExtractDataFromSheet(DataTable table)
        {
            List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
            string[] headers = new string[table.Columns.Count];

            // Lấy tiêu đề cột từ dòng đầu tiên (row 0)
            for (int col = 0; col < table.Columns.Count; col++)
            {
                headers[col] = table.Rows[0][col].ToString();
            }

            // Lấy dữ liệu từ dòng 2 trở đi (row index 1 trở đi)
            for (int row = 1; row < table.Rows.Count; row++)
            {
                Dictionary<string, string> rowData = new Dictionary<string, string>();
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    rowData[headers[col]] = table.Rows[row][col].ToString();
                }
                data.Add(rowData);
            }

            return data;
        }

    }
}