using ExcelDataReader;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Data;

namespace SeleniumTests
{
    public class RegisterTests
    {
        private ChromeDriver driver;
        private string baseUrl = "https://localhost:7053/"; private string baseUrldk = "https://localhost:7053/user/register";private string baseUrldn = "https://localhost:7053/User/Login";
        private List<Dictionary<string, string>> loginData;
        private List<Dictionary<string, string>> dangnhapData;
        private List<Dictionary<string, string>> dangxuatData;
        private string filePath;
        private readonly object successMessage;
        public object ExpectedConditions { get; private set; }

        [SetUp]
        public void Setup()
        {
            string filePath = @"D:\Register_WithExcelData.xlsx";
            loginData = ReadExcel(filePath, "dangky");
            dangnhapData = ReadExcel(filePath, "dangnhap");
            dangxuatData = ReadExcel(filePath, "dangxuat");
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(baseUrl);
        }
        [Test]
        public void Test_OpenWeb()
        {
            Assert.That(driver.Title, Is.EqualTo("NHATHUOCSO1VIETNAM"));
        }

        [Test]
        public void Test_DangKy_DangNhap_DangXuat()
        {
            var testDataDK = loginData.FirstOrDefault();
            var testDataDN = dangnhapData.FirstOrDefault();

            if (testDataDK != null)
            {
                driver.Navigate().GoToUrl(baseUrldk);
                Thread.Sleep(2000);

                // Nhập dữ liệu vào form đăng ký
                driver.FindElement(By.Name("HoTen")).SendKeys(testDataDK["HoTen"]);
                driver.FindElement(By.Name("Username")).SendKeys(testDataDK["Username"]);
                driver.FindElement(By.Name("Email")).SendKeys(testDataDK["Email"]);
                driver.FindElement(By.Name("Sdt")).SendKeys(testDataDK["Sdt"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(testDataDK["Matkhau"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(Keys.Enter);

                Thread.Sleep(2000);
            }

            if (testDataDN != null)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(2000);

                // Nhập thông tin đăng nhập
                driver.FindElement(By.Name("Username")).SendKeys(testDataDN["Username"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(testDataDN["Password"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(Keys.Enter);

                Thread.Sleep(3000);

                DangXuat();
                Thread.Sleep(2000);
            }
        }

        [Test]
        public void TestDang_Ky_Thanh_Cong()
        {
            var testDataDK = loginData.FirstOrDefault();
            if (testDataDK != null)
            {
                driver.Navigate().GoToUrl(baseUrldk);
                Thread.Sleep(2000);

                // Nhập dữ liệu vào form đăng ký
                driver.FindElement(By.Name("HoTen")).SendKeys(testDataDK["HoTen"]);
                driver.FindElement(By.Name("Username")).SendKeys(testDataDK["Username"]);
                driver.FindElement(By.Name("Email")).SendKeys(testDataDK["Email"]);
                driver.FindElement(By.Name("Sdt")).SendKeys(testDataDK["Sdt"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(testDataDK["Matkhau"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(Keys.Enter);

                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDang_ky_Email_Trong()
        {
            var testData = loginData.Skip(2).Take(4).ToList();

            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldk); // Chuyển đến trang đăng ký
                Thread.Sleep(2000);

                // **Nhập dữ liệu vào form đăng ký**
                IWebElement HoTenInput = driver.FindElement(By.Name("HoTen"));
                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement EmailInput = driver.FindElement(By.Name("Email"));
                IWebElement SdtInput = driver.FindElement(By.Name("Sdt"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                HoTenInput.SendKeys(data["HoTen"]);
                UsernameInput.SendKeys(data["Username"]);
                EmailInput.SendKeys(data["Email"]);
                SdtInput.SendKeys(data["Sdt"]);
                MatkhauInput.SendKeys(data["Matkhau"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(2000);

                // **Kiểm tra kết quả đăng ký**
                // Kiểm tra phần tử hiển thị
                Assert.That(successMessage, Is.True, "Đăng ký không thành công, Email trống cần nhập.");

                // Kiểm tra nội dung của thông báo (giả sử có nội dung cụ thể)
                Assert.That(successMessage, Is.EqualTo("Đăng ký thành công!"), "Thông báo đăng ký không đúng.");

            }
        }
        [Test]
        public void TestDangKy_tên_đăng_nhập_trùng()
        {
            var testData = loginData.Skip(1).Take(3).ToList(); // Bỏ qua 2 dòng đầu, lấy 4 dòng tiếp theo

            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldk); // Chuyển đến trang đăng ký
                Thread.Sleep(2000);

                // **Nhập dữ liệu vào form đăng ký**
                IWebElement HoTenInput = driver.FindElement(By.Name("HoTen"));
                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement EmailInput = driver.FindElement(By.Name("Email"));
                IWebElement SdtInput = driver.FindElement(By.Name("Sdt"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                HoTenInput.SendKeys(data["HoTen"]);
                UsernameInput.SendKeys(data["Username"]);
                EmailInput.SendKeys(data["Email"]);
                SdtInput.SendKeys(data["Sdt"]);
                MatkhauInput.SendKeys(data["Matkhau"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(2000);

                // **Kiểm tra kết quả đăng ký**
                // Kiểm tra phần tử hiển thị
                Assert.That(successMessage, Is.True, "Đăng ký không thành công, tên đăng nhập trùng.");

                // Kiểm tra nội dung của thông báo (giả sử có nội dung cụ thể)
                Assert.That(successMessage, Is.EqualTo("Đăng ký thành công!"), "Thông báo đăng ký không đúng.");

            }
        }
        [Test]
        public void TestDangKy_mật_khẩu_yếu()
        {
            var testData = loginData.Skip(3).Take(5).ToList(); // Bỏ qua 2 dòng đầu, lấy 4 dòng tiếp theo

            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldk); // Chuyển đến trang đăng ký
                Thread.Sleep(2000);

                // **Nhập dữ liệu vào form đăng ký**
                IWebElement HoTenInput = driver.FindElement(By.Name("HoTen"));
                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement EmailInput = driver.FindElement(By.Name("Email"));
                IWebElement SdtInput = driver.FindElement(By.Name("Sdt"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                HoTenInput.SendKeys(data["HoTen"]);
                UsernameInput.SendKeys(data["Username"]);
                EmailInput.SendKeys(data["Email"]);
                SdtInput.SendKeys(data["Sdt"]);
                MatkhauInput.SendKeys(data["Matkhau"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(2000);

                // **Kiểm tra kết quả đăng ký**
                // Kiểm tra phần tử hiển thị
                Assert.That(successMessage, Is.True, "Đăng ký không thành công, mật khẩu yếu.");

                // Kiểm tra nội dung của thông báo (giả sử có nội dung cụ thể)
                Assert.That(successMessage, Is.EqualTo("Đăng ký thành công!"), "Thông báo đăng ký không đúng.");

            }
        }
        [Test]
        public void TestDangKyHọTen_chưa_nhập()
        {
            var testData = loginData.Skip(4).Take(6).ToList(); // Bỏ qua 2 dòng đầu, lấy 4 dòng tiếp theo

            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldk); // Chuyển đến trang đăng ký
                Thread.Sleep(2000);

                // **Nhập dữ liệu vào form đăng ký**
                IWebElement HoTenInput = driver.FindElement(By.Name("HoTen"));
                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement EmailInput = driver.FindElement(By.Name("Email"));
                IWebElement SdtInput = driver.FindElement(By.Name("Sdt"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                HoTenInput.SendKeys(data["HoTen"]);
                UsernameInput.SendKeys(data["Username"]);
                EmailInput.SendKeys(data["Email"]);
                SdtInput.SendKeys(data["Sdt"]);
                MatkhauInput.SendKeys(data["Matkhau"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(2000);

                // **Kiểm tra kết quả đăng ký**
                // Kiểm tra phần tử hiển thị
                Assert.That(successMessage, Is.True, "Đăng ký không thành công, HọTen chưa nhập.");

                // Kiểm tra nội dung của thông báo (giả sử có nội dung cụ thể)
                Assert.That(successMessage, Is.EqualTo("Đăng ký thành công!"), "Thông báo đăng ký không đúng.");

            }
        }
        [Test]
        public void TestDangKy_matkhaucokytudatbiet()
        {
            var testData = loginData.Skip(5).Take(7).ToList(); // Bỏ qua 2 dòng đầu, lấy 4 dòng tiếp theo

            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldk); // Chuyển đến trang đăng ký
                Thread.Sleep(2000);

                // **Nhập dữ liệu vào form đăng ký**
                IWebElement HoTenInput = driver.FindElement(By.Name("HoTen"));
                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement EmailInput = driver.FindElement(By.Name("Email"));
                IWebElement SdtInput = driver.FindElement(By.Name("Sdt"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                HoTenInput.SendKeys(data["HoTen"]);
                UsernameInput.SendKeys(data["Username"]);
                EmailInput.SendKeys(data["Email"]);
                SdtInput.SendKeys(data["Sdt"]);
                MatkhauInput.SendKeys(data["Matkhau"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(2000);

                // **Kiểm tra kết quả đăng ký**
                // Kiểm tra phần tử hiển thị
                Assert.That(successMessage, Is.True, "Đăng ký không thành công, mật khẩu có chứa ký tự đặt biệt.");

                // Kiểm tra nội dung của thông báo (giả sử có nội dung cụ thể)
                Assert.That(successMessage, Is.EqualTo("Đăng ký thành công!"), "Thông báo đăng ký không đúng.");

            }
        }
        [Test]
        public void TestDangKy_usernameco_kytu_datbiet()
        {
            var testData = loginData.Skip(6).Take(8).ToList(); // Bỏ qua 2 dòng đầu, lấy 4 dòng tiếp theo

            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldk); // Chuyển đến trang đăng ký
                Thread.Sleep(2000);

                // **Nhập dữ liệu vào form đăng ký**
                IWebElement HoTenInput = driver.FindElement(By.Name("HoTen"));
                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement EmailInput = driver.FindElement(By.Name("Email"));
                IWebElement SdtInput = driver.FindElement(By.Name("Sdt"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                HoTenInput.SendKeys(data["HoTen"]);
                UsernameInput.SendKeys(data["Username"]);
                EmailInput.SendKeys(data["Email"]);
                SdtInput.SendKeys(data["Sdt"]);
                MatkhauInput.SendKeys(data["Matkhau"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(2000);

                // **Kiểm tra kết quả đăng ký**
                // Kiểm tra phần tử hiển thị
                Assert.That(successMessage, Is.True, "Đăng ký không thành công, HọTen chưa nhập.");

                // Kiểm tra nội dung của thông báo (giả sử có nội dung cụ thể)
                Assert.That(successMessage, Is.EqualTo("Đăng ký thành công!"), "Thông báo đăng ký không đúng.");

            }
        }
        [Test]
        public void TestDang_Nhap_Thanh_Cong()
        {
            var testDataDN = dangnhapData.FirstOrDefault();
            if (testDataDN != null)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(2000);

                // Nhập thông tin đăng nhập
                driver.FindElement(By.Name("Username")).SendKeys(testDataDN["Username"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(testDataDN["Password"]);
                driver.FindElement(By.Name("Matkhau")).SendKeys(Keys.Enter);

                Thread.Sleep(3000);

                DangXuat();
                Thread.Sleep(2000);
            }
        }

        [Test]
        public void TestDangNhap_matkhaukhongdung()
        {
            var testData = dangnhapData.Skip(1).Take(3).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(3000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, mật khẩu không đúng.");

                DangXuat();
                Thread.Sleep(2000);

                //// Đăng nhập lại sau khi đăng xuất
                //driver.Navigate().GoToUrl(baseUrldn);
                //Thread.Sleep(2000);

                //IWebElement UsernameInput2 = driver.FindElement(By.Name("Username"));
                //IWebElement MatkhauInput2 = driver.FindElement(By.Name("Matkhau"));

                //UsernameInput2.SendKeys(data["Username"]);
                //MatkhauInput2.SendKeys(data["Password"]);
                //MatkhauInput2.SendKeys(Keys.Enter);

                //Thread.Sleep(3000);
                //Console.WriteLine($"✅ Đăng nhập lại thành công với {data["Username"]}");
            }
        }

        [Test]
        public void TestDangNhapUsernamecokytudatbiet()
        {
            var testData = dangnhapData.Skip(2).Take(4).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(2000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, Username có ký tự đặc biệt.");

                DangXuat();
                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDangNhapmatkhaucokytudatbiet()
        {
            var testData = dangnhapData.Skip(3).Take(5).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(2000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, mật khẩu có ký tự đặc biệt.");

                DangXuat();
                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDangNhap_usernamebotrong()
        {
            var testData = dangnhapData.Skip(4).Take(6).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(3000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, username không đc bỏ trống.");

                DangXuat();
                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDangNhap_matkhautrong()
        {
            var testData = dangnhapData.Skip(5).Take(7).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(3000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, mật khẩu bỏ trống.");

                DangXuat();
                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDangNhap_botrongca2()
        {
            var testData = dangnhapData.Skip(6).Take(8).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(3000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, username và mật khẩu không đc bỏ trống.");

                DangXuat();
                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDangNhap_taikhoanbikhoa()
        {
            var testData = dangnhapData.Skip(7).Take(9).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(3000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, tài khoản này đã bị khóa.");

                DangXuat();
                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDangNhap_nhapsaimatkhau_username()
        {
            var testData = dangnhapData.Skip(8).Take(10).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(3000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                MatkhauInput.SendKeys(Keys.Enter);

                Thread.Sleep(3000);
                Assert.That(successMessage, Is.True, "Đăng nhập không thành công, tên và mật khẩu không đúng.");

                DangXuat();
                Thread.Sleep(2000);
            }
        }
        [Test]
        public void TestDangxuat_ThanhCong()
        {
            var testData = dangxuatData.Take(2).ToList();
            foreach (var data in testData)
            {
                driver.Navigate().GoToUrl(baseUrldn);
                Thread.Sleep(3000);

                IWebElement UsernameInput = driver.FindElement(By.Name("Username"));
                IWebElement MatkhauInput = driver.FindElement(By.Name("Matkhau"));
                IWebElement LoginButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));

                UsernameInput.SendKeys(data["Username"]);
                MatkhauInput.SendKeys(data["Password"]);
                LoginButton.Click();

                Thread.Sleep(3000);

                // Kiểm tra đăng nhập thành công bằng cách xác minh phần tử nào đó trên trang
                IWebElement logoutButton = driver.FindElement(By.CssSelector("a[href='/user/dangxuat']"));
               

                // Thực hiện đăng xuất
                DangXuat();
                Thread.Sleep(2000);
            }
        }

        [Test]
        public void TestDangXuat_KhiChuaDangNhap()
        {
            driver.Navigate().GoToUrl(baseUrl);
            Thread.Sleep(2000);

            var logoutButtons = driver.FindElements(By.CssSelector("a[href='/user/dangxuat']"));
            Assert.That(logoutButtons.Count == 0, "đăng xuất không thành công do chưa đăng nhập");
        }

        [TearDown]
        public void TearDown()
        {
            driver.Dispose();
        }

        private void DangNhap(string Username, string password)
        {
            driver.Navigate().GoToUrl($"{baseUrl}user/login");

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // Chờ ô nhập Username xuất hiện
            IWebElement usernameInput = wait.Until(d => d.FindElement(By.Id("Username")));
            usernameInput.SendKeys(Username);

            // Chờ ô nhập Password xuất hiện
            IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));
            passwordInput.SendKeys(password);

            // Nhấn nút đăng nhập
            IWebElement loginButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            loginButton.Click();

            // Chờ điều hướng thành công (ví dụ: về trang chủ)
            wait.Until(d => !d.Url.Contains("login"));

            Console.WriteLine($"✅ Đăng nhập thành công với tài khoản: {Username}");
        }

        private void DangXuat()
        {
            Actions actions = new Actions(driver);

            // Tìm icon người dùng và hover chuột vào đó
            IWebElement userIcon = driver.FindElement(By.CssSelector("i.bx-user-check"));
            actions.MoveToElement(userIcon).Perform();
            Thread.Sleep(3000); // Chờ menu hiển thị

            // Nhấn vào nút đăng xuất
            IWebElement logoutButton = driver.FindElement(By.CssSelector("a[href='/user/dangxuat']"));
            logoutButton.Click();
            Thread.Sleep(3000);
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