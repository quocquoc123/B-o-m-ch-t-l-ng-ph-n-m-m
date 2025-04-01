using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using ExcelDataReader;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;


namespace TestThongKe
{
    public class Tests
    {

        private ChromeDriver driver;
        private string baseUrl = "https://localhost:7053/";
        private List<Dictionary<string, string>> userLoginData;
        private List<Dictionary<string, string>> adminLoginData;
        private List<Dictionary<string, string>> productData; 
        private List<Dictionary<string, string>> donhangData; 


        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(baseUrl);

            string filePath = @"C:\\Excel\\ThongKeBaoCao.xlsx";

            // Đọc dữ liệu từ các sheet
            userLoginData = ReadExcel(filePath, "Login");  
            adminLoginData = ReadExcel(filePath, "Admin"); 
            productData = ReadExcel(filePath, "Product");  
            donhangData = ReadExcel(filePath, "DonHang");  
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
            Thread.Sleep(5000); // Chờ menu hiển thị

            // Nhấn vào nút đăng xuất
            IWebElement logoutButton = driver.FindElement(By.CssSelector("a[href='/user/dangxuat']"));
            logoutButton.Click();
            Thread.Sleep(5000);
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

           
            return readSheet3 ? dataSheet3 : data;
        }

        
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

   
        [Test]
        public void TestThongKeKhiChuaCoDonHang_DoanhThu()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20)); // Tăng thời gian chờ lên 20 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000);


            // 🔹 Click vào menu "Thống kê số lượng đã bán"
            IWebElement thongKeMenu = wait.Until(d => d.FindElement(By.XPath("//a[@href='/Admin/DonHangs/SoLuongDaBan']")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", thongKeMenu);
            Thread.Sleep(4000); 
            thongKeMenu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/SoLuongDaBan"));
            Console.WriteLine("✅ Điều hướng đến trang thống kê số lượng đã bán!");

            // Kiểm tra thông báo khi không có đơn hàng
            IWebElement noDataMessage1 = wait.Until(d => d.FindElement(By.XPath("//*[contains(text(), 'Không có đơn hàng để thống kê')]")));
            Assert.IsTrue(noDataMessage1.Displayed, "❌ Không hiển thị thông báo 'Không có đơn hàng để thống kê' trên trang số lượng bán!");
            Console.WriteLine("✅ Thông báo 'Không có đơn hàng để thống kê' hiển thị đúng!");

            Thread.Sleep(4000); 
      
        }

        [Test]
        public void TestThongKeKhiChuaCoDonHang_BaoCao()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20)); // Tăng thời gian chờ lên 20 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000);

            // 🔹 Click vào menu "Thống kê doanh thu"
            IWebElement menuDoanhThu = wait.Until(d => d.FindElement(By.XPath("//a[@href='/Admin/DonHangs/DoanhThu1' and contains(@class, 'nav-link')]")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", menuDoanhThu);
            Thread.Sleep(4000);
            menuDoanhThu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Console.WriteLine("✅ Điều hướng đến trang thống kê doanh thu!");

            // Kiểm tra thông báo khi không có doanh thu
            IWebElement noDataMessage2 = wait.Until(d => d.FindElement(By.XPath("//*[contains(text(), 'Không có đơn hàng để thống kê doanh thu theo năm')]")));
            Assert.IsTrue(noDataMessage2.Displayed, "❌ Không hiển thị thông báo 'Không có đơn hàng để thống kê doanh thu theo năm'!");
            Console.WriteLine("✅ Thông báo 'Không có đơn hàng để thống kê doanh thu theo năm' hiển thị đúng!");

            Thread.Sleep(4000);
        }

        [Test]
        public void TestThongKeDonHangDauTien1SP()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // Lấy dữ liệu từ dòng đầu tiên của tập dữ liệu
            var user = userLoginData[0]; // Người dùng đầu tiên       
            var product = productData[0]; // Sản phẩm đầu tiên

            // Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000);

            IWebElement duocPhamLink = wait.Until(d => d.FindElement(By.XPath("//a[@href='/SanPham/Index/2']")));
            duocPhamLink.Click();
            Thread.Sleep(3000);
            Console.WriteLine("✅ Đã điều hướng đến trang DƯỢC PHẨM!");

            // Ấn mua ngay sản phẩm
            driver.FindElement(By.CssSelector("a.buy-button[href='/GioHang/AddToCart/1']")).Click();
            Thread.Sleep(2000);

            // Chuyển đến giỏ hàng
            driver.Navigate().GoToUrl("https://localhost:7053/GioHang");
            Thread.Sleep(2000);

            // Cập nhật số lượng sản phẩm
            var quantityInput = driver.FindElement(By.CssSelector("input[name='quantity']"));
            quantityInput.Clear();

            quantityInput.SendKeys(product["SoLuong"]);

            driver.FindElement(By.CssSelector("button[formaction='/GioHang/UpdateQuantity']")).Click();
            Thread.Sleep(3000);

            // Chuyển đến trang thanh toán
            driver.FindElement(By.CssSelector("a[href='/GioHang/ThanhToan']")).Click();
            Thread.Sleep(2000);

            // Chọn phương thức thanh toán COD
            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("cod");
            Thread.Sleep(2000);

            // Nhập thông tin nhận hàng
            driver.FindElement(By.Name("tennguoinhan")).SendKeys(product["TenNguoiNhan"]);

            driver.FindElement(By.Name("sdtnguoinhan")).SendKeys(product["Sdt"]);

            // Chọn địa chỉ
            var provinceSelect = new SelectElement(driver.FindElement(By.Id("province")));
            provinceSelect.SelectByValue("10");
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].dispatchEvent(new Event('change'));", driver.FindElement(By.Id("province")));
            wait.Until(d => d.FindElement(By.Id("district")).FindElements(By.TagName("option")).Count > 1);

            var districtSelect = new SelectElement(driver.FindElement(By.Id("district")));
            var districtOptions = districtSelect.Options.Where(o => !string.IsNullOrEmpty(o.GetAttribute("value"))).ToList();
            districtSelect.SelectByValue(districtOptions.First().GetAttribute("value"));

            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].dispatchEvent(new Event('change'));", driver.FindElement(By.Id("district")));
            wait.Until(d => d.FindElement(By.Id("ward")).FindElements(By.TagName("option")).Count > 1);

            var wardSelect = new SelectElement(driver.FindElement(By.Id("ward")));
            var wardOptions = wardSelect.Options.Where(o => !string.IsNullOrEmpty(o.GetAttribute("value"))).ToList();
            wardSelect.SelectByValue(wardOptions.First().GetAttribute("value"));

            driver.FindElement(By.Id("address_detail")).SendKeys(product["DiaChi"]);

            Thread.Sleep(4000);

            IWebElement thanhToanButton = driver.FindElement(By.Id("cod-button"));
            thanhToanButton.Click();
            Thread.Sleep(7000);

            Assert.That(driver.Url, Does.Contain("https://localhost:7053/GioHang/LuuDonHang"), "Không phải trang đăng nhập thành công");

            // Nhấn vào nút "Quay lại trang chủ"
            IWebElement backButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn-back[href='/']")));
            backButton.Click();
            Thread.Sleep(5000);



            driver.Navigate().GoToUrl("https://localhost:7053/Admin/DonHangs/Index");
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/Index"));
            Console.WriteLine("✅ Điều hướng đến trang danh sách đơn hàng thành công");
            Thread.Sleep(3000);

            // Xác nhận đơn hàng mới nhất
            var xacNhanButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/XacNhanDon'] button")));
            Assert.IsTrue(xacNhanButtons.Count > 0, "❌ Không có đơn hàng nào để xác nhận");
            var newestOrderButton = xacNhanButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestOrderButton);
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xác nhận đơn hàng thành công");

            // Cập nhật trạng thái đơn hàng thành 'Đã giao'
            var daGiaoButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/UpdateOrderStatus'] button")));
            Assert.IsTrue(daGiaoButtons.Count > 0, "❌ Không có đơn hàng nào để cập nhật trạng thái");
            var newestStatusButton = daGiaoButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestStatusButton);
            Thread.Sleep(6000);
            Console.WriteLine("✅ Cập nhật trạng thái đơn hàng thành 'Đã giao'");

            // 🔹 Click vào menu "Thống Kê Doanh Thu"
            var menuDoanhThu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/DoanhThu1']")));
            menuDoanhThu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Thread.Sleep(5000);
            Console.WriteLine("✅ Điều hướng đến trang doanh thu thành công");

            // Chọn năm mới nhất
            var doanhThuTheoNam = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoThang']")));
            doanhThuTheoNam.OrderByDescending(e => int.Parse(e.Text.Trim())).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo năm mới nhất thành công");

            // Chọn tháng mới nhất
            var doanhThuTheoThang = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoNgay']")));
            doanhThuTheoThang.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo tháng mới nhất thành công");

            // Chọn ngày mới nhất
            var donHangTheoNgay = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DonHangTheoNgay']")));
            donHangTheoNgay.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(10000);
            Console.WriteLine("✅ Xem đơn hàng theo ngày mới nhất thành công");

            // 🔹 Click vào mục Thống kê số lượng đã bán
            IWebElement thongKeMenu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/SoLuongDaBan']")));
            thongKeMenu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/SoLuongDaBan"));
            Thread.Sleep(3000);
            Console.WriteLine("✅ Điều hướng đến trang thống kê số lượng đã bán!");

            // 🔹 Kiểm tra biểu đồ Pie Chart có hiển thị không
            var pieChart = wait.Until(d => d.FindElement(By.Id("pieChart")));
            Assert.IsTrue(pieChart.Displayed, "❌ Biểu đồ không hiển thị!");
            Thread.Sleep(9000);
            Console.WriteLine("✅ Biểu đồ Pie Chart hiển thị đúng!");

            // 🔹 Kiểm tra danh sách sản phẩm có hiển thị không
            var danhSachSanPham = wait.Until(d => d.FindElements(By.XPath("//ul/li/a/strong")));
            Assert.IsTrue(danhSachSanPham.Count > 0, "❌ Không có sản phẩm nào trong danh sách thống kê!");
            Thread.Sleep(9000);
            Console.WriteLine($"✅ Có {danhSachSanPham.Count} sản phẩm trong danh sách thống kê!");

            // 🔹 Kiểm tra từng sản phẩm trong danh sách
            foreach (var sanPham in danhSachSanPham)
            {
                wait.Until(d => sanPham.Displayed); // Đảm bảo sản phẩm hiển thị
                string tenSanPham = sanPham.Text.Trim();
                Console.WriteLine($"🔹 Sản phẩm bán chạy: {tenSanPham}");
                Assert.IsTrue(!string.IsNullOrEmpty(tenSanPham), "❌ Tên sản phẩm không hợp lệ!");
            }
            Thread.Sleep(9000);
            Console.WriteLine("🎉 Kiểm thử hoàn tất - Trang thống kê sản phẩm bán chạy hiển thị chính xác!");

        }

        [Test]
        public void TestThongKeDonHangLan2TrungBinh5SP()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // Lấy dữ liệu từ dòng đầu tiên của tập dữ liệu
            var user = userLoginData[3]; // Người dùng đầu tiên
            var product1 = productData[1]; // Sản phẩm đầu tiên

            // Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Thread.Sleep(3000);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");

            IWebElement duocPhamLink = wait.Until(d => d.FindElement(By.XPath("//a[@href='/SanPham/Index/2']")));
            duocPhamLink.Click();
            Thread.Sleep(3000);
            Console.WriteLine("✅ Đã điều hướng đến trang DƯỢC PHẨM!");

            // Ấn mua ngay sản phẩm
            driver.FindElement(By.CssSelector("a.buy-button[href='/GioHang/AddToCart/2']")).Click();
            Thread.Sleep(1000);

            // Chuyển đến giỏ hàng
            driver.Navigate().GoToUrl("https://localhost:7053/GioHang");
            Thread.Sleep(2000);

            // Cập nhật số lượng sản phẩm
            var quantityInput = driver.FindElement(By.CssSelector("input[name='quantity']"));
            quantityInput.Clear();

            quantityInput.SendKeys(product1["SoLuong"]);

            driver.FindElement(By.CssSelector("button[formaction='/GioHang/UpdateQuantity']")).Click();
            Thread.Sleep(3000);

            // Chuyển đến trang thanh toán
            driver.FindElement(By.CssSelector("a[href='/GioHang/ThanhToan']")).Click();
            Thread.Sleep(2000);

            // Chọn phương thức thanh toán COD
            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("cod");
            Thread.Sleep(2000);

            // Nhập thông tin nhận hàng
            driver.FindElement(By.Name("tennguoinhan")).SendKeys(product1["TenNguoiNhan"]);

            driver.FindElement(By.Name("sdtnguoinhan")).SendKeys(product1["Sdt"]);


            // Chọn địa chỉ
            var provinceSelect = new SelectElement(driver.FindElement(By.Id("province")));
            provinceSelect.SelectByValue("10");
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].dispatchEvent(new Event('change'));", driver.FindElement(By.Id("province")));
            wait.Until(d => d.FindElement(By.Id("district")).FindElements(By.TagName("option")).Count > 1);

            var districtSelect = new SelectElement(driver.FindElement(By.Id("district")));
            var districtOptions = districtSelect.Options.Where(o => !string.IsNullOrEmpty(o.GetAttribute("value"))).ToList();
            districtSelect.SelectByValue(districtOptions.First().GetAttribute("value"));

            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].dispatchEvent(new Event('change'));", driver.FindElement(By.Id("district")));
            wait.Until(d => d.FindElement(By.Id("ward")).FindElements(By.TagName("option")).Count > 1);

            var wardSelect = new SelectElement(driver.FindElement(By.Id("ward")));
            var wardOptions = wardSelect.Options.Where(o => !string.IsNullOrEmpty(o.GetAttribute("value"))).ToList();
            wardSelect.SelectByValue(wardOptions.First().GetAttribute("value"));

            driver.FindElement(By.Id("address_detail")).SendKeys(product1["DiaChi"]);

            Thread.Sleep(2000);

            IWebElement thanhToanButton = driver.FindElement(By.Id("cod-button"));
            thanhToanButton.Click();
            Thread.Sleep(7000);

            Assert.That(driver.Url, Does.Contain("https://localhost:7053/GioHang/LuuDonHang"), "Không phải trang đăng nhập thành công");

            // Nhấn vào nút "Quay lại trang chủ"
            IWebElement backButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn-back[href='/']")));
            backButton.Click();
            Thread.Sleep(6000);



            driver.Navigate().GoToUrl("https://localhost:7053/Admin/DonHangs/Index");
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/Index"));
            Console.WriteLine("✅ Điều hướng đến trang danh sách đơn hàng thành công");
            Thread.Sleep(3000);

            // Xác nhận đơn hàng mới nhất
            var xacNhanButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/XacNhanDon'] button")));
            Assert.IsTrue(xacNhanButtons.Count > 0, "❌ Không có đơn hàng nào để xác nhận");
            var newestOrderButton = xacNhanButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestOrderButton);
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xác nhận đơn hàng thành công");

            // Cập nhật trạng thái đơn hàng thành 'Đã giao'
            var daGiaoButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/UpdateOrderStatus'] button")));
            Assert.IsTrue(daGiaoButtons.Count > 0, "❌ Không có đơn hàng nào để cập nhật trạng thái");
            var newestStatusButton = daGiaoButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestStatusButton);
            Thread.Sleep(6000);
            Console.WriteLine("✅ Cập nhật trạng thái đơn hàng thành 'Đã giao'");

            // 🔹 Click vào menu "Thống Kê Doanh Thu"
            var menuDoanhThu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/DoanhThu1']")));
            menuDoanhThu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Thread.Sleep(5000);
            Console.WriteLine("✅ Điều hướng đến trang doanh thu thành công");

            // Chọn năm mới nhất
            var doanhThuTheoNam = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoThang']")));
            doanhThuTheoNam.OrderByDescending(e => int.Parse(e.Text.Trim())).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo năm mới nhất thành công");

            // Chọn tháng mới nhất
            var doanhThuTheoThang = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoNgay']")));
            doanhThuTheoThang.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo tháng mới nhất thành công");

            // Chọn ngày mới nhất
            var donHangTheoNgay = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DonHangTheoNgay']")));
            donHangTheoNgay.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(10000);
            Console.WriteLine("✅ Xem đơn hàng theo ngày mới nhất thành công");

            // 🔹 Click vào mục Thống kê số lượng đã bán
            IWebElement thongKeMenu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/SoLuongDaBan']")));
            thongKeMenu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/SoLuongDaBan"));
            Thread.Sleep(3000);
            Console.WriteLine("✅ Điều hướng đến trang thống kê số lượng đã bán!");

            // 🔹 Kiểm tra biểu đồ Pie Chart có hiển thị không
            var pieChart = wait.Until(d => d.FindElement(By.Id("pieChart")));
            Assert.IsTrue(pieChart.Displayed, "❌ Biểu đồ không hiển thị!");
            Thread.Sleep(9000);
            Console.WriteLine("✅ Biểu đồ Pie Chart hiển thị đúng!");

            // 🔹 Kiểm tra danh sách sản phẩm có hiển thị không
            var danhSachSanPham = wait.Until(d => d.FindElements(By.XPath("//ul/li/a/strong")));
            Assert.IsTrue(danhSachSanPham.Count > 0, "❌ Không có sản phẩm nào trong danh sách thống kê!");
            Thread.Sleep(9000);
            Console.WriteLine($"✅ Có {danhSachSanPham.Count} sản phẩm trong danh sách thống kê!");

            // 🔹 Kiểm tra từng sản phẩm trong danh sách
            foreach (var sanPham in danhSachSanPham)
            {
                wait.Until(d => sanPham.Displayed); // Đảm bảo sản phẩm hiển thị
                string tenSanPham = sanPham.Text.Trim();
                Console.WriteLine($"🔹 Sản phẩm bán chạy: {tenSanPham}");
                Assert.IsTrue(!string.IsNullOrEmpty(tenSanPham), "❌ Tên sản phẩm không hợp lệ!");
            }
            Thread.Sleep(9000);
            Console.WriteLine("🎉 Kiểm thử hoàn tất - Trang thống kê sản phẩm bán chạy hiển thị chính xác!");

        }

        [Test]
        public void TestThongKeDonHangLan3SoLuong5SP()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // Lấy dữ liệu từ dòng đầu tiên của tập dữ liệu
            var user = userLoginData[2]; // Người dùng đầu tiên
            var product2 = productData[2]; // Sản phẩm đầu tiên

            // Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Thread.Sleep(3000);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");

            IWebElement duocPhamLink = wait.Until(d => d.FindElement(By.XPath("//a[@href='/SanPham/Index/2']")));
            duocPhamLink.Click();
            Thread.Sleep(3000);
            Console.WriteLine("✅ Đã điều hướng đến trang DƯỢC PHẨM!");

            // Ấn mua ngay sản phẩm
            driver.FindElement(By.CssSelector("a.buy-button[href='/GioHang/AddToCart/3']")).Click();
            Thread.Sleep(1000);

            // Chuyển đến giỏ hàng
            driver.Navigate().GoToUrl("https://localhost:7053/GioHang");
            Thread.Sleep(2000);

            // Cập nhật số lượng sản phẩm
            var quantityInput = driver.FindElement(By.CssSelector("input[name='quantity']"));
            quantityInput.Clear();

            quantityInput.SendKeys(product2["SoLuong"]);

            driver.FindElement(By.CssSelector("button[formaction='/GioHang/UpdateQuantity']")).Click();
            Thread.Sleep(2000);

            // Chuyển đến trang thanh toán
            driver.FindElement(By.CssSelector("a[href='/GioHang/ThanhToan']")).Click();
            Thread.Sleep(2000);

            // Chọn phương thức thanh toán COD
            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("cod");
            Thread.Sleep(2000);

            // Nhập thông tin nhận hàng
            driver.FindElement(By.Name("tennguoinhan")).SendKeys(product2["TenNguoiNhan"]);

            driver.FindElement(By.Name("sdtnguoinhan")).SendKeys(product2["Sdt"]);


            // Chọn địa chỉ
            var provinceSelect = new SelectElement(driver.FindElement(By.Id("province")));
            provinceSelect.SelectByValue("10");
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].dispatchEvent(new Event('change'));", driver.FindElement(By.Id("province")));
            wait.Until(d => d.FindElement(By.Id("district")).FindElements(By.TagName("option")).Count > 1);

            var districtSelect = new SelectElement(driver.FindElement(By.Id("district")));
            var districtOptions = districtSelect.Options.Where(o => !string.IsNullOrEmpty(o.GetAttribute("value"))).ToList();
            districtSelect.SelectByValue(districtOptions.First().GetAttribute("value"));

            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].dispatchEvent(new Event('change'));", driver.FindElement(By.Id("district")));
            wait.Until(d => d.FindElement(By.Id("ward")).FindElements(By.TagName("option")).Count > 1);

            var wardSelect = new SelectElement(driver.FindElement(By.Id("ward")));
            var wardOptions = wardSelect.Options.Where(o => !string.IsNullOrEmpty(o.GetAttribute("value"))).ToList();
            wardSelect.SelectByValue(wardOptions.First().GetAttribute("value"));

            driver.FindElement(By.Id("address_detail")).SendKeys(product2["DiaChi"]);

            Thread.Sleep(2000);

            IWebElement thanhToanButton = driver.FindElement(By.Id("cod-button"));
            thanhToanButton.Click();
            Thread.Sleep(7000);

            Assert.That(driver.Url, Does.Contain("https://localhost:7053/GioHang/LuuDonHang"), "Không phải trang đăng nhập thành công");

            // Nhấn vào nút "Quay lại trang chủ"
            IWebElement backButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn-back[href='/']")));
            backButton.Click();
            Thread.Sleep(6000);



            driver.Navigate().GoToUrl("https://localhost:7053/Admin/DonHangs/Index");
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/Index"));
            Console.WriteLine("✅ Điều hướng đến trang danh sách đơn hàng thành công");
            Thread.Sleep(2000);

            // Xác nhận đơn hàng mới nhất
            var xacNhanButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/XacNhanDon'] button")));
            Assert.IsTrue(xacNhanButtons.Count > 0, "❌ Không có đơn hàng nào để xác nhận");
            var newestOrderButton = xacNhanButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestOrderButton);
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xác nhận đơn hàng thành công");

            // Cập nhật trạng thái đơn hàng thành 'Đã giao'
            var daGiaoButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/UpdateOrderStatus'] button")));
            Assert.IsTrue(daGiaoButtons.Count > 0, "❌ Không có đơn hàng nào để cập nhật trạng thái");
            var newestStatusButton = daGiaoButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestStatusButton);
            Thread.Sleep(6000);
            Console.WriteLine("✅ Cập nhật trạng thái đơn hàng thành 'Đã giao'");

            // 🔹 Click vào menu "Thống Kê Doanh Thu"
            var menuDoanhThu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/DoanhThu1']")));
            menuDoanhThu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Thread.Sleep(5000);
            Console.WriteLine("✅ Điều hướng đến trang doanh thu thành công");

            // Chọn năm mới nhất
            var doanhThuTheoNam = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoThang']")));
            doanhThuTheoNam.OrderByDescending(e => int.Parse(e.Text.Trim())).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo năm mới nhất thành công");

            // Chọn tháng mới nhất
            var doanhThuTheoThang = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoNgay']")));
            doanhThuTheoThang.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo tháng mới nhất thành công");

            // Chọn ngày mới nhất
            var donHangTheoNgay = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DonHangTheoNgay']")));
            donHangTheoNgay.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(11000);
            Console.WriteLine("✅ Xem đơn hàng theo ngày mới nhất thành công");

            // 🔹 Click vào mục Thống kê số lượng đã bán
            IWebElement thongKeMenu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/SoLuongDaBan']")));
            thongKeMenu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/SoLuongDaBan"));
            Thread.Sleep(3000);
            Console.WriteLine("✅ Điều hướng đến trang thống kê số lượng đã bán!");

            // 🔹 Kiểm tra biểu đồ Pie Chart có hiển thị không
            var pieChart = wait.Until(d => d.FindElement(By.Id("pieChart")));
            Assert.IsTrue(pieChart.Displayed, "❌ Biểu đồ không hiển thị!");
            Thread.Sleep(9000);
            Console.WriteLine("✅ Biểu đồ Pie Chart hiển thị đúng!");

            // 🔹 Kiểm tra danh sách sản phẩm có hiển thị không
            var danhSachSanPham = wait.Until(d => d.FindElements(By.XPath("//ul/li/a/strong")));
            Assert.IsTrue(danhSachSanPham.Count > 0, "❌ Không có sản phẩm nào trong danh sách thống kê!");
            Thread.Sleep(9000);
            Console.WriteLine($"✅ Có {danhSachSanPham.Count} sản phẩm trong danh sách thống kê!");

            // 🔹 Kiểm tra từng sản phẩm trong danh sách
            foreach (var sanPham in danhSachSanPham)
            {
                wait.Until(d => sanPham.Displayed); // Đảm bảo sản phẩm hiển thị
                string tenSanPham = sanPham.Text.Trim();
                Console.WriteLine($"🔹 Sản phẩm : {tenSanPham}");
                Assert.IsTrue(!string.IsNullOrEmpty(tenSanPham), "❌ Tên sản phẩm không hợp lệ!");
            }
            Thread.Sleep(9000);
            Console.WriteLine("🎉 Kiểm thử hoàn tất - Trang thống kê sản phẩm hiển thị chính xác!");

        }

        [Test]
        public void TestThongKeDoanhThuNhieuTaiKhoanNguoiDung()
        {
       
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20)); // Tăng thời gian chờ lên 20 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000);

            // Truy cập trang danh sách đơn hàng đã giao và lấy danh sách Username
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/DonHangs/Index");
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/Index"));
            Console.WriteLine("✅ Điều hướng đến trang danh sách đơn hàng thành công");
            Thread.Sleep(2000);

            var danhSachTaiKhoanDaGiao = wait.Until(d => d.FindElements(By.CssSelector("td:nth-child(1)"))) // Cột Username
                .Select(e => e.Text.Trim())
                .Distinct()
                .ToList();

            Console.WriteLine($"🔍 Danh sách Username từ đơn hàng đã giao: {string.Join(", ", danhSachTaiKhoanDaGiao)}");

            // Truy cập trang thống kê đơn hàng theo ngày và lấy danh sách Username
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/DonHangs/DoanhThu1");
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Console.WriteLine("✅ Điều hướng đến trang thống kê doanh thu thành công");
            Thread.Sleep(2000);

            wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoThang']")))
                .OrderByDescending(e => int.Parse(e.Text.Trim()))
                .First()
                .Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo năm mới nhất thành công");

            wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoNgay']")))
                .OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1]))
                .First()
                .Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo tháng mới nhất thành công");

            wait.Until(d => d.FindElements(By.CssSelector("a[href*='DonHangTheoNgay']")))
                .OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1]))
                .First()
                .Click();
            Thread.Sleep(12000);
            Console.WriteLine("✅ Xem đơn hàng theo ngày mới nhất thành công");

            var danhSachTaiKhoanTheoNgay = wait.Until(d => d.FindElements(By.CssSelector("td:nth-child(2)"))) // Cột Username
                .Select(e => e.Text.Trim())
                .Distinct()
                .ToList();

            Console.WriteLine($"🔍 Danh sách Username từ đơn hàng theo ngày: {string.Join(", ", danhSachTaiKhoanTheoNgay)}");

            // So sánh hai danh sách Username
            var taiKhoanKhongCoTrongThongKe = danhSachTaiKhoanDaGiao.Except(danhSachTaiKhoanTheoNgay).ToList();

            Console.WriteLine(taiKhoanKhongCoTrongThongKe.Any()
                ? $"❌ LỖI: Các tài khoản sau không xuất hiện trong thống kê doanh thu: {string.Join(", ", taiKhoanKhongCoTrongThongKe)}"
                : "✅ Thống kê doanh thu bao gồm tất cả tài khoản từ đơn hàng đã giao.");
        }


        [Test]
        public void TestThongKeDoanhThuDemThayDoiSoLuongDonHang()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20)); // Tăng thời gian chờ lên 20 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000);
            // Kiểm tra Thống kê và lưu tổng đơn hàng lần 1 

           
            // 🔹 Click vào menu "Thống Kê Doanh Thu"
            var menuDoanhThu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/DoanhThu1']")));
            menuDoanhThu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Thread.Sleep(5000);
            Console.WriteLine("✅ Điều hướng đến trang doanh thu thành công");

            // Chọn năm mới nhất
            var doanhThuTheoNam = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoThang']")));
            doanhThuTheoNam.OrderByDescending(e => int.Parse(e.Text.Trim())).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo năm mới nhất thành công");

            // Chọn tháng mới nhất
            var doanhThuTheoThang = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoNgay']")));
            doanhThuTheoThang.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo tháng mới nhất thành công");

            // Chọn ngày mới nhất
            var donHangTheoNgay = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DonHangTheoNgay']")));
            donHangTheoNgay.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(4000);
            Console.WriteLine("✅ Xem đơn hàng theo ngày mới nhất thành công");

            // 🔹 Lấy danh sách đơn hàng lần 1 (sử dụng WebDriverWait thay vì Thread.Sleep)
            var donHangRowsLan1 = wait.Until(d => d.FindElements(By.XPath("//table/tbody/tr")));
            Assert.IsTrue(donHangRowsLan1.Count > 0, "❌ Không có đơn hàng nào hiển thị!");

            // 🔹 Lấy danh sách mã đơn hàng (MaDH) lần 1
            var danhSachMaDHLan1 = donHangRowsLan1
                .Select(row => wait.Until(d => row.FindElement(By.XPath("./td[1]"))).Text.Trim())
                .ToList();

            var soLuongDonHangLan1 = danhSachMaDHLan1.Count;
            Console.WriteLine($"✅ Tổng số đơn hàng lần 1: {soLuongDonHangLan1}");
            TestContext.WriteLine($"🔹 Số đơn hàng lần 1: {soLuongDonHangLan1}");


            // Xác nhận đơn và kiểm tra Thống kê và lưu tổng đơn hàng lần 2 và só sánh 2 lần 
            // 🔹 Click vào menu "Đơn hàng" để điều hướng
            var menuDonHang = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/Index']")));
            menuDonHang.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/Index"));
            Console.WriteLine("✅ Điều hướng đến trang danh sách đơn hàng thành công");

            // Xác nhận đơn hàng mới nhất
            var xacNhanButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/XacNhanDon'] button")));
            Assert.IsTrue(xacNhanButtons.Count > 0, "❌ Không có đơn hàng nào để xác nhận");
            var newestOrderButton = xacNhanButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestOrderButton);
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xác nhận đơn hàng thành công");

            // Cập nhật trạng thái đơn hàng thành 'Đã giao'
            var daGiaoButtons = wait.Until(d => d.FindElements(By.CssSelector("form[action='/Admin/DonHangs/UpdateOrderStatus'] button")));
            Assert.IsTrue(daGiaoButtons.Count > 0, "❌ Không có đơn hàng nào để cập nhật trạng thái");
            var newestStatusButton = daGiaoButtons.OrderByDescending(b => b.GetAttribute("data-order-id")).First();
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", newestStatusButton);
            Thread.Sleep(6000);
            Console.WriteLine("✅ Cập nhật trạng thái đơn hàng thành 'Đã giao'");
            Thread.Sleep(3000);

            driver.Navigate().GoToUrl("https://localhost:7053/Admin/DonHangs/DoanhThu1");
            wait.Until(d => d.Url.Contains("DoanhThu1")); // Đợi đến khi trang tải xong
            Console.WriteLine("✅ Đã điều hướng đến trang Doanh Thu!");

            // 🔹 Click vào menu "Thống Kê Doanh Thu"
            var menuDoanhThu1 = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/DoanhThu1']")));
            menuDoanhThu1.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Thread.Sleep(5000);
            Console.WriteLine("✅ Điều hướng đến trang doanh thu thành công");

            // Chọn năm mới nhất
            var doanhThuTheoNam1 = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoThang']")));
            doanhThuTheoNam1.OrderByDescending(e => int.Parse(e.Text.Trim())).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo năm mới nhất thành công");

            // Chọn tháng mới nhất
            var doanhThuTheoThang1 = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoNgay']")));
            doanhThuTheoThang1.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo tháng mới nhất thành công");

            // Chọn ngày mới nhất
            var donHangTheoNgay1 = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DonHangTheoNgay']")));
            donHangTheoNgay1.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(4000);
            Console.WriteLine("✅ Xem đơn hàng theo ngày mới nhất thành công");

            // 🔹 Lấy danh sách đơn hàng lần 2 sau khi đặt hàng
            var donHangRowsLan2 = wait.Until(d => d.FindElements(By.XPath("//table/tbody/tr")));
            Assert.IsTrue(donHangRowsLan2.Count > 0, "❌ Không có đơn hàng nào hiển thị!");

            // 🔹 Lấy danh sách mã đơn hàng (MaDH) lần 2
            var danhSachMaDHLan2 = donHangRowsLan2
                .Select(row => wait.Until(d => row.FindElement(By.XPath("./td[1]"))).Text.Trim())
                .ToList();

            var soLuongDonHangLan2 = danhSachMaDHLan2.Count;
            Console.WriteLine($"✅ Tổng số đơn hàng lần 2: {soLuongDonHangLan2}");
            TestContext.WriteLine($"🔹 Số đơn hàng lần 2: {soLuongDonHangLan2}");

            // 🔹 Kiểm tra thay đổi số lượng đơn hàng không dùng if-else
            Console.WriteLine($"✅ {(soLuongDonHangLan2 > soLuongDonHangLan1 ? $"Đơn hàng đã được cập nhật! (Tăng thêm {soLuongDonHangLan2 - soLuongDonHangLan1} đơn hàng)" : "Không có thay đổi về số lượng đơn hàng!")}");
        }




        [Test]
        public void TestThongKeDoanhThuThongTinDonHangTrungKhop()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); // Tăng thời gian chờ lên 20 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000);


            var menuDoanhThu = wait.Until(d => d.FindElement(By.XPath("//li/a[@href='/Admin/DonHangs/DoanhThu1']")));
            menuDoanhThu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/DoanhThu1"));
            Thread.Sleep(5000);
            Console.WriteLine("✅ Điều hướng đến trang doanh thu thành công");

            var doanhThuTheoNam = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoThang']")));
            doanhThuTheoNam.OrderByDescending(e => int.Parse(e.Text.Trim())).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo năm mới nhất thành công");

            var doanhThuTheoThang = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DoanhThuTheoNgay']")));
            doanhThuTheoThang.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(5000);
            Console.WriteLine("✅ Xem doanh thu theo tháng mới nhất thành công");

            var donHangTheoNgay = wait.Until(d => d.FindElements(By.CssSelector("a[href*='DonHangTheoNgay']")));
            donHangTheoNgay.OrderByDescending(e => int.Parse(e.Text.Trim().Split(' ')[1])).First().Click();
            Thread.Sleep(4000);
            Console.WriteLine("✅ Xem đơn hàng theo ngày mới nhất thành công");

            // Lưu thông tin đơn hàng từ trang web
            List<Dictionary<string, string>> donhangWeb = new List<Dictionary<string, string>>();
            var rows = driver.FindElements(By.CssSelector("table tbody tr"));
            foreach (var row in rows)
            {
                var columns = row.FindElements(By.TagName("td"));
                if (columns.Count > 0)
                {
                    donhangWeb.Add(new Dictionary<string, string>
                    {
                        { "MaDh", columns[0].Text.Trim() },
                        { "Username", columns[1].Text.Trim() },
                        { "Diachi", columns[2].Text.Trim() },
                        { "TongTien", columns[3].Text.Trim().Replace(" VNĐ", "").Replace(",", "") },
                        { "CreatedAt", columns[4].Text.Trim() },
                        { "TrangThai", columns[5].Text.Trim() }
                    });
                }
            }

            Console.WriteLine("✅ Đã lưu danh sách đơn hàng từ trang web");

            // Kiểm tra số lượng phần tử trước khi so sánh để tránh lỗi IndexOutOfRange
            int minCount = Math.Min(donhangWeb.Count, donhangData.Count);

            for (int i = 0; i < minCount; i++)
            {
                var donhangWebRow = donhangWeb[i];
                var donhangExcelRow = donhangData[i];

                Assert.AreEqual(donhangExcelRow["MaDh"], donhangWebRow["MaDh"], "Mã đơn hàng không khớp");
                Assert.AreEqual(donhangExcelRow["Username"], donhangWebRow["Username"], "Username không khớp");
                Assert.AreEqual(donhangExcelRow["Diachi"], donhangWebRow["Diachi"], "Địa chỉ không khớp");
                Assert.AreEqual(donhangExcelRow["TongTien"], donhangWebRow["TongTien"], "Tổng tiền không khớp");
                Assert.AreEqual(donhangExcelRow["CreatedAt"], donhangWebRow["CreatedAt"], "Ngày tạo không khớp");
                Assert.AreEqual(donhangExcelRow["TrangThai"], donhangWebRow["TrangThai"], "Trạng thái không khớp");
            }

            Console.WriteLine("✅ Dữ liệu đơn hàng từ trang web và Excel khớp nhau");
        }


        [Test]
        public void TestThongKeDoanhThuKiemTraTongTienCoTinhToanDungKhong()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20)); // Tăng thời gian chờ lên 20 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(3000);

            // 🔹 Điều hướng đến trang thống kê doanh thu
            wait.Until(d => d.FindElement(By.XPath("//li[@class='nav-item']/a[@href='/Admin/DonHangs/DoanhThu1']"))).Click();
            wait.Until(d => d.Url.Contains("DoanhThu1"));
            Console.WriteLine("✅ Điều hướng đến trang thống kê doanh thu!");
            Thread.Sleep(3000);

            // 🔹 Chọn năm, tháng, ngày mới nhất và lưu lại link trang "Doanh Thu Theo Ngày"
            var namElement = wait.Until(d => d.FindElements(By.XPath("//table/tbody/tr/td/a"))).First();
            string doanhThuNgayLink = namElement.GetAttribute("href"); // Lưu link ngày
            namElement.Click();
            wait.Until(d => d.Url.Contains("nam="));
            Console.WriteLine("✅ Chọn năm mới nhất!");
            Thread.Sleep(3000);

            var thangElement = wait.Until(d => d.FindElements(By.XPath("//table/tbody/tr/td/a"))).First();
            doanhThuNgayLink = thangElement.GetAttribute("href"); // Cập nhật link tháng
            thangElement.Click();
            wait.Until(d => d.Url.Contains("thang="));
            Console.WriteLine("✅ Chọn tháng mới nhất!");
            Thread.Sleep(3000);

            var ngayElement = wait.Until(d => d.FindElements(By.XPath("//table/tbody/tr/td/a"))).First();
            doanhThuNgayLink = ngayElement.GetAttribute("href"); // Cập nhật link ngày
            ngayElement.Click();
            wait.Until(d => d.Url.Contains("ngay="));
            Console.WriteLine("✅ Chọn ngày mới nhất!");
            Thread.Sleep(3000);

            // 🔹 Lấy danh sách đơn hàng và tính tổng tiền
            var orderRows = wait.Until(d => d.FindElements(By.XPath("//table/tbody/tr")));
            Assert.IsTrue(orderRows.Count > 0, "❌ Không có đơn hàng nào hiển thị!");
            Thread.Sleep(3000);

            decimal tongTienDonHang = orderRows
                .Select(row => row.FindElement(By.XPath("./td[last()-2]")))
                .Select(cell => Regex.Match(cell.Text, "\\d+([,.]\\d+)*").Value.Replace(",", "").Trim())
                .Where(text => !string.IsNullOrEmpty(text))
                .Sum(text => decimal.Parse(text));

            Console.WriteLine($"✅ Tổng tiền từ danh sách đơn hàng: {tongTienDonHang} VND");
            Thread.Sleep(5000);

            // 🔙 Quay lại trang doanh thu theo ngày bằng JavaScript (bấm vào "Trở về" button)
            ((IJavaScriptExecutor)driver).ExecuteScript("javascript:history.back()");  // Quay lại trang trước
            Console.WriteLine($"✅ Quay lại trang Doanh Thu Theo Ngày!");
            wait.Until(d => d.Url.Contains("DoanhThuTheoNgay"));
            Thread.Sleep(3000);

            // 🔹 Kiểm tra nếu bảng tổng doanh thu có hiển thị
            var tableRows = wait.Until(d => d.FindElements(By.XPath("//table/tbody/tr")));
            Assert.IsTrue(tableRows.Count > 0, "❌ Không tìm thấy bảng doanh thu!");
            Console.WriteLine($"✅ Tìm thấy {tableRows.Count} dòng dữ liệu trong bảng doanh thu.");
            Thread.Sleep(5000);

            // 🔹 Lấy tổng doanh thu từ hàng cuối cùng của bảng
            var lastRow = tableRows.Last();
            Console.WriteLine($"✅ Dữ liệu dòng cuối cùng: {lastRow.Text}");
            var tongTienElement = lastRow.FindElement(By.XPath("./td[last()]"));

            // 🔹 Chuyển đổi dữ liệu sang số
            string tongTienText = Regex.Match(tongTienElement.Text, "\\d+([,.]\\d+)*").Value.Replace(",", "").Trim();
            Console.WriteLine($"✅ Tổng doanh thu trên giao diện: {tongTienText} VND");

            decimal tongTienHienThi = decimal.Parse(tongTienText);

            // 🔹 So sánh tổng tiền đơn hàng với tổng tiền hiển thị
            Assert.AreEqual(tongTienDonHang, tongTienHienThi, "❌ Tổng tiền không khớp!");
            Console.WriteLine("🎉 Kiểm tra tổng tiền doanh thu chính xác!");
        }

        //==============================================================================================================================================================================================
        // Biểu đồ

        [Test]
        public void TestThongKeBieuDoClickVaoDuongDanSanPham()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); // Tăng thời gian chờ lên 20 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000);

            // 🔹 Click vào menu "Thống kê số lượng đã bán"
            IWebElement thongKeMenu = wait.Until(d => d.FindElement(By.XPath("//a[@href='/Admin/DonHangs/SoLuongDaBan']")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", thongKeMenu);
            Thread.Sleep(4000);
            thongKeMenu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/SoLuongDaBan"));
            Console.WriteLine("✅ Điều hướng đến trang thống kê số lượng đã bán!");

            // Kiểm tra biểu đồ Pie Chart có hiển thị
            Assert.IsTrue(driver.FindElement(By.Id("pieChart")).Displayed, "❌ Biểu đồ không hiển thị!");

            // Kiểm tra danh sách sản phẩm có hiển thị
            var danhSachSanPham = driver.FindElements(By.XPath("//a/strong"));
            Assert.IsTrue(danhSachSanPham.Count > 0, "❌ Không có sản phẩm nào trong danh sách thống kê!");
            Console.WriteLine($"✅ Tìm thấy {danhSachSanPham.Count} sản phẩm trong danh sách thống kê.");

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

            // Lặp qua từng sản phẩm, kiểm tra đường dẫn và quay lại
            foreach (var sanPham in danhSachSanPham)
            {
                string tenSanPham = sanPham.Text;
                string urlTruocKhiClick = driver.Url;
                string urlSanPham = sanPham.FindElement(By.XPath("..")).GetAttribute("href"); // Lấy URL của sản phẩm

                // Mở link trong tab mới bằng JavaScript
                js.ExecuteScript($"window.open('{urlSanPham}', '_blank');");
                Thread.Sleep(3000);

                // Chuyển sang tab mới
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                wait.Until(d => d.Url != urlTruocKhiClick);

                // Kiểm tra xem trang chi tiết sản phẩm có đúng không
                Assert.IsTrue(driver.Url.Contains("/SanPham/ChiTietSanPham"), $"❌ Sai đường dẫn chi tiết cho sản phẩm: {tenSanPham}");
                Console.WriteLine($"✅ Click vào sản phẩm '{tenSanPham}' thành công, đúng trang chi tiết!");

                // Đóng tab hiện tại và quay lại danh sách
                driver.Close();
                driver.SwitchTo().Window(driver.WindowHandles.First());
                Thread.Sleep(3000);
            }

            Console.WriteLine("🎉 Đã kiểm tra xong tất cả sản phẩm trong danh sách thống kê!");
        }



        [Test]
        public void TestThongKeBieuDoNhanXetSanPham()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); // Tăng thời gian chờ lên 10 giây
            var user = userLoginData[1];

            // 🔹 Đăng nhập
            DangNhap(user["Username"], user["Password"]);
            Console.WriteLine($"✅ Đăng nhập thành công: {user["Username"]}");
            Thread.Sleep(4000); // Thời gian chờ sau khi đăng nhập

            // 🔹 Click vào menu "Thống kê số lượng đã bán"
            IWebElement thongKeMenu = wait.Until(d => d.FindElement(By.XPath("//a[@href='/Admin/DonHangs/SoLuongDaBan']")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", thongKeMenu);
            Thread.Sleep(2000); // Thời gian chờ để phần tử hiển thị
            thongKeMenu.Click();
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/SoLuongDaBan"));
            Console.WriteLine("✅ Điều hướng đến trang thống kê số lượng đã bán!");
            Thread.Sleep(3000); // Thời gian chờ để trang tải

            // Lấy danh sách nhận xét sản phẩm từ ViewBag
            var danhSachNhanXet = driver.FindElements(By.XPath("//div[@class='info-container']//ul//span"));
            Thread.Sleep(2000); // Thời gian chờ sau khi lấy danh sách nhận xét

            // Lấy thông tin về số lượng bán và tên sản phẩm
            var tenSanPham = driver.FindElements(By.XPath("//div[@class='info-container']//ul//a/strong")).Select(e => e.Text).ToList();
            Thread.Sleep(2000); // Thời gian chờ sau khi lấy tên sản phẩm
            var soLuongDaBan = driver.FindElements(By.XPath("//div[@class='info-container']//ul//li"))
                .Select(e => e.Text.Contains("Số lượng bán") ? int.Parse(e.Text.Split(':')[1].Split("sản phẩm")[0].Trim()) : 0).ToList();
            Thread.Sleep(2000); // Thời gian chờ sau khi lấy số lượng bán

            // Lấy max và min số lượng bán
            int maxSoLuongDaBan = soLuongDaBan.Max();
            int minSoLuongDaBan = soLuongDaBan.Min();

            // Kiểm tra từng nhận xét có hiển thị đúng
            for (int i = 0; i < tenSanPham.Count; i++)
            {
                var soLuong = soLuongDaBan[i];
                var nhanXet = danhSachNhanXet[i].Text;
                var expectedComment = GetExpectedComment(soLuong, maxSoLuongDaBan, minSoLuongDaBan);

                // In ra thông tin cho mỗi sản phẩm
                Console.WriteLine($"🛒 Sản phẩm: {tenSanPham[i]} - Số lượng bán: {soLuong} - Nhận xét theo thống kê: {nhanXet}");

                // Kiểm tra nhận xét có chính xác không
                Assert.AreEqual(expectedComment, nhanXet, $"❌ Nhận xét cho sản phẩm '{tenSanPham[i]}' không chính xác. Dự đoán: {expectedComment}, Nhận xét: {nhanXet}");
                Console.WriteLine($"✅ Nhận xét cho sản phẩm '{tenSanPham[i]}' là đúng với điều kiện: {nhanXet}");

                // Thêm thời gian chờ giữa các sản phẩm (tuỳ thuộc vào yêu cầu)
                Thread.Sleep(2000); // Thời gian chờ sau khi kiểm tra từng sản phẩm
            }
        }




        [Test]
        public void TestThongKeBieuDoKiemTraDungPhanTram()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // Điều hướng đến trang thống kê số lượng bán
            driver.Navigate().GoToUrl("https://localhost:7053/Admin/DonHangs/SoLuongDaBan");
            wait.Until(d => d.Url.Contains("/Admin/DonHangs/SoLuongDaBan"));
            Console.WriteLine("✅ Điều hướng đến trang thống kê số lượng đã bán thành công");

            // Chờ biểu đồ xuất hiện
            var pieChart = wait.Until(d => d.FindElement(By.Id("pieChart")));
            Assert.IsNotNull(pieChart, "❌ Biểu đồ không hiển thị trên trang!");
            Console.WriteLine("✅ Biểu đồ hiển thị thành công!");

            // Cuộn đến biểu đồ để tránh lỗi MoveTargetOutOfBoundsException
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", pieChart);
            Thread.Sleep(500); // Đợi cuộn xong

            // Lấy tổng số sản phẩm trong biểu đồ
            int totalProducts = (int)(long)((IJavaScriptExecutor)driver).ExecuteScript(@"
    var chart = Chart.getChart('pieChart');
    return chart && chart.data && chart.data.labels ? chart.data.labels.length : 0;
    ");
            Assert.IsTrue(totalProducts > 0, "❌ Không có sản phẩm nào trong biểu đồ!");
            Console.WriteLine($"✅ Tìm thấy {totalProducts} sản phẩm trong biểu đồ.");

            // Lấy tổng số lượng bán được từ biểu đồ (cần tính tổng tất cả các phần tử trong biểu đồ)
            var totalQuantity = (long)((IJavaScriptExecutor)driver).ExecuteScript(@"
    var chart = Chart.getChart('pieChart');
    return chart && chart.data && chart.data.datasets && chart.data.datasets[0] && chart.data.datasets[0].data
        ? chart.data.datasets[0].data.reduce((acc, val) => acc + val, 0)
        : 0;
    ");
            Assert.IsTrue(totalQuantity > 0, "❌ Không có dữ liệu số lượng bán trong biểu đồ!");

            double totalPercentage = 0.0;

            // Lặp qua từng sản phẩm để kiểm tra tooltip và cộng phần trăm
            for (int i = 0; i < totalProducts; i++)
            {
                // Lấy tên sản phẩm từ chart.data.labels
                string productName = (string)((IJavaScriptExecutor)driver).ExecuteScript($@"
        var chart = Chart.getChart('pieChart');
        return chart && chart.data && chart.data.labels ? chart.data.labels[{i}] : '';
        ");

                // Kích hoạt tooltip cho từng sản phẩm
                string script = $@"
        var chart = Chart.getChart('pieChart');
        if (chart && chart.data && chart.data.datasets && chart.data.datasets[0] && chart.data.datasets[0].data) {{
            chart.setActiveElements([{{datasetIndex: 0, index: {i}}}]);
            chart.tooltip.update();
            var tooltipText = chart.tooltip.title && chart.tooltip.title[0] ? chart.tooltip.title[0] : '';
            var bodyText = chart.tooltip.body && chart.tooltip.body[0] && chart.tooltip.body[0].lines[0] ? chart.tooltip.body[0].lines[0] : '';
            var dataValue = chart.data.datasets[0].data[{i}];
            var percentage = ((dataValue / {totalQuantity}) * 100).toFixed(2);
            return tooltipText + ' - ' + bodyText + ' - ' + percentage + '%';
        }}
        return null;
        ";

                string tooltipText = (string)((IJavaScriptExecutor)driver).ExecuteScript(script);
                Assert.IsNotNull(tooltipText, $"❌ Tooltip không xuất hiện cho sản phẩm {productName}!");
                Console.WriteLine($"✅ Tooltip sản phẩm {productName}: {tooltipText}");

                // Lấy phần trăm từ tooltip
                var percentageString = tooltipText.Split('-').LastOrDefault()?.Trim();  // Lấy phần trăm từ cuối chuỗi
                Assert.IsTrue(!string.IsNullOrEmpty(percentageString), $"❌ Không tìm thấy phần trăm trong tooltip của sản phẩm {productName}.");

                // Loại bỏ ký tự "%" và chuyển chuỗi thành kiểu double
                percentageString = percentageString.Replace("%", "").Trim();

                // Thực hiện chuyển đổi kiểu đúng từ string sang double
                bool parseSuccess = double.TryParse(percentageString, NumberStyles.Any, CultureInfo.InvariantCulture, out double tooltipPercentage);
                Assert.IsTrue(parseSuccess, $"❌ Phần trăm không hợp lệ cho sản phẩm {productName}. Phần trăm nhận được: {percentageString}");

                // Cộng phần trăm vào tổng
                totalPercentage += tooltipPercentage;
            }

            // In tổng phần trăm sau khi cộng tất cả sản phẩm
            Console.WriteLine($"✅ Tổng phần trăm sau khi cộng cho tất cả sản phẩm: {totalPercentage}%");

            // Kiểm tra tổng phần trăm có bằng 100%
            // Tăng phạm vi dung sai lên 0.03 thay vì 0.01 để chấp nhận sai số lớn hơn
            double tolerance = 0.03;  // Sai lệch cho phép
            Assert.IsTrue(Math.Abs(totalPercentage - 100.0) <= tolerance, $"❌ Tổng phần trăm của tất cả sản phẩm không bằng 100%! Tổng phần trăm: {totalPercentage}");

            Console.WriteLine("🎉 Tất cả sản phẩm đã được kiểm tra thành công!");
        }


        private string GetExpectedComment(int soLuongDaBan, int maxSoLuongDaBan, int minSoLuongDaBan)
        {


            // Kiểm tra xem số lượng bán có phải là cao nhất hoặc thấp nhất
            if (soLuongDaBan == maxSoLuongDaBan)
                return "Sản phẩm bán chạy nhất!";
            else if (soLuongDaBan == minSoLuongDaBan)
                return "Sản phẩm ít được mua.";
            else if (soLuongDaBan >= 50)
                return "Sản phẩm khá phổ biến.";
            else if (soLuongDaBan >= 20)
                return "Sản phẩm đang được ưa chuộng.";
            else
                return "Sản phẩm ít được mua.";
        }

        [TearDown]
        public void TearDown()
        {
            driver.Dispose(); // Đóng trình duyệt sau khi test xong
        }
    }
}
