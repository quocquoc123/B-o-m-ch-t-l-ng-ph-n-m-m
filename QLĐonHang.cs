using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Interactions;
using System;


namespace TestProject

{
    public class Tests
    {

        private ChromeDriver driver;
        private string baseUrl = "https://localhost:7053/";
        private List<Dictionary<string, string>> loginData;
        private List<Dictionary<string, string>> loginData1;
        private List<Dictionary<string, string>> ThanhToan;
    
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(baseUrl);

            // Đọc dữ liệu từ file Excel
            string filePath = @"D:\TestData1.xlsx";
            loginData = ReadExcel(filePath, "Login");
            loginData1 = ReadExcel(filePath, "Login1");
            ThanhToan = ReadExcel(filePath, "ThanhToan");
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

       

        [Test]
        public void TestDangNhapAdmin()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }
        }





        [Test]
        public void TestDangNhapAcout()
        {
            foreach (var data in loginData1)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }
        }





        [Test]
        public void Test_CapNhapTrangThaiXacNhan()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            try
            {
                IWebElement xacNhanButton = driver.FindElement(By.XPath("//button[contains(text(),'Xác nhận đơn')]"));
                if (xacNhanButton.Displayed && xacNhanButton.Enabled)
                {
                    Console.WriteLine("Có nút Xác nhận đơn");
                    Thread.Sleep(2000);

                    IWebElement parentRow = xacNhanButton.FindElement(By.XPath("./ancestor::tr"));

                    // Lấy ID đơn hàng để tìm lại sau khi DOM thay đổi
                    string donHangID = parentRow.FindElement(By.XPath(".//td[1]")).Text; // Điều chỉnh cột chứa ID đơn hàng

                    // Lấy trạng thái hiện tại trước khi xác nhận
                    IWebElement trangThaiElement = parentRow.FindElement(By.XPath(".//td[5]")); // Điều chỉnh nếu cần
                    string trangThaiTruoc = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái trước khi xác nhận: " + trangThaiTruoc);

                    // Nhấn vào nút xác nhận
                    xacNhanButton.Click();
                    Thread.Sleep(3000);


                    // Chờ hệ thống cập nhật trạng thái đơn hàng
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(driver =>
                    {
                        try
                        {
                            // Tìm lại hàng chứa đơn hàng sau khi DOM thay đổi
                            IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));

                            // Lấy trạng thái mới
                            IWebElement trangThaiMoiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                            return trangThaiMoiElement.Text != trangThaiTruoc; // Chờ đến khi trạng thái thay đổi
                        }
                        catch (NoSuchElementException)
                        {
                            return false; // Nếu chưa tìm thấy thì tiếp tục đợi
                        }
                    });

                    // Sau khi chờ trạng thái thay đổi, tìm lại dòng đơn hàng
                    IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));
                    trangThaiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                    string trangThaiMoi = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái sau khi xác nhận: " + trangThaiMoi);

                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Không tìm thấy nút Xác nhận đơn");
            }


        }

        [Test]
        public void Test_CapNhapTrangThaiDaGiao()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            try
            {
                // Tìm nút "Đã giao"
                IWebElement daGiaoButton = driver.FindElement(By.XPath("//button[contains(text(),'Đã giao')]"));

                if (daGiaoButton.Displayed && daGiaoButton.Enabled)
                {
                    Console.WriteLine("Có nút Đã giao");
                    Thread.Sleep(2000);

                    // Tìm phần tử cha (hàng chứa nút này)
                    IWebElement parentRow = daGiaoButton.FindElement(By.XPath("./ancestor::tr"));

                    // Lấy ID đơn hàng để tìm lại sau khi DOM thay đổi
                    string donHangID = parentRow.FindElement(By.XPath(".//td[1]")).Text; // Điều chỉnh cột chứa ID đơn hàng

                    // Lấy trạng thái hiện tại trước khi bấm nút
                    IWebElement trangThaiElement = parentRow.FindElement(By.XPath(".//td[5]")); // Điều chỉnh nếu cần
                    string trangThaiTruoc = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái trước khi bấm 'Đã giao': " + trangThaiTruoc);

                    // Nhấn vào nút "Đã giao"
                    daGiaoButton.Click();
                    Thread.Sleep(3000); // Đợi hệ thống cập nhật trạng thái

                    // Chờ hệ thống cập nhật trạng thái đơn hàng
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(driver =>
                    {
                        try
                        {
                            // Tìm lại hàng chứa đơn hàng sau khi DOM thay đổi
                            IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));

                            // Lấy trạng thái mới
                            IWebElement trangThaiMoiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                            return trangThaiMoiElement.Text != trangThaiTruoc; // Chờ đến khi trạng thái thay đổi
                        }
                        catch (NoSuchElementException)
                        {
                            return false; // Nếu chưa tìm thấy thì tiếp tục đợi
                        }
                    });

                    // Sau khi chờ trạng thái thay đổi, tìm lại dòng đơn hàng
                    IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));
                    trangThaiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                    string trangThaiMoi = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái sau khi bấm 'Đã giao': " + trangThaiMoi);
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Không tìm thấy nút Đã giao");
            }
        }


        [Test]
        public void Test_CapNhapTrangThaiHuyDonHang()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            try
            {
                // Tìm nút "Hủy đơn hàng"
                IWebElement huyDonButton = driver.FindElement(By.XPath("//button[contains(text(),'Hủy đơn hàng')]"));

                if (huyDonButton.Displayed && huyDonButton.Enabled)
                {
                    Console.WriteLine("Có nút Hủy đơn hàng");
                    Thread.Sleep(2000);

                    // Tìm phần tử cha (hàng chứa nút này)
                    IWebElement parentRow = huyDonButton.FindElement(By.XPath("./ancestor::tr"));

                    // Lấy ID đơn hàng để tìm lại sau khi DOM thay đổi
                    string donHangID = parentRow.FindElement(By.XPath(".//td[1]")).Text; // Điều chỉnh cột chứa ID đơn hàng

                    // Lấy trạng thái hiện tại trước khi bấm nút
                    IWebElement trangThaiElement = parentRow.FindElement(By.XPath(".//td[5]")); // Điều chỉnh nếu cần
                    string trangThaiTruoc = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái trước khi bấm 'Hủy đơn hàng': " + trangThaiTruoc);

                    // Nhấn vào nút "Hủy đơn hàng"
                    huyDonButton.Click();
                    Thread.Sleep(3000); // Đợi hệ thống cập nhật trạng thái

                    // Chờ hệ thống cập nhật trạng thái đơn hàng
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(driver =>
                    {
                        try
                        {
                            // Tìm lại hàng chứa đơn hàng sau khi DOM thay đổi
                            IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));

                            // Lấy trạng thái mới
                            IWebElement trangThaiMoiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                            return trangThaiMoiElement.Text != trangThaiTruoc; // Chờ đến khi trạng thái thay đổi
                        }
                        catch (NoSuchElementException)
                        {
                            return false; // Nếu chưa tìm thấy thì tiếp tục đợi
                        }
                    });

                    // Sau khi chờ trạng thái thay đổi, tìm lại dòng đơn hàng
                    IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));
                    trangThaiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                    string trangThaiMoi = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái sau khi bấm 'Hủy đơn hàng': " + trangThaiMoi);
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Không tìm thấy nút Hủy đơn hàng");
            }

        }


        [Test]
        public void Test_ChapNhanHuyDonHang()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            try
            {
                // Tìm nút "Hủy đơn hàng"
                IWebElement huyDonButton = driver.FindElement(By.XPath("//button[contains(text(),'Chấp nhận hủy đơn hàng')]"));

                if (huyDonButton.Displayed && huyDonButton.Enabled)
                {
                    Console.WriteLine("Có nút Chấp nhận hủy đơn hàng");
                    Thread.Sleep(2000);

                    // Tìm phần tử cha (hàng chứa nút này)
                    IWebElement parentRow = huyDonButton.FindElement(By.XPath("./ancestor::tr"));

                    // Lấy ID đơn hàng để tìm lại sau khi DOM thay đổi
                    string donHangID = parentRow.FindElement(By.XPath(".//td[1]")).Text; // Điều chỉnh cột chứa ID đơn hàng

                    // Lấy trạng thái hiện tại trước khi bấm nút
                    IWebElement trangThaiElement = parentRow.FindElement(By.XPath(".//td[5]")); // Điều chỉnh nếu cần
                    string trangThaiTruoc = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái trước khi bấm 'Chấp nhận hủy đơn hàng': " + trangThaiTruoc);

                    // Nhấn vào nút "Hủy đơn hàng"
                    huyDonButton.Click();
                    Thread.Sleep(3000); // Đợi hệ thống cập nhật trạng thái

                    // Chờ hệ thống cập nhật trạng thái đơn hàng
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(driver =>
                    {
                        try
                        {
                            // Tìm lại hàng chứa đơn hàng sau khi DOM thay đổi
                            IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));

                            // Lấy trạng thái mới
                            IWebElement trangThaiMoiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                            return trangThaiMoiElement.Text != trangThaiTruoc; // Chờ đến khi trạng thái thay đổi
                        }
                        catch (NoSuchElementException)
                        {
                            return false; // Nếu chưa tìm thấy thì tiếp tục đợi
                        }
                    });

                    // Sau khi chờ trạng thái thay đổi, tìm lại dòng đơn hàng
                    IWebElement updatedRow = driver.FindElement(By.XPath($"//tr[td[1][contains(text(), '{donHangID}')]]"));
                    trangThaiElement = updatedRow.FindElement(By.XPath(".//td[5]"));
                    string trangThaiMoi = trangThaiElement.Text;
                    Console.WriteLine("Trạng thái sau khi bấm 'Chấp nhận hủy đơn hàng': " + trangThaiMoi);
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Không tìm thấy nút Chấp nhận hủy đơn hàng");
            }

        }


        [Test]
        public void Test_PhanTrangItHon5Don()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            IReadOnlyCollection<IWebElement> danhSachDonHang = driver.FindElements(By.XPath("//table//tbody//tr"));
            int soLuongDonHang = danhSachDonHang.Count;

            Console.WriteLine("Số lượng đơn hàng hiển thị: " + soLuongDonHang + " Đơn hàng");
            Thread.Sleep(3000);
        }

        [Test]
        public void Test_PhanTrang5Don()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            IReadOnlyCollection<IWebElement> danhSachDonHang = driver.FindElements(By.XPath("//table//tbody//tr"));
            int soLuongDonHang = danhSachDonHang.Count;

            Console.WriteLine("Số lượng đơn hàng hiển thị: " + soLuongDonHang + " Đơn hàng");
            Thread.Sleep(3000);
        }


        [Test]
        public void Test_PhanTrangHon5DonHang()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            // Biến đếm số trang
            int soTrang = 1;

            while (true)
            {
                // Đếm số lượng đơn hàng trên trang hiện tại
                IReadOnlyCollection<IWebElement> danhSachDonHang = driver.FindElements(By.XPath("//table//tbody//tr"));
                int soLuongDonHang = danhSachDonHang.Count;
                Console.WriteLine("Trang " + soTrang + " có: " + soLuongDonHang + " đơn hàng.");
                Thread.Sleep(2000);

                try
                {
                    // Tìm nút "Next"
                    IWebElement nextButton = driver.FindElement(By.XPath("//a[contains(text(),'Next')]"));
                    if (nextButton.Displayed && nextButton.Enabled)
                    {
                        nextButton.Click(); // Nhấn để chuyển trang
                        Thread.Sleep(3000); // Chờ trang tải lại
                        soTrang++; // Tăng số trang lên
                    }
                    else
                    {
                        Console.WriteLine("Không còn trang nào nữa.");
                        break; // Thoát khỏi vòng lặp nếu không còn trang tiếp theo
                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine("Không tìm thấy nút Next, đã đến trang cuối cùng.");
                    break; // Thoát khỏi vòng lặp
                }
            }

        }


        [Test]
        public void Test_XuatHoaDonDaGiao()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            try
            {
                // Kiểm tra có đơn hàng nào có trạng thái "Đã giao" không
                IWebElement daGiaoRow = driver.FindElement(By.XPath("//tr[td[contains(text(),'Đã giao')]]"));

                if (daGiaoRow != null)
                {
                    Console.WriteLine(" Có đơn hàng với trạng thái 'Đã giao'.");

                    // Tìm nút "Xuất hóa đơn PDF" trong hàng đơn hàng này
                    IWebElement xuatHoaDonButton = daGiaoRow.FindElement(By.XPath(".//button[contains(text(),'Xuất Hóa Đơn PDF')]"));


                    if (xuatHoaDonButton.Displayed && xuatHoaDonButton.Enabled)
                    {
                        Console.WriteLine(" Nhấn vào nút 'Xuất Hóa Đơn PDF'...");
                        xuatHoaDonButton.Click();
                        Thread.Sleep(5000); // Chờ hộp thoại xuất hiện

                        // Kiểm tra xem hộp thoại lưu tệp có xuất hiện không
                        if (KiemTraHopThoaiLuuTep())
                        {
                            Console.WriteLine(" Xuất hóa đơn thành công!");
                        }
                        else
                        {
                            Console.WriteLine("Không thấy hộp thoại lưu tệp, có thể lỗi khi xuất hóa đơn.");
                        }
                    }
                    else
                    {
                        Console.WriteLine(" Không tìm thấy nút 'Xuất Hóa Đơn PDF'.");
                    }
                }
                else
                {
                    Console.WriteLine("Không có đơn hàng nào với trạng thái 'Đã giao'.");
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine(" Không tìm thấy đơn hàng nào với trạng thái 'Đã giao' hoặc nút 'Xuất Hóa Đơn PDF'.");
            }


        }



        [Test]
        public void Test_XuatHoaDonChuaXacNhan()
        {
            IWebElement IconDN = driver.FindElement(By.CssSelector(".bx-user"));
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(3000);

            try
            {
                // Kiểm tra có đơn hàng nào có trạng thái "Đã giao" không
                IWebElement daGiaoRow = driver.FindElement(By.XPath("//tr[td[contains(text(),'Dang Cho')]]"));

                if (daGiaoRow != null)
                {
                    Console.WriteLine(" Có đơn hàng với trạng thái 'Đang chờ'.");

                    // Tìm nút "Xuất hóa đơn PDF" trong hàng đơn hàng này
                    IWebElement xuatHoaDonButton = daGiaoRow.FindElement(By.XPath(".//button[contains(text(),'Xuất Hóa Đơn PDF')]"));


                    if (xuatHoaDonButton.Displayed && xuatHoaDonButton.Enabled)
                    {
                        Console.WriteLine(" Nhấn vào nút 'Xuất Hóa Đơn PDF'...");
                        xuatHoaDonButton.Click();
                        Thread.Sleep(5000); // Chờ hộp thoại xuất hiện

                        // Kiểm tra xem hộp thoại lưu tệp có xuất hiện không
                        if (KiemTraHopThoaiLuuTep())
                        {
                            Console.WriteLine("Thông báo không được xuất file trạng thái đang chờ");
                        }
                        else
                        {
                            Console.WriteLine("Thông báo chạy đúng ");
                        }
                    }
                    else
                    {
                        Console.WriteLine(" Không tìm thấy nút 'Xuất Hóa Đơn PDF'.");
                    }
                }
                else
                {
                    Console.WriteLine("Không có đơn hàng nào với trạng thái 'Đang chờ'.");
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine(" Không tìm thấy đơn hàng nào với trạng thái 'Đang chờ' hoặc nút 'Xuất Hóa Đơn PDF'.");
            }


        }


        private bool KiemTraHopThoaiLuuTep()
        {
            try
            {
                // Chờ 3 giây xem có hộp thoại lưu xuất hiện không
                Thread.Sleep(3000);

                // Cách kiểm tra file có tải về hay không (Chỉ hoạt động nếu có quyền truy cập thư mục)
                string downloadPath = "C:\\Users\\YourUsername\\Downloads"; // Cập nhật đúng đường dẫn thư mục tải xuống của bạn
                string[] files = Directory.GetFiles(downloadPath, "*.pdf");

                return files.Length > 0; // Nếu có file PDF trong thư mục => Xuất hóa đơn thành công
            }
            catch (Exception)
            {
                return false;
            }
        }


        [Test]
        public void Test_MuaHangKTDonHang()
        {
            foreach (var data in loginData1)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }

            IWebElement duocpham = driver.FindElement(By.LinkText("DƯỢC PHẨM"));
            duocpham.Click();
            Thread.Sleep(2000);

            IWebElement chitietsp = driver.FindElement(By.ClassName("product-card"));
            chitietsp.Click();
            Thread.Sleep(2000);

            IWebElement themvaogiohang = driver.FindElement(By.LinkText("Thêm vào giỏ hàng"));
            themvaogiohang.Click();
            Thread.Sleep(2000);

            IWebElement thanhtoan = driver.FindElement(By.LinkText("Thanh toán"));
            thanhtoan.Click();
            Thread.Sleep(2000);


            foreach (var data in ThanhToan)
            {
                IWebElement MaGiamGia = driver.FindElement(By.Name("maKhuyenMai"));
                MaGiamGia.SendKeys(data["maKhuyenMai"]);
                Thread.Sleep(2000);

                IWebElement apMGG = driver.FindElement(By.XPath("//button[contains(text(), 'Áp dụng')]"));
                apMGG.Click();
                Thread.Sleep(2000);

                IWebElement ChonTT = driver.FindElement(By.Id("payment-method"));
                ChonTT.Click();
                Thread.Sleep(2000);

                IWebElement optionCOD = driver.FindElement(By.XPath("//option[@value='cod']"));
                optionCOD.Click();
                Thread.Sleep(2000);

                IWebElement Ten = driver.FindElement(By.Name("tennguoinhan"));
                Ten.SendKeys(data["tennguoinhan"]);
                Thread.Sleep(2000);

                IWebElement sdt = driver.FindElement(By.Name("sdtnguoinhan"));
                sdt.SendKeys(data["sdtnguoinhan"]);
                Thread.Sleep(2000);

                IWebElement Tinh = driver.FindElement(By.Id("province"));
                Tinh.Click();
                Thread.Sleep(2000);

                IWebElement LamDong = driver.FindElement(By.XPath("//option[@value='68']"));
                LamDong.Click();

                IWebElement Quan = driver.FindElement(By.Id("district"));
                Quan.Click();
                Thread.Sleep(2000);

                IWebElement BaoLoc = driver.FindElement(By.XPath("//option[@value='673']"));
                BaoLoc.Click();

                IWebElement Xa = driver.FindElement(By.Id("ward"));
                Xa.Click();
                Thread.Sleep(2000);

                IWebElement Dl = driver.FindElement(By.XPath("//option[@value='Xã Đại Lào']"));
                Dl.Click();

                IWebElement DiaChi = driver.FindElement(By.Id("address_detail"));
                DiaChi.SendKeys(data["address_detail"]);
                Thread.Sleep(2000);
            }

            IWebElement ThanhToan1 = driver.FindElement(By.Id("cod-button"));
            ThanhToan1.Click();
            Thread.Sleep(3000);

            IWebElement userCheckIcon = driver.FindElement(By.CssSelector("i.bx-user-check"));
            userCheckIcon.Click();
            Thread.Sleep(2000);

            Actions action = new Actions(driver);
            IWebElement dang = driver.FindElement(By.CssSelector("a[href='/home/donhang']"));
            action.MoveToElement(dang).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị



            IWebElement dangxuat = driver.FindElement(By.XPath("//a[contains(text(),'Đăng xuất')]"));
            dangxuat.Click();
            Thread.Sleep(2000);

            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }
            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(2000);

            foreach (var data in ThanhToan)
            {
                IWebElement firstOrderRow = driver.FindElement(By.XPath("//table/tbody/tr[1]"));
                IWebElement hoTenElement = firstOrderRow.FindElement(By.XPath("./td[8]"));
                string actualHoTen = hoTenElement.Text.Trim();
                Assert.AreEqual(data["tennguoinhan"], actualHoTen, "Họ tên của đơn hàng đầu tiên không khớp!");

                IWebElement sdtElement = firstOrderRow.FindElement(By.XPath("./td[9]"));
                string actualSdt = sdtElement.Text.Trim();
                Assert.AreEqual(data["sdtnguoinhan"], actualSdt, "SDT của đơn hàng đầu tiên không khớp!");


                IWebElement DCElement = firstOrderRow.FindElement(By.XPath("./td[2]"));
                string actualDC = DCElement.Text.Trim();
                string expectedDC = data["address_detail"] + ", Xã Đại Lào, Thành phố Bảo Lộc, Tỉnh Lâm Đồng";
                Assert.AreEqual(expectedDC, actualDC, "Địa chỉ của đơn hàng đầu tiên không khớp!");

            }

        }


        [Test]
        public void Test_NguoiDungHuyDon()
        {
            foreach (var data in loginData1)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }



            Actions action = new Actions(driver);
            IWebElement dang = driver.FindElement(By.CssSelector("a[href='/home/donhang']"));
            action.MoveToElement(dang).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement donHangButton = driver.FindElement(By.XPath("//a[contains(text(),'Đơn hàng của tôi')]"));
            donHangButton.Click();
            Thread.Sleep(2000);

            IWebElement firstCancelButton = driver.FindElement(By.XPath("(//a[contains(@class, 'btnHuyDonHang')])[1]"));
            firstCancelButton.Click();
            Thread.Sleep(2000);

            IAlert alert = driver.SwitchTo().Alert(); // Chuyển hướng đến hộp thoại
            alert.Accept(); // Nhấn "OK"
            Thread.Sleep(2000);

            Actions action1 = new Actions(driver);
            IWebElement dang1 = driver.FindElement(By.CssSelector("a[href='/home/donhang']"));
            action1.MoveToElement(dang1).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị


            IWebElement dangxuat = driver.FindElement(By.XPath("//a[contains(text(),'Đăng xuất')]"));
            dangxuat.Click();
            Thread.Sleep(2000);

            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }

            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(2000);

            IWebElement firstOrderRow = driver.FindElement(By.XPath("//table/tbody/tr[1]")); // Chọn hàng đầu tiên
            IWebElement trangThaiElement = firstOrderRow.FindElement(By.XPath("./td[5]")); // Chọn cột 5
            string actualTrangThai = trangThaiElement.Text.Trim(); // Lấy nội dung và loại bỏ khoảng trắng
            Assert.AreEqual("Chờ xác nhận hủy đơn", actualTrangThai, "Trạng thái không khớp với mong đợi!");

        }


        [Test]
        public void Test_AdminCapNhatTrangThaiXacNhan()
        {


          
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }
            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(2000);

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement firstConfirmButton = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'btn-success11')])[1]")));
            wait.Until(d => firstConfirmButton.Displayed && firstConfirmButton.Enabled);
            firstConfirmButton.Click();
            Thread.Sleep(2000);


            Actions action1 = new Actions(driver);
            IWebElement dang1 = driver.FindElement(By.ClassName("icon-user-fullname"));
            action1.MoveToElement(dang1).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement dangxuat = driver.FindElement(By.XPath("//a[@href='/user/DangXuat']"));
            dangxuat.Click();
            Thread.Sleep(2000);

            foreach (var data in loginData1)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }

            Actions action = new Actions(driver);
            IWebElement dang = driver.FindElement(By.CssSelector("a[href='/home/donhang']"));
            action.MoveToElement(dang).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement donHangButton = driver.FindElement(By.XPath("//a[contains(text(),'Đơn hàng của tôi')]"));
            donHangButton.Click();
            Thread.Sleep(2000);

            IWebElement firstOrderRow = driver.FindElement(By.XPath("//table/tbody/tr[1]")); // Chọn hàng đầu tiên
            IWebElement trangThaiElement = firstOrderRow.FindElement(By.XPath("./td[3]")); // Chọn cột 5
            string actualTrangThai = trangThaiElement.Text.Trim(); // Lấy nội dung và loại bỏ khoảng trắng
            Assert.AreEqual("Đã xác nhận đơn hàng sẽ sớm được giao đến bạn", actualTrangThai, "Trạng thái không khớp với mong đợi!");

        }


        [Test]
        public void Test_AdminCapNhatTrangThaiDaGiao()
        {



            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }
            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(2000);

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement firstConfirmButton = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'btn btn-primary')])[1]")));
            wait.Until(d => firstConfirmButton.Displayed && firstConfirmButton.Enabled);
            firstConfirmButton.Click();
            Thread.Sleep(2000);


            Actions action1 = new Actions(driver);
            IWebElement dang1 = driver.FindElement(By.ClassName("icon-user-fullname"));
            action1.MoveToElement(dang1).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement dangxuat = driver.FindElement(By.XPath("//a[@href='/user/DangXuat']"));
            dangxuat.Click();
            Thread.Sleep(2000);

            foreach (var data in loginData1)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }

            Actions action = new Actions(driver);
            IWebElement dang = driver.FindElement(By.CssSelector("a[href='/home/donhang']"));
            action.MoveToElement(dang).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement donHangButton = driver.FindElement(By.XPath("//a[contains(text(),'Đơn hàng của tôi')]"));
            donHangButton.Click();
            Thread.Sleep(2000);

            IWebElement firstOrderRow = driver.FindElement(By.XPath("//table/tbody/tr[1]")); // Chọn hàng đầu tiên
            IWebElement trangThaiElement = firstOrderRow.FindElement(By.XPath("./td[3]")); // Chọn cột 5
            string actualTrangThai = trangThaiElement.Text.Trim(); // Lấy nội dung và loại bỏ khoảng trắng
            Assert.AreEqual("Đã giao", actualTrangThai, "Trạng thái không khớp với mong đợi!");

        }



        [Test]
        public void Test_AdminHuyDonHang()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }
            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(2000);

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement firstConfirmButton = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'btn btn-danger')])[1]")));
            wait.Until(d => firstConfirmButton.Displayed && firstConfirmButton.Enabled);
            firstConfirmButton.Click();
            Thread.Sleep(2000);


            Actions action1 = new Actions(driver);
            IWebElement dang1 = driver.FindElement(By.ClassName("icon-user-fullname"));
            action1.MoveToElement(dang1).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement dangxuat = driver.FindElement(By.XPath("//a[@href='/user/DangXuat']"));
            dangxuat.Click();
            Thread.Sleep(2000);

            foreach (var data in loginData1)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }

            Actions action = new Actions(driver);
            IWebElement dang = driver.FindElement(By.CssSelector("a[href='/home/donhang']"));
            action.MoveToElement(dang).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement donHangButton = driver.FindElement(By.XPath("//a[contains(text(),'Đơn hàng của tôi')]"));
            donHangButton.Click();
            Thread.Sleep(2000);

            IWebElement firstOrderRow = driver.FindElement(By.XPath("//table/tbody/tr[1]")); // Chọn hàng đầu tiên
            IWebElement trangThaiElement = firstOrderRow.FindElement(By.XPath("./td[3]")); // Chọn cột 5
            string actualTrangThai = trangThaiElement.Text.Trim(); // Lấy nội dung và loại bỏ khoảng trắng
            Assert.AreEqual("Hủy đơn hàng thành công", actualTrangThai, "Trạng thái không khớp với mong đợi!");

        }

        [Test]
        public void Test_AdminChapNHanHuy()
        {
            foreach (var data in loginData)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }
            IWebElement donHangLink = driver.FindElement(By.LinkText("Đơn hàng"));
            donHangLink.Click();
            Thread.Sleep(2000);

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement firstConfirmButton = wait.Until(d => d.FindElement(By.XPath("(//button[contains(@class, 'btn btn-warning')])[1]")));
            wait.Until(d => firstConfirmButton.Displayed && firstConfirmButton.Enabled);
            firstConfirmButton.Click();
            Thread.Sleep(2000);


            Actions action1 = new Actions(driver);
            IWebElement dang1 = driver.FindElement(By.ClassName("icon-user-fullname"));
            action1.MoveToElement(dang1).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement dangxuat = driver.FindElement(By.XPath("//a[@href='/user/DangXuat']"));
            dangxuat.Click();
            Thread.Sleep(2000);

            foreach (var data in loginData1)
            {
                driver.Navigate().GoToUrl(baseUrl);
                Thread.Sleep(2000);

                IWebElement loginIcon = driver.FindElement(By.ClassName("bx-user"));
                loginIcon.Click();
                Thread.Sleep(2000);

                IWebElement usernameInput = driver.FindElement(By.Id("Username"));
                IWebElement passwordInput = driver.FindElement(By.Id("Matkhau"));

                usernameInput.SendKeys(data["Username"]);
                passwordInput.SendKeys(data["Matkhau"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
            }

            Actions action = new Actions(driver);
            IWebElement dang = driver.FindElement(By.CssSelector("a[href='/home/donhang']"));
            action.MoveToElement(dang).Perform();
            Thread.Sleep(2000); // Chờ form hiển thị

            IWebElement donHangButton = driver.FindElement(By.XPath("//a[contains(text(),'Đơn hàng của tôi')]"));
            donHangButton.Click();
            Thread.Sleep(2000);

            IWebElement firstOrderRow = driver.FindElement(By.XPath("//table/tbody/tr[1]")); // Chọn hàng đầu tiên
            IWebElement trangThaiElement = firstOrderRow.FindElement(By.XPath("./td[3]")); // Chọn cột 5
            string actualTrangThai = trangThaiElement.Text.Trim(); // Lấy nội dung và loại bỏ khoảng trắng
            Assert.AreEqual("Hủy đơn hàng thành công", actualTrangThai, "Trạng thái không khớp với mong đợi!");

        }










        [TearDown]
        public void TearDown()
        {
            driver.Dispose();
        }

    }
}