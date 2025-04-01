using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Data;

namespace TestThanhToan
{
    public class Tests
    {
        private IWebDriver driver;
        private IWebDriver paypal;

        private List<Dictionary<string, string>> loginData;
        private List<Dictionary<string, string>> cartData;
        private List<Dictionary<string, string>> vnpayData;

        private string  vnpay=   "https://sandbox.vnpayment.vn/paymentv2/Transaction/PaymentMethod.html?token=57b0ee997bed40ac8f2bc3319f073d3c";
        private string baseUrl = "https://localhost:7053";
        private string paypalUrl = "https://api.sandbox.paypal.com";


        [SetUp]
        public void Setup()
        {
                     driver = new ChromeDriver();
            driver.Navigate().GoToUrl(baseUrl);
            string filePath = @"C:\File\HK2_Nam3\Đảm Bảo Chất Lượng Phần Mềm_LT_Ngân\TestData.xlsx";
            loginData = ReadExcel(filePath, "Login");
            cartData = ReadExcel(filePath, "Cart");
            vnpayData = ReadExcel(filePath, "VnPay");

            Thread.Sleep(2000);
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
        public void DangNhap()
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
                passwordInput.SendKeys(data["Password"]);
                passwordInput.SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }
        }

        [Test]
        public void TestMuaHangCODThanhCong()
        {

            var product = cartData[1];
            DangNhap();            // Sản phẩm đầu tiên


            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);

            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("cod");
            Thread.Sleep(2000);
          
                driver.FindElement(By.Name("tennguoinhan")).SendKeys(product["TenNguoiNhan"]);



                driver.FindElement(By.Name("sdtnguoinhan")).SendKeys(product["SoDienThoai"]);


                driver.FindElement(By.Id("address_detail")).SendKeys(product["DiaChi"]);
                Thread.Sleep(1000);

          
            var province = new SelectElement(driver.FindElement(By.Id("province")));
            province.SelectByText(product["DiaChiTinh"]);
            Thread.Sleep(2000);


            var district = new SelectElement(driver.FindElement(By.Id("district")));
            district.SelectByIndex(1);
            Thread.Sleep(2000);

            var ward = new SelectElement(driver.FindElement(By.Id("ward")));
            ward.SelectByIndex(1);

            Thread.Sleep(2000);


            IWebElement thanhToanButton = driver.FindElement(By.Id("cod-button"));

            thanhToanButton.Click();
            Thread.Sleep(7000);

            Assert.That(driver.Url, Does.Contain("https://localhost:7053/GioHang/LuuDonHang"), "Không phải trang đăng nhập thành công");


        }
        [Test]
        public void TestMuaHangCODVoiNhapThongTinDonHangBangKyTuDacBiet()
        {

            var product = cartData[3];
            DangNhap();            // Sản phẩm đầu tiên


            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);

            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("cod");
            Thread.Sleep(2000);

            driver.FindElement(By.Name("tennguoinhan")).SendKeys(product["TenNguoiNhan"]);



            driver.FindElement(By.Name("sdtnguoinhan")).SendKeys(product["SoDienThoai"]);


            driver.FindElement(By.Id("address_detail")).SendKeys(product["DiaChi"]);
            Thread.Sleep(1000);


            var province = new SelectElement(driver.FindElement(By.Id("province")));
            province.SelectByText(product["DiaChiTinh"]);
            Thread.Sleep(2000);


            var district = new SelectElement(driver.FindElement(By.Id("district")));
            district.SelectByIndex(1);
            Thread.Sleep(2000);

            var ward = new SelectElement(driver.FindElement(By.Id("ward")));
            ward.SelectByIndex(1);

            Thread.Sleep(2000);


            //IWebElement thanhToanButton = driver.FindElement(By.Id("cod-button"));

            //thanhToanButton.Click();
            IWebElement inputField = driver.FindElement(By.Name("tennguoinhan"));

            string validationMessage = (string)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].validationMessage;", inputField);

            Assert.That(validationMessage, Is.EqualTo("Please fill out this field."));


        }
        [Test]
        public void TestMuaHangPayPalThanhCong()
        {
            var product = cartData[0]; // Sản phẩm đầu tiên


            DangNhap();


            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);

            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("paypal");
            Thread.Sleep(2000);

          
                driver.FindElement(By.Name("tennguoinhan")).SendKeys(product["TenNguoiNhan"]);



                driver.FindElement(By.Name("sdtnguoinhan")).SendKeys(product["SoDienThoai"]);


                driver.FindElement(By.Id("address_detail")).SendKeys(product["DiaChi"]);
                Thread.Sleep(1000);

            
            var province = new SelectElement(driver.FindElement(By.Id("province")));
            province.SelectByText(product["DiaChiTinh"]);
            Thread.Sleep(2000);


            var district = new SelectElement(driver.FindElement(By.Id("district")));
            district.SelectByIndex(1);
            Thread.Sleep(2000);

            var ward = new SelectElement(driver.FindElement(By.Id("ward")));
            ward.SelectByIndex(1);

            Thread.Sleep(2000);

            IWebElement thanhToanButton = driver.FindElement(By.Id("paypal-button"));

            thanhToanButton.Click();
            Thread.Sleep(100);

            driver.FindElement(By.Id("email")).SendKeys("sb-d9mod33035231@personal.example.com");
            Thread.Sleep(1000);

            driver.FindElement(By.Id("btnNext")).Click();
            Thread.Sleep(1000);

            driver.FindElement(By.Id("password")).SendKeys("u|^iXU7K\r\n");
            Thread.Sleep(1000);

            driver.FindElement(By.Id("payment-submit-btn\r\n")).Click();
            Thread.Sleep(2000);

            IWebElement successMessage = driver.FindElement(By.TagName("h2"));
            Assert.That(successMessage.Text, Is.EqualTo("Bạn đã mua hàng thành công!"));

        }
   
        [Test]
        public void TestMuaHangVnPay()
        {
            var product = cartData[0]; // Sản phẩm đầu tiên
            var vnpay = vnpayData[0];
            DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);



            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("vnpay");
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('form-vnpay').classList.remove('hidden');");
            Thread.Sleep(1000); // Chờ form hiển thị


            IWebElement cartTitle = driver.FindElement(By.XPath("(//input[@name='tennguoinhan'])[2]"));
            cartTitle.SendKeys(product["TenNguoiNhan"]);
            Thread.Sleep(3000);

            // Chọn phần tử thứ 2 với name là 'sdtnguoinhan'
            IWebElement cartTitle2 = driver.FindElement(By.XPath("(//input[@name='sdtnguoinhan'])[2]"));
            cartTitle2.SendKeys(product["SoDienThoai"]);
            Thread.Sleep(3000);

            // Chọn phần tử với id 'address_detail_vnpay' (giả sử có duy nhất)
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).Clear();
                driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).SendKeys(product["DiaChi"]);


                Thread.Sleep(1000);


            IWebElement provinceDropdown = driver.FindElement(By.Id("province-vnpay"));
            SelectElement provinceSelect = new SelectElement(provinceDropdown);
            provinceSelect.SelectByIndex(1); // Chọn tỉnh (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn quận/huyện
            IWebElement districtDropdown = driver.FindElement(By.Id("district-vnpay"));
            SelectElement districtSelect = new SelectElement(districtDropdown);
            districtSelect.SelectByIndex(1); // Chọn quận/huyện (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn xã/phường
            IWebElement wardDropdown = driver.FindElement(By.Id("ward-vnpay"));
            SelectElement wardSelect = new SelectElement(wardDropdown);
            wardSelect.SelectByIndex(1); // Chọn xã/phường (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);


            // Click nút thanh toán VNPAY


            IWebElement thanhToanButton = driver.FindElement(By.CssSelector("button.btn-success"));
            thanhToanButton.Click();
            Thread.Sleep(3000);

            // Tìm phần tử div có các thuộc tính data-bs-toggle, data-bs-target, và class="list-method-button"
            IWebElement accordionListButton = driver.FindElement(By.XPath("//div[@class='list-method-button' and @data-bs-toggle='collapse' and @data-bs-target='#accordionList2']"));
            accordionListButton.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click

            // Tìm phần tử div có class 'list-bank-item-inner' và style chứa background-image với URL cụ thể
            IWebElement bankItem = driver.FindElement(By.XPath("//div[@class='list-bank-item-inner' and contains(@style, 'background-image: url(/paymentv2/images/img/logos/bank/big/ncb.svg)')]"));
            bankItem.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click


            // Tìm input có placeholder là 'Nhập số thẻ' và nhập số thẻ vào
            // Tìm input có id là 'card_number_mask' và nhập số thẻ vào
            IWebElement cardNumberInput = driver.FindElement(By.Id("card_number_mask"));
            cardNumberInput.SendKeys(vnpay["SoThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập số thẻ

            // Tìm input có id là 'cardHolder' và nhập tên vào
            IWebElement cardHolderInput = driver.FindElement(By.Id("cardHolder"));
            cardHolderInput.SendKeys(vnpay["TenThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập tên

            // Tìm input có id là 'cardDate' và nhập dữ liệu vào
            IWebElement cardDateInput = driver.FindElement(By.Id("cardDate"));
            cardDateInput.SendKeys(vnpay["NgayPhatHanh"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập dữ liệu

           driver.FindElement(By.Id("btnContinue")).Click();
            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            IWebElement dongYButton = driver.FindElement(By.XPath("//span[@class='ubtn-text' and contains(text(),'Đồng ý & Tiếp tục')]"));
            Thread.Sleep(1000); // Chờ sau khi click

            // Click vào nút
            dongYButton.Click();
            Thread.Sleep(1000); // Chờ sau khi click

            IWebElement a = driver.FindElement(By.Id("otpvalue"));
            a.SendKeys("123456");
            Thread.Sleep(1000); // Chờ sau khi click

            driver.FindElement(By.Id("btnConfirm")).Click();
            Thread.Sleep(3000); // Chờ sau khi click


            IWebElement successMessage = driver.FindElement(By.TagName("h2"));
            Assert.That(successMessage.Text, Is.EqualTo("Bạn đã mua hàng thành công!"));



        }
        [Test]
        public void TestMuaHangVnPayNhungNhapSaiTaiKhoan()
        {
            var product = cartData[0]; // Sản phẩm đầu tiên
            var vnpay = vnpayData[1];

            DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);



            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("vnpay");
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('form-vnpay').classList.remove('hidden');");
            Thread.Sleep(1000); // Chờ form hiển thị


            IWebElement cartTitle = driver.FindElement(By.XPath("(//input[@name='tennguoinhan'])[2]"));
            cartTitle.SendKeys(product["TenNguoiNhan"]);
            Thread.Sleep(3000);

            // Chọn phần tử thứ 2 với name là 'sdtnguoinhan'
            IWebElement cartTitle2 = driver.FindElement(By.XPath("(//input[@name='sdtnguoinhan'])[2]"));
            cartTitle2.SendKeys(product["SoDienThoai"]);
            Thread.Sleep(3000);

            // Chọn phần tử với id 'address_detail_vnpay' (giả sử có duy nhất)
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).Clear();
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).SendKeys(product["DiaChi"]);


            Thread.Sleep(1000);


            IWebElement provinceDropdown = driver.FindElement(By.Id("province-vnpay"));
            SelectElement provinceSelect = new SelectElement(provinceDropdown);
            provinceSelect.SelectByIndex(1); // Chọn tỉnh (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn quận/huyện
            IWebElement districtDropdown = driver.FindElement(By.Id("district-vnpay"));
            SelectElement districtSelect = new SelectElement(districtDropdown);
            districtSelect.SelectByIndex(1); // Chọn quận/huyện (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn xã/phường
            IWebElement wardDropdown = driver.FindElement(By.Id("ward-vnpay"));
            SelectElement wardSelect = new SelectElement(wardDropdown);
            wardSelect.SelectByIndex(1); // Chọn xã/phường (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);


            // Click nút thanh toán VNPAY


            IWebElement thanhToanButton = driver.FindElement(By.CssSelector("button.btn-success"));
            thanhToanButton.Click();
            Thread.Sleep(3000);

            // Tìm phần tử div có các thuộc tính data-bs-toggle, data-bs-target, và class="list-method-button"
            IWebElement accordionListButton = driver.FindElement(By.XPath("//div[@class='list-method-button' and @data-bs-toggle='collapse' and @data-bs-target='#accordionList2']"));
            accordionListButton.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click

            // Tìm phần tử div có class 'list-bank-item-inner' và style chứa background-image với URL cụ thể
            IWebElement bankItem = driver.FindElement(By.XPath("//div[@class='list-bank-item-inner' and contains(@style, 'background-image: url(/paymentv2/images/img/logos/bank/big/ncb.svg)')]"));
            bankItem.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click


            // Tìm input có placeholder là 'Nhập số thẻ' và nhập số thẻ vào
            // Tìm input có id là 'card_number_mask' và nhập số thẻ vào
            IWebElement cardNumberInput = driver.FindElement(By.Id("card_number_mask"));
            cardNumberInput.SendKeys(vnpay["SoThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập số thẻ
            // Tìm input có id là 'cardHolder' và nhập tên vào
            IWebElement cardHolderInput = driver.FindElement(By.Id("cardHolder"));
            cardHolderInput.SendKeys(vnpay["TenThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập tên

            // Tìm input có id là 'cardDate' và nhập dữ liệu vào
            IWebElement cardDateInput = driver.FindElement(By.Id("cardDate"));
            cardDateInput.SendKeys(vnpay["NgayPhatHanh"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập dữ liệu

            driver.FindElement(By.Id("btnContinue")).Click();
            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            IWebElement dongYButton = driver.FindElement(By.XPath("//span[@class='ubtn-text' and contains(text(),'Đồng ý & Tiếp tục')]"));
            Thread.Sleep(1000); // Chờ sau khi click
             cardNumberInput = driver.FindElement(By.Id("card_number_mask"));

            // Click vào nút

            // Tìm phần tử thông báo lỗi bằng class cụ thể
            string errorMessage = cardNumberInput.GetAttribute("data-parsley-length-message");

            // Kiểm tra xem thông báo có đúng không
            Assert.That(errorMessage, Is.EqualTo("Số thẻ không hợp lệ"), "Thông báo lỗi không đúng!");





        }
        [Test]
        public void TestMuaHangVnPayNhungDeTrongSoThe()
        {
            var product = cartData[0]; // Sản phẩm đầu tiên
            var vnpay = vnpayData[0];
            DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);



            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("vnpay");
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('form-vnpay').classList.remove('hidden');");
            Thread.Sleep(1000); // Chờ form hiển thị


            IWebElement cartTitle = driver.FindElement(By.XPath("(//input[@name='tennguoinhan'])[2]"));
            cartTitle.SendKeys(product["TenNguoiNhan"]);
            Thread.Sleep(3000);

            // Chọn phần tử thứ 2 với name là 'sdtnguoinhan'
            IWebElement cartTitle2 = driver.FindElement(By.XPath("(//input[@name='sdtnguoinhan'])[2]"));
            cartTitle2.SendKeys(product["SoDienThoai"]);
            Thread.Sleep(3000);

            // Chọn phần tử với id 'address_detail_vnpay' (giả sử có duy nhất)
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).Clear();
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).SendKeys(product["DiaChi"]);


            Thread.Sleep(1000);


            IWebElement provinceDropdown = driver.FindElement(By.Id("province-vnpay"));
            SelectElement provinceSelect = new SelectElement(provinceDropdown);
            provinceSelect.SelectByIndex(1); // Chọn tỉnh (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn quận/huyện
            IWebElement districtDropdown = driver.FindElement(By.Id("district-vnpay"));
            SelectElement districtSelect = new SelectElement(districtDropdown);
            districtSelect.SelectByIndex(1); // Chọn quận/huyện (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn xã/phường
            IWebElement wardDropdown = driver.FindElement(By.Id("ward-vnpay"));
            SelectElement wardSelect = new SelectElement(wardDropdown);
            wardSelect.SelectByIndex(1); // Chọn xã/phường (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);


            // Click nút thanh toán VNPAY


            IWebElement thanhToanButton = driver.FindElement(By.CssSelector("button.btn-success"));
            thanhToanButton.Click();
            Thread.Sleep(3000);

            // Tìm phần tử div có các thuộc tính data-bs-toggle, data-bs-target, và class="list-method-button"
            IWebElement accordionListButton = driver.FindElement(By.XPath("//div[@class='list-method-button' and @data-bs-toggle='collapse' and @data-bs-target='#accordionList2']"));
            accordionListButton.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click

            // Tìm phần tử div có class 'list-bank-item-inner' và style chứa background-image với URL cụ thể
            IWebElement bankItem = driver.FindElement(By.XPath("//div[@class='list-bank-item-inner' and contains(@style, 'background-image: url(/paymentv2/images/img/logos/bank/big/ncb.svg)')]"));
            bankItem.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click


            // Tìm input có placeholder là 'Nhập số thẻ' và nhập số thẻ vào
            // Tìm input có id là 'card_number_mask' và nhập số thẻ vào
            IWebElement cardNumberInput = driver.FindElement(By.Id("card_number_mask"));
            Thread.Sleep(2000); // Chờ một chút sau khi nhập số thẻ

            IWebElement cardHolderInput = driver.FindElement(By.Id("cardHolder"));
            cardHolderInput.SendKeys(vnpay["TenThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập tên

            // Tìm input có id là 'cardDate' và nhập dữ liệu vào
            IWebElement cardDateInput = driver.FindElement(By.Id("cardDate"));
            cardDateInput.SendKeys(vnpay["NgayPhatHanh"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập dữ liệu/ Chờ một chút sau khi nhập dữ liệu

            driver.FindElement(By.Id("btnContinue")).Click();
            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            Thread.Sleep(1000); // Chờ sau khi click
                                // Tìm input số thẻ
             cardNumberInput = driver.FindElement(By.Id("card_number_mask"));

            // Lấy giá trị của thuộc tính `data-parsley-required-message`
            string requiredMessage = cardNumberInput.GetAttribute("data-parsley-required-message");

            // Kiểm tra xem thông báo có đúng không
            Assert.That(requiredMessage, Is.EqualTo("Quý khách vui lòng nhập Số thẻ"), "Thông báo lỗi không đúng!");






        }
        [Test]
        public void TestMuaHangKhiChuaDangNhap()
        {


            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();


            Assert.That(driver.Url, Does.Contain("https://localhost:7053/User/Login"), "Không phải trang đăng nhập!");


        }
        [Test]

        public void TestMuaHangKhiKhongNhapThongTin()
        {

                DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



          

            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);

            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("cod");
            Thread.Sleep(2000);

         

            IWebElement inputField = driver.FindElement(By.Name("tennguoinhan"));

            string validationMessage = (string)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].validationMessage;", inputField);

            Assert.That(validationMessage, Is.EqualTo("Please fill out this field."));


        }

        [Test]
        public void TestThanhToanKhiNhapSaiTaiKhoanPayPal()
        {

            var product = cartData[0]; // Sản phẩm đầu tiên


            DangNhap();


            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);

            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("paypal");
            Thread.Sleep(2000);


            driver.FindElement(By.Name("tennguoinhan")).SendKeys(product["TenNguoiNhan"]);



            driver.FindElement(By.Name("sdtnguoinhan")).SendKeys(product["SoDienThoai"]);


            driver.FindElement(By.Id("address_detail")).SendKeys(product["DiaChi"]);
            Thread.Sleep(1000);


            var province = new SelectElement(driver.FindElement(By.Id("province")));
            province.SelectByText(product["DiaChiTinh"]);
            Thread.Sleep(2000);


            var district = new SelectElement(driver.FindElement(By.Id("district")));
            district.SelectByIndex(1);
            Thread.Sleep(2000);

            var ward = new SelectElement(driver.FindElement(By.Id("ward")));
            ward.SelectByIndex(1);

            Thread.Sleep(2000);

            IWebElement thanhToanButton = driver.FindElement(By.Id("paypal-button"));

            thanhToanButton.Click();
            Thread.Sleep(100);

            driver.FindElement(By.Id("email")).SendKeys("sb-d9mod33035231@personal.example.com");
            Thread.Sleep(1000);

            driver.FindElement(By.Id("btnNext")).Click();
            Thread.Sleep(3000);

            driver.FindElement(By.Id("password")).SendKeys("u|^iX\r\n");
            Thread.Sleep(3000);

            driver.FindElement(By.Id("btnLogin")).Click();
            Thread.Sleep(2000);

            // ✅ Dùng Assert.That để kiểm tra thông báo lỗi
            IWebElement errorMessage = driver.FindElement(By.ClassName("notification-critical"));
            string actualErrorText = errorMessage.Text;
            string expectedErrorText = "Một số thông tin của bạn không khớp. Vui lòng thử lại, thay đổi địa chỉ email hoặc yêu cầu trợ giúp nếu bạn quên mật khẩu.";

            Assert.That(actualErrorText, Is.EqualTo(expectedErrorText), "Thông báo lỗi không đúng!");
        }



        [Test]
        public void TestThanhToanKhiTaiKhoanKhongConDuTien()
        {

            var product = cartData[0]; // Sản phẩm đầu tiên
            var vnpay = vnpayData[2];    
                DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);



            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("vnpay");
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('form-vnpay').classList.remove('hidden');");
            Thread.Sleep(1000); // Chờ form hiển thị


            IWebElement cartTitle = driver.FindElement(By.XPath("(//input[@name='tennguoinhan'])[2]"));
            cartTitle.SendKeys(product["TenNguoiNhan"]);
            Thread.Sleep(3000);

            // Chọn phần tử thứ 2 với name là 'sdtnguoinhan'
            IWebElement cartTitle2 = driver.FindElement(By.XPath("(//input[@name='sdtnguoinhan'])[2]"));
            cartTitle2.SendKeys(product["SoDienThoai"]);
            Thread.Sleep(3000);

            // Chọn phần tử với id 'address_detail_vnpay' (giả sử có duy nhất)
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).Clear();
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).SendKeys(product["DiaChi"]);


            Thread.Sleep(1000);


            IWebElement provinceDropdown = driver.FindElement(By.Id("province-vnpay"));
            SelectElement provinceSelect = new SelectElement(provinceDropdown);
            provinceSelect.SelectByIndex(1); // Chọn tỉnh (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn quận/huyện
            IWebElement districtDropdown = driver.FindElement(By.Id("district-vnpay"));
            SelectElement districtSelect = new SelectElement(districtDropdown);
            districtSelect.SelectByIndex(1); // Chọn quận/huyện (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn xã/phường
            IWebElement wardDropdown = driver.FindElement(By.Id("ward-vnpay"));
            SelectElement wardSelect = new SelectElement(wardDropdown);
            wardSelect.SelectByIndex(1); // Chọn xã/phường (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);


            // Click nút thanh toán VNPAY


            IWebElement thanhToanButton = driver.FindElement(By.CssSelector("button.btn-success"));
            thanhToanButton.Click();
            Thread.Sleep(3000);

            // Tìm phần tử div có các thuộc tính data-bs-toggle, data-bs-target, và class="list-method-button"
            IWebElement accordionListButton = driver.FindElement(By.XPath("//div[@class='list-method-button' and @data-bs-toggle='collapse' and @data-bs-target='#accordionList2']"));
            accordionListButton.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click

            // Tìm phần tử div có class 'list-bank-item-inner' và style chứa background-image với URL cụ thể
            IWebElement bankItem = driver.FindElement(By.XPath("//div[@class='list-bank-item-inner' and contains(@style, 'background-image: url(/paymentv2/images/img/logos/bank/big/ncb.svg)')]"));
            bankItem.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click


            // Tìm input có placeholder là 'Nhập số thẻ' và nhập số thẻ vào
            // Tìm input có id là 'card_number_mask' và nhập số thẻ vào
            IWebElement cardNumberInput = driver.FindElement(By.Id("card_number_mask"));
            cardNumberInput.SendKeys(vnpay["SoThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập số thẻ

            // Tìm input có id là 'cardHolder' và nhập tên vào
            IWebElement cardHolderInput = driver.FindElement(By.Id("cardHolder"));
            cardHolderInput.SendKeys(vnpay["TenThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập tên

            // Tìm input có id là 'cardDate' và nhập dữ liệu vào
            IWebElement cardDateInput = driver.FindElement(By.Id("cardDate"));
            cardDateInput.SendKeys(vnpay["NgayPhatHanh"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập dữ liệu

            driver.FindElement(By.Id("btnContinue")).Click();
            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            Thread.Sleep(3000); // Chờ sau khi click
                                // Tìm input số thẻ


            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            IWebElement dongYButton = driver.FindElement(By.XPath("//span[@class='ubtn-text' and contains(text(),'Đồng ý & Tiếp tục')]"));
            Thread.Sleep(1000); // Chờ sau khi click

            // Click vào nút
            dongYButton.Click();
            Thread.Sleep(1000); // Chờ sau khi click



            IWebElement errorMessageElement = driver.FindElement(By.XPath("//label[contains(@id, 'lb_message_error')]"));
             string errorMessageText = errorMessageElement.Text;
            Thread.Sleep(2000);


            var allLabels = driver.FindElements(By.TagName("label"));
            foreach (var label in allLabels)
            {
                Console.WriteLine("Label found: " + label.Text);
            }

            Assert.That(errorMessageText, Is.EqualTo("Tài khoản của khách hàng không đủ số dư để thực hiện giao dịch"));

            Thread.Sleep(1000); // Chờ sau khi click


        }
        [Test]
        public void TestThanhToanKhiTaiKhoanBiKhoa()
        {

            var product = cartData[0]; // Sản phẩm đầu tiên
            var vnpay = vnpayData[3];

            DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);



            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("vnpay");
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('form-vnpay').classList.remove('hidden');");
            Thread.Sleep(1000); // Chờ form hiển thị


            IWebElement cartTitle = driver.FindElement(By.XPath("(//input[@name='tennguoinhan'])[2]"));
            cartTitle.SendKeys(product["TenNguoiNhan"]);
            Thread.Sleep(3000);

            // Chọn phần tử thứ 2 với name là 'sdtnguoinhan'
            IWebElement cartTitle2 = driver.FindElement(By.XPath("(//input[@name='sdtnguoinhan'])[2]"));
            cartTitle2.SendKeys(product["SoDienThoai"]);
            Thread.Sleep(3000);

            // Chọn phần tử với id 'address_detail_vnpay' (giả sử có duy nhất)
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).Clear();
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).SendKeys(product["DiaChi"]);


            Thread.Sleep(1000);


            IWebElement provinceDropdown = driver.FindElement(By.Id("province-vnpay"));
            SelectElement provinceSelect = new SelectElement(provinceDropdown);
            provinceSelect.SelectByIndex(1); // Chọn tỉnh (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn quận/huyện
            IWebElement districtDropdown = driver.FindElement(By.Id("district-vnpay"));
            SelectElement districtSelect = new SelectElement(districtDropdown);
            districtSelect.SelectByIndex(1); // Chọn quận/huyện (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn xã/phường
            IWebElement wardDropdown = driver.FindElement(By.Id("ward-vnpay"));
            SelectElement wardSelect = new SelectElement(wardDropdown);
            wardSelect.SelectByIndex(1); // Chọn xã/phường (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);


            // Click nút thanh toán VNPAY


            IWebElement thanhToanButton = driver.FindElement(By.CssSelector("button.btn-success"));
            thanhToanButton.Click();
            Thread.Sleep(3000);

            // Tìm phần tử div có các thuộc tính data-bs-toggle, data-bs-target, và class="list-method-button"
            IWebElement accordionListButton = driver.FindElement(By.XPath("//div[@class='list-method-button' and @data-bs-toggle='collapse' and @data-bs-target='#accordionList2']"));
            accordionListButton.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click

            // Tìm phần tử div có class 'list-bank-item-inner' và style chứa background-image với URL cụ thể
            IWebElement bankItem = driver.FindElement(By.XPath("//div[@class='list-bank-item-inner' and contains(@style, 'background-image: url(/paymentv2/images/img/logos/bank/big/ncb.svg)')]"));
            bankItem.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click


            // Tìm input có placeholder là 'Nhập số thẻ' và nhập số thẻ vào
            // Tìm input có id là 'card_number_mask' và nhập số thẻ vào
            IWebElement cardNumberInput = driver.FindElement(By.Id("card_number_mask"));
            cardNumberInput.SendKeys(vnpay["SoThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập số thẻ

            // Tìm input có id là 'cardHolder' và nhập tên vào
            IWebElement cardHolderInput = driver.FindElement(By.Id("cardHolder"));
            cardHolderInput.SendKeys(vnpay["TenThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập tên

            // Tìm input có id là 'cardDate' và nhập dữ liệu vào
            IWebElement cardDateInput = driver.FindElement(By.Id("cardDate"));
            cardDateInput.SendKeys(vnpay["NgayPhatHanh"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập dữ liệu
            driver.FindElement(By.Id("btnContinue")).Click();
            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            Thread.Sleep(3000); // Chờ sau khi click
                                // Tìm input số thẻ


            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            IWebElement dongYButton = driver.FindElement(By.XPath("//span[@class='ubtn-text' and contains(text(),'Đồng ý & Tiếp tục')]"));
            Thread.Sleep(1000); // Chờ sau khi click

            // Click vào nút
            dongYButton.Click();
            Thread.Sleep(1000); // Chờ sau khi click



            IWebElement errorMessageElement = driver.FindElement(By.XPath("//label[contains(@id, 'lb_message_error')]"));
            string errorMessageText = errorMessageElement.Text;
            Thread.Sleep(2000);


            var allLabels = driver.FindElements(By.TagName("label"));
            foreach (var label in allLabels)
            {
                Console.WriteLine("Label found: " + label.Text);
            }

            Assert.That(errorMessageText, Is.EqualTo("Thẻ bị khóa. Vui lòng liên hệ ngân hàng phát hành để được hỗ trợ."));

            Thread.Sleep(1000); // Chờ sau khi click


        }
        [Test]
        public void TestThanhToanKhiTheChuaKichHoat()
        {

            var product = cartData[0]; // Sản phẩm đầu tiên
            var vnpay = vnpayData[4];

            DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);



            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("vnpay");
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('form-vnpay').classList.remove('hidden');");
            Thread.Sleep(1000); // Chờ form hiển thị


            IWebElement cartTitle = driver.FindElement(By.XPath("(//input[@name='tennguoinhan'])[2]"));
            cartTitle.SendKeys(product["TenNguoiNhan"]);
            Thread.Sleep(3000);

            // Chọn phần tử thứ 2 với name là 'sdtnguoinhan'
            IWebElement cartTitle2 = driver.FindElement(By.XPath("(//input[@name='sdtnguoinhan'])[2]"));
            cartTitle2.SendKeys(product["SoDienThoai"]);
            Thread.Sleep(3000);

            // Chọn phần tử với id 'address_detail_vnpay' (giả sử có duy nhất)
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).Clear();
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).SendKeys(product["DiaChi"]);


            Thread.Sleep(1000);


            IWebElement provinceDropdown = driver.FindElement(By.Id("province-vnpay"));
            SelectElement provinceSelect = new SelectElement(provinceDropdown);
            provinceSelect.SelectByIndex(1); // Chọn tỉnh (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn quận/huyện
            IWebElement districtDropdown = driver.FindElement(By.Id("district-vnpay"));
            SelectElement districtSelect = new SelectElement(districtDropdown);
            districtSelect.SelectByIndex(1); // Chọn quận/huyện (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn xã/phường
            IWebElement wardDropdown = driver.FindElement(By.Id("ward-vnpay"));
            SelectElement wardSelect = new SelectElement(wardDropdown);
            wardSelect.SelectByIndex(1); // Chọn xã/phường (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);


            // Click nút thanh toán VNPAY


            IWebElement thanhToanButton = driver.FindElement(By.CssSelector("button.btn-success"));
            thanhToanButton.Click();
            Thread.Sleep(3000);

            // Tìm phần tử div có các thuộc tính data-bs-toggle, data-bs-target, và class="list-method-button"
            IWebElement accordionListButton = driver.FindElement(By.XPath("//div[@class='list-method-button' and @data-bs-toggle='collapse' and @data-bs-target='#accordionList2']"));
            accordionListButton.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click

            // Tìm phần tử div có class 'list-bank-item-inner' và style chứa background-image với URL cụ thể
            IWebElement bankItem = driver.FindElement(By.XPath("//div[@class='list-bank-item-inner' and contains(@style, 'background-image: url(/paymentv2/images/img/logos/bank/big/ncb.svg)')]"));
            bankItem.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click


            // Tìm input có placeholder là 'Nhập số thẻ' và nhập số thẻ vào
            // Tìm input có id là 'card_number_mask' và nhập số thẻ vào
            IWebElement cardNumberInput = driver.FindElement(By.Id("card_number_mask"));
            cardNumberInput.SendKeys(vnpay["SoThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập số thẻ

            // Tìm input có id là 'cardHolder' và nhập tên vào
            IWebElement cardHolderInput = driver.FindElement(By.Id("cardHolder"));
            cardHolderInput.SendKeys(vnpay["TenThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập tên

            // Tìm input có id là 'cardDate' và nhập dữ liệu vào
            IWebElement cardDateInput = driver.FindElement(By.Id("cardDate"));
            cardDateInput.SendKeys(vnpay["NgayPhatHanh"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập dữ liệu
            driver.FindElement(By.Id("btnContinue")).Click();
            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            Thread.Sleep(3000); // Chờ sau khi click
                                // Tìm input số thẻ


            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            IWebElement dongYButton = driver.FindElement(By.XPath("//span[@class='ubtn-text' and contains(text(),'Đồng ý & Tiếp tục')]"));
            Thread.Sleep(1000); // Chờ sau khi click

            // Click vào nút
            dongYButton.Click();
            Thread.Sleep(1000); // Chờ sau khi click



            IWebElement errorMessageElement = driver.FindElement(By.XPath("//label[contains(@id, 'lb_message_error')]"));
            string errorMessageText = errorMessageElement.Text;
            Thread.Sleep(2000);


            var allLabels = driver.FindElements(By.TagName("label"));
            foreach (var label in allLabels)
            {
                Console.WriteLine("Label found: " + label.Text);
            }

            Assert.That(errorMessageText, Is.EqualTo("Thẻ chưa được kích hoạt"));

            Thread.Sleep(1000); // Chờ sau khi click


        }
        [Test]
        public void TestThanhToanKhiTheHetHan()
        {

            var product = cartData[0]; // Sản phẩm đầu tiên
            var vnpay = vnpayData[5];

            DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            driver.FindElement(By.Id("payment-method")).Click();
            Thread.Sleep(4000);



            var paymentMethodSelect = new SelectElement(driver.FindElement(By.Id("payment-method")));
            paymentMethodSelect.SelectByValue("vnpay");
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('form-vnpay').classList.remove('hidden');");
            Thread.Sleep(1000); // Chờ form hiển thị


            IWebElement cartTitle = driver.FindElement(By.XPath("(//input[@name='tennguoinhan'])[2]"));
            cartTitle.SendKeys(product["TenNguoiNhan"]);
            Thread.Sleep(3000);

            // Chọn phần tử thứ 2 với name là 'sdtnguoinhan'
            IWebElement cartTitle2 = driver.FindElement(By.XPath("(//input[@name='sdtnguoinhan'])[2]"));
            cartTitle2.SendKeys(product["SoDienThoai"]);
            Thread.Sleep(3000);

            // Chọn phần tử với id 'address_detail_vnpay' (giả sử có duy nhất)
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).Clear();
            driver.FindElement(By.XPath("//input[@id='address_detail_vnpay']")).SendKeys(product["DiaChi"]);


            Thread.Sleep(1000);


            IWebElement provinceDropdown = driver.FindElement(By.Id("province-vnpay"));
            SelectElement provinceSelect = new SelectElement(provinceDropdown);
            provinceSelect.SelectByIndex(1); // Chọn tỉnh (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn quận/huyện
            IWebElement districtDropdown = driver.FindElement(By.Id("district-vnpay"));
            SelectElement districtSelect = new SelectElement(districtDropdown);
            districtSelect.SelectByIndex(1); // Chọn quận/huyện (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);

            // Chọn xã/phường
            IWebElement wardDropdown = driver.FindElement(By.Id("ward-vnpay"));
            SelectElement wardSelect = new SelectElement(wardDropdown);
            wardSelect.SelectByIndex(1); // Chọn xã/phường (có thể thay bằng SelectByValue nếu bạn biết giá trị cụ thể)
            Thread.Sleep(2000);


            // Click nút thanh toán VNPAY


            IWebElement thanhToanButton = driver.FindElement(By.CssSelector("button.btn-success"));
            thanhToanButton.Click();
            Thread.Sleep(3000);

            // Tìm phần tử div có các thuộc tính data-bs-toggle, data-bs-target, và class="list-method-button"
            IWebElement accordionListButton = driver.FindElement(By.XPath("//div[@class='list-method-button' and @data-bs-toggle='collapse' and @data-bs-target='#accordionList2']"));
            accordionListButton.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click

            // Tìm phần tử div có class 'list-bank-item-inner' và style chứa background-image với URL cụ thể
            IWebElement bankItem = driver.FindElement(By.XPath("//div[@class='list-bank-item-inner' and contains(@style, 'background-image: url(/paymentv2/images/img/logos/bank/big/ncb.svg)')]"));
            bankItem.Click();
            Thread.Sleep(2000); // Chờ một chút sau khi click


            // Tìm input có placeholder là 'Nhập số thẻ' và nhập số thẻ vào
            // Tìm input có id là 'card_number_mask' và nhập số thẻ vào
            IWebElement cardNumberInput = driver.FindElement(By.Id("card_number_mask"));
            cardNumberInput.SendKeys(vnpay["SoThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập số thẻ

            // Tìm input có id là 'cardHolder' và nhập tên vào
            IWebElement cardHolderInput = driver.FindElement(By.Id("cardHolder"));
            cardHolderInput.SendKeys(vnpay["TenThe"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập tên

            // Tìm input có id là 'cardDate' và nhập dữ liệu vào
            IWebElement cardDateInput = driver.FindElement(By.Id("cardDate"));
            cardDateInput.SendKeys(vnpay["NgayPhatHanh"]);
            Thread.Sleep(2000); // Chờ một chút sau khi nhập dữ liệu
            driver.FindElement(By.Id("btnContinue")).Click();
            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            Thread.Sleep(3000); // Chờ sau khi click
                                // Tìm input số thẻ


            // Tìm phần tử có class 'ubtn-text' và chứa text 'Đồng ý & Tiếp tục'
            IWebElement dongYButton = driver.FindElement(By.XPath("//span[@class='ubtn-text' and contains(text(),'Đồng ý & Tiếp tục')]"));
            Thread.Sleep(1000); // Chờ sau khi click

            // Click vào nút
            dongYButton.Click();
            Thread.Sleep(1000); // Chờ sau khi click



            IWebElement errorMessageElement = driver.FindElement(By.XPath("//label[contains(@id, 'lb_message_error')]"));
            string errorMessageText = errorMessageElement.Text;
            Thread.Sleep(2000);


            var allLabels = driver.FindElements(By.TagName("label"));
            foreach (var label in allLabels)
            {
                Console.WriteLine("Label found: " + label.Text);
            }

            Assert.That(errorMessageText, Is.EqualTo("Thẻ bị hết hạn"));

            Thread.Sleep(1000); // Chờ sau khi click


        }
        [Test]
        public void TestNhapMaGiamGiaThanhCong()
        {

            var product = cartData[0]; // Sản phẩm đầu tiên

            DangNhap();

            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            IWebElement discountInput = driver.FindElement(By.Name("maKhuyenMai"));
          
                discountInput.SendKeys(product["MaKhuyenMai"]);
            
            var submitButton = driver.FindElement(By.XPath("//button[text()='Áp dụng']"));
            submitButton.Click();


            var giamGiaElement = driver.FindElement(By.XPath("//div[contains(@style, 'display: flex')]//h5[contains(@style, 'color: #FFE5B4')]"));
            Thread.Sleep(2000);

            // Lấy giá trị giảm giá hiển thị
            string giamGiaText = giamGiaElement.Text;
            Thread.Sleep(2000);

            // Assert để kiểm tra xem giá trị có đúng không
            Assert.That(giamGiaText, Is.EqualTo("20000 vnđ"), "Giảm giá không đúng!");
                        Thread.Sleep(2000);

        }
        [Test]
        public void MaGiamGiaQuaHan()
        {

            var product = cartData[2]; // Sản phẩm đầu tiên

            DangNhap();


            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            IWebElement discountInput = driver.FindElement(By.Name("maKhuyenMai"));
     

                discountInput.SendKeys(product["MaKhuyenMai"]);
            
            Thread.Sleep(2000);

            var submitButton = driver.FindElement(By.XPath("//button[text()='Áp dụng']"));
            submitButton.Click();
         
            Thread.Sleep(2000);




            var giamGiaElement = driver.FindElement(By.XPath("//div[contains(@style, 'display: flex')]//h5[contains(@style, 'color: #FFE5B4')]"));

            // Lấy giá trị giảm giá hiển thị
            string giamGiaText = giamGiaElement.Text;

            // Assert để kiểm tra xem giá trị có đúng không
            Assert.That(giamGiaText, Is.EqualTo("0 vnđ"));


        }
        [Test]
        public void MaGiamKhongDung()
        {

            var product = cartData[1]; // Sản phẩm đầu tiên

            DangNhap();


            IWebElement xemChiTietButton = driver.FindElement(By.XPath("//div[@class='product-container']//a[contains(text(),'Xem chi tiết')]"));
            xemChiTietButton.Click();

            Thread.Sleep(2000);

            driver.FindElement(By.ClassName("nutthemvaogio")).Click();
            Thread.Sleep(2000);



            IWebElement thanhToanButtons = driver.FindElement(By.LinkText("Thanh toán"));
            thanhToanButtons.Click();

            IWebElement discountInput = driver.FindElement(By.Name("maKhuyenMai"));
          

                discountInput.SendKeys(product["MaKhuyenMai"]);
            
            Thread.Sleep(2000);

            var submitButton = driver.FindElement(By.XPath("//button[text()='Áp dụng']"));
            submitButton.Click();

            Thread.Sleep(2000);




            var giamGiaElement = driver.FindElement(By.XPath("//div[contains(@style, 'display: flex')]//h5[contains(@style, 'color: #FFE5B4')]"));

            // Lấy giá trị giảm giá hiển thị
            string giamGiaText = giamGiaElement.Text;

            // Assert để kiểm tra xem giá trị có đúng không
            Assert.That(giamGiaText, Is.EqualTo("0 vnđ"));


        }
     
    
        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }
    }
}
    
