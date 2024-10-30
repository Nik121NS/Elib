using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Elib
{
    public partial class Form1 : Form
    {

        HttpClientHandler Handler;
        HttpClient Client;
        HtmlParser Parser;
        Dictionary<string, string> Types = new Dictionary<string, string>();

        public Form1()
        {
            InitializeComponent();
            groupBox2.Enabled = false;
            ChromeOptions chromeOptions = new ChromeOptions();
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            Driver = new ChromeDriver(service, chromeOptions);
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            Driver.Manage().Window.Minimize();
            Parser = new HtmlParser();
            JS = (IJavaScriptExecutor)Driver;
            var lines = File.ReadAllLines("Преподаватели.txt");
            foreach(var line in lines)
            {
                if(string.IsNullOrEmpty(line))
                {
                    continue;
                }
                listBox1.Items.Add(line);
                Educators.Add(new Educator() { FIO=line});
            }
        }

        WebDriver Driver;
        IJavaScriptExecutor JS;
        private void CheckCaptcha()
        {
            if(Driver.PageSource.Contains("С Вашего IP-адреса"))
            {
                Driver.Manage().Window.Size = new Size(400, 450);
            }

            Task.Run(() =>
            {
                while(true)
                {
                    if(!Driver.PageSource.Contains("С Вашего IP-адреса"))
                    {
                        break;
                    }
                    Task.Delay(1000).Wait();
                }
            }).Wait(TimeSpan.FromSeconds(30));
            if (Driver.PageSource.Contains("С Вашего IP-адреса"))
            {
                throw new Exception("Капча не распознана");
            }
            Driver.Manage().Window.Minimize();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Handler = new HttpClientHandler();
                Handler.AllowAutoRedirect = true;
                Handler.CookieContainer = new System.Net.CookieContainer();
                //Handler.Proxy = new WebProxy()
                //{
                //    Address = new Uri("http://178.218.44.79:3128"),
                //    BypassProxyOnLocal = false,
                //    UseDefaultCredentials = false,
                //};
                //Handler.ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;
                Client = new HttpClient(Handler,true);
                Parser = new HtmlParser();



                var r1 = Get("https://elibrary.ru/defaultx.asp");
                if (Auth(textBox1.Text, textBox2.Text))
                {
                    groupBox2.Enabled = true;
                    var autorPage = Get("http://www.elibrary.ru/authors.asp");
                    if (autorPage.Contains("Поиск авторов"))
                    {
                        var doc = Parser.ParseDocument(autorPage);
                        var select = doc.QuerySelector("select[id=\"rubriccode\"]");
                        foreach (var item in select.Children.Skip(1))
                        {
                            Types.Add(item.TextContent, item.GetAttribute("value"));
                        }
                    }
                    else
                    {
                        MessageBox.Show("Авторы не найдены","Ошибка");
                    }
                }
                else
                {
                    MessageBox.Show("Ошибка авторизации","Ошибка");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
            }
        }
        bool Auth(string login, string pass)
        {
            var data = new Dictionary<string, string>()
            {
                {"rpage","https://www.elibrary.ru/defaultx.asp" },
                {"login",login },
                {"password",pass },
            };

            var auth = Post("https://www.elibrary.ru/start_session.asp", new FormUrlEncodedContent(data));


            if (auth.Contains(login))
            {
                return true;
            }
            else
            {
                if(auth.Contains("С Вашего IP-адреса пуступает необычно много запросов"))
                {
                    MessageBox.Show("При авторизации возникла капча","Капча");
                }
                return false;
            }
        }
        bool AuthSelenium(string login, string pass)
        {
            Driver.Navigate().GoToUrl("https://www.elibrary.ru/defaultx.asp");
            CheckCaptcha();
            var loginInput = Driver.FindElement(By.Id("login"));
            var passInput = Driver.FindElement(By.Id("password"));
            var btn = Driver.FindElement(By.ClassName("butred"));
            if(loginInput != null && passInput!=null && btn!=null)
            {
                loginInput.SendKeys(login);
                passInput.SendKeys(pass);
                btn.Click();

                CheckCaptcha();

                return Driver.PageSource.Contains("Закрыть сессию");
            }
            else
            {
                return false;
            }
        }
        String Post(String url, HttpContent data)
        {
            var req = Client.PostAsync(url, data).Result;
            var res = req.Content.ReadAsStringAsync().Result;

            if(res.Contains("С Вашего IP-адреса пуступает необычно много запросов "))
            {
                throw new Exception("Во время запроса возникла капча");
            }

            return res;
        }
        String Get(String url)
        {
            var req = Client.GetAsync(url).Result;
            var res = req.Content.ReadAsStringAsync().Result;
            if (res.Contains("С Вашего IP-адреса пуступает необычно много запросов "))
            {
                throw new Exception("Во время запроса возникла капча");
            }
            return res;
        }

        List<Educator> Educators = new List<Educator>();

        private void button2_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    var data = new Dictionary<string, string>
            //    {
            //        {"authors_all","" },
            //        {"pagenum","" },
            //        {"authorbox_name","" },
            //        {"selid","" },
            //        {"orgid","548" },
            //        {"orgadminid","" },
            //        {"surname",textBox3.Text },
            //        {"codetype","" },
            //        {"town","Новосибирск" },
            //        {"countryid","RUS" },
            //        {"orgname","Сибирский государственный университет геосистем и технологий" },
            //        {"rubriccode","" },
            //        {"metrics","1" },
            //        {"show_nopubl","on" },
            //        {"sortorder","0" },
            //        {"order","0" },
            //        {"authorboxid","0" },
            //    };

            //    var req = Post("https://elibrary.ru/authors.asp", new FormUrlEncodedContent(data));

            //    var doc = Parser.ParseDocument(req);
            //    var table = doc.QuerySelector("table[id=\"restab\"]").Children.First().Children.Skip(3).Where(x => x.TagName == "TR");

            //    if (table.Count() == 0)
            //    {
            //        MessageBox.Show("Не найден данный автор", "Ошибка");
            //    }
            //    else
            //    {
            //        if (table.Count() > 1)
            //        {
            //            MessageBox.Show("Найдено несколько авторов по данному запросу", "Ошибка");
            //        }
            //        else
            //        {
            //            Educator.Id = Regex.Match(table.First().GetAttribute("id"), @"(\d+)").Groups[1].Value;
            //            Educator.FIO = table.First().Children[2].FirstChild.TextContent.Trim().Replace("  ", " ");
            //            Educator.LinkOnPublish = $"https://www.elibrary.ru/author_items.asp?authorid={Educator.Id}&pubrole=100&show_refs=1&show_option=0";

            //            if (textBox3.Text == Educator.FIO)
            //            {
            //                GetLinks(Educator.LinkOnPublish);
            //                if (Educator.Published.Count > 0)
            //                {
            //                    if (CreateExcel())
            //                    {
            //                        MessageBox.Show("Файл сохранен", "Выполнено");
            //                    }
            //                    else
            //                    {
            //                        MessageBox.Show("Файл не сохранен", "Ошибка");
            //                    }
            //                }
            //            }
            //            else
            //            {
            //                MessageBox.Show($"Расхождение в фио. На сайте нашли '{Educator.FIO}'", "Ошибка");
            //            }
            //        }
            //    }
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Ошибка");
            //}
        }

        private void CreateExcel()
        {
            var publishByAutor = Educators.OrderBy(x => x.FIO);
            foreach(var publish in publishByAutor.Where(x=>x.Published.Count>0))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                DirectoryInfo dirInfo = null;
                if(!Directory.Exists("Выгрузки"))
                {
                    dirInfo = Directory.CreateDirectory("Выгрузки");
                }    
                else
                {
                    dirInfo = new DirectoryInfo("Выгрузки");
                }
                if (File.Exists($"Выгрузки\\{publish.FIO}.xlsx"))
                {
                    File.Delete($"Выгрузки\\{publish.FIO}.xlsx");
                }
                using (ExcelPackage excel = new ExcelPackage($"Выгрузки\\{publish.FIO}.xlsx"))
                {
                    excel.Workbook.Worksheets.Add("Выгрузка");
                    var sheet = excel.Workbook.Worksheets[0];
                    sheet.Cells[1, 1].Value = "ФИО";
                    sheet.Cells[1, 2].Value = "Название статьи";
                    sheet.Cells[1, 3].Value = "Ссылка на статью";
                    sheet.Cells[1, 4].Value = "Анотация";
                    sheet.Cells[1, 5].Value = "Цитирование";
                    sheet.Cells[1, 6].Value = "Журнал";
                    sheet.Cells[1, 7].Value = "Авторы";
                    sheet.Cells[1, 8].Value = "Входит в РИНЦ";
                    sheet.Cells[1, 9].Value = "Входит в ядро РИНЦ";
                    sheet.Cells[1, 10].Value = "Рецензии";
                    sheet.Cells[1, 11].Value = "Ссылка для цитирования";

                    for (int i = 0; i < publish.Published.Count; i++)
                    {
                        sheet.Cells[i + 2, 1].Value = publish.FIO;
                        sheet.Cells[i + 2, 2].Value = publish.Published[i].Name;
                        sheet.Cells[i + 2, 3].Value = publish.Published[i].Link;
                        sheet.Cells[i + 2, 4].Value = publish.Published[i].Anotation;
                        sheet.Cells[i + 2, 5].Value = publish.Published[i].Citat;
                        sheet.Cells[i + 2, 6].Value = publish.Published[i].Journal;
                        sheet.Cells[i + 2, 7].Value = publish.Published[i].Autors;
                        sheet.Cells[i + 2, 8].Value = publish.Published[i].RINC;
                        sheet.Cells[i + 2, 9].Value = publish.Published[i].Core_RINC;
                        sheet.Cells[i + 2, 10].Value = publish.Published[i].Recenz;
                        sheet.Cells[i + 2, 11].Value = publish.Published[i].Citir;
                    }

                    sheet.Cells[1, 1, publish.Published.Count, 8].AutoFitColumns(70,70);


                    printSheet2(excel,publish.Published);
                    printSheet3(excel,publish.Published);

                    excel.Save();
                }
                Process.Start("explorer.exe", dirInfo.FullName);
            }
        }

        private void printSheet2(ExcelPackage excel,List<Publish>publish)
        {
            excel.Workbook.Worksheets.Add("Приложение №2");
            var sheet = excel.Workbook.Worksheets[1];
            sheet.Cells[1, 1].Value = "№ п/п";
            sheet.Cells[1, 2].Value = "Наименование работы";
            sheet.Cells[1, 3].Value = "Вид работы";
            sheet.Cells[1, 4].Value = "Выходные данные";
            sheet.Cells[1, 5].Value = "Фамилии и инициалы соавторов";
            sheet.Cells[1, 6].Value = "Аннотация";
            sheet.Cells[1, 7].Value = "Объем работы";

            for (int i = 0; i < publish.Count; i++)
            {
                sheet.Cells[i + 2, 1].Value = i+1;
                sheet.Cells[i + 2, 2].Value = publish[i].Name;
                sheet.Cells[i + 2, 3].Value = "Научная статья";
                sheet.Cells[i + 2, 4].Value = publish[i].Citir;
                sheet.Cells[i + 2, 5].Value = publish[i].Autors;
                sheet.Cells[i + 2, 6].Value = publish[i].Anotation;
                sheet.Cells[i + 2, 7].Value = publish[i].CountPage;
            }
            sheet.Cells[1, 1, 1, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            sheet.Cells[1, 1, publish.Count, 6].AutoFitColumns(70,70);
        }
        private void printSheet3(ExcelPackage excel, List<Publish> publish)
        {
            excel.Workbook.Worksheets.Add("Список трудов претендента");
            var sheet = excel.Workbook.Worksheets[2];
            sheet.Cells[1, 1].Value = "№ п/п";
            sheet.Cells[1, 2].Value = "Наименование учебных изданий, научных трудов и патентов на изобретения и иные объекты интеллектуальной собственности";
            sheet.Cells[1, 3].Value = "Форма учебных изданий и научных трудов";
            sheet.Cells[1, 4].Value = "Выходные данные";
            sheet.Cells[1, 5].Value = "Объем";
            sheet.Cells[1, 6].Value = "Соавторы";

            for (int i = 0; i < publish.Count; i++)
            {
                sheet.Cells[i + 2, 1].Value = i + 1;
                sheet.Cells[i + 2, 2].Value = publish[i].Name;
                sheet.Cells[i + 2, 3].Value = "Научная статья";
                sheet.Cells[i + 2, 4].Value = publish[i].Citir;
                sheet.Cells[i + 2, 5].Value = publish[i].CountPage;
                sheet.Cells[i + 2, 6].Value = publish[i].Autors;
            }
            sheet.Cells[1, 1, 1, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            sheet.Cells[1, 1, publish.Count, 6].AutoFitColumns(70, 70);
        }

        private void GetLinks(Educator educator)
        {
            Driver.Navigate().GoToUrl(educator.LinkOnPublish);
            //var page = Get(educator.LinkOnPublish);
            var doc = Parser.ParseDocument(Driver.PageSource);

            var countPage = doc.QuerySelector("#pages > table > tbody > tr")?.Children.Where(x=>x.GetAttribute("bgcolor")!=null).ToList();
            int iterPage = 0;
            do
            {
                iterPage++;
                Task.Delay(1000).Wait();
                doc = Parser.ParseDocument(Driver.PageSource);
                var table = doc.QuerySelector("table[id=\"restab\"]").Children[0].Children.Skip(3);
                foreach (var item in table)
                {
                    try
                    {
                        var id = Regex.Match(item.GetAttribute("id"), @"(\d+)").Groups[1].Value;
                        string name = "";
                        string link = "";
                        if (item.Children[1].Children[0].Children[0].Children.Count() > 0)
                        {
                            name = item.Children[1].Children[0].Children[0].Children[0].TextContent;
                            link = $"https://www.elibrary.ru/{item.Children[1].Children[0].GetAttribute("href")}";
                        }
                        else
                        {
                            continue;
                            //name = item.Children[1].Children[0].Children[0].TextContent;
                            //link = $"https://www.elibrary.ru/{item.Children[1].Children[0].GetAttribute("href")}";
                        }
                        educator.Published.Add(new Publish { Id = id, Name = name, Link = link });
                    }
                    catch (Exception ex)
                    {

                    }
                }
                if(countPage==null)
                {
                    break;
                }
                if(iterPage!=countPage.Count())
                {
                    JS.ExecuteScript(countPage[iterPage].Children[0].GetAttribute("href"));
                }

            } while (iterPage != countPage.Count());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(AuthSelenium(textBox1.Text,textBox2.Text))
            {
                button8.Enabled = true;
                Driver.Navigate().GoToUrl("https://elibrary.ru/authors.asp");
                Task.Delay(3000).Wait();
                var btnClear = Driver.FindElement(By.XPath("//*[@id=\"show_param\"]/table[6]/tbody/tr[2]/td[5]/div"));
                if(btnClear != null)
                {
                    btnClear.Click();
                }
                SetTematics();
                comboBox1.SelectedIndex = 0;
                groupBox2.Enabled = true;
            }
        }

        private List<Tematic>Tematics = new List<Tematic>();

        private void SetTematics()
        {
            if(Tematics.Count==0)
            {
                var doc = Parser.ParseDocument(Driver.PageSource);
                var tematics = doc.QuerySelector("select[id=\"rubriccode\"]");
                foreach (var tematic in tematics.Children)
                {
                    var t = new Tematic() { Text = tematic.TextContent.Trim(), Value = tematic.GetAttribute("value") };
                    Tematics.Add(t);
                    comboBox1.Items.Add(t);
                }
            }
            
        }

        //private void GetResultAutors(Educator educator)
        //{
        //    var cookies = Driver.Manage().Cookies.AllCookies;
        //    Handler = new HttpClientHandler();
        //    Handler.AllowAutoRedirect = true;
        //    Handler.CookieContainer = new CookieContainer();
        //    foreach(var cookie in cookies)
        //    {
        //        var c = new System.Net.Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain);
        //        Handler.CookieContainer.Add(c);
        //    }
        //    Client = new HttpClient(Handler);

        //    string idTematic = "";
        //    if (SelectedTematic != null)
        //    {
        //        idTematic = SelectedTematic.Value;
        //    }

        //    var data = new Dictionary<string, string>
        //        {
        //            {"authors_all","" },
        //            {"pagenum","" },
        //            {"authorbox_name","" },
        //            {"selid","" },
        //            {"orgid","548" },
        //            {"orgadminid","" },
        //            {"surname",educator.FIO },
        //            {"codetype","SPIN" },//было пусто
        //            {"codevalue","" },//не было
        //            {"town","Новосибирск" },
        //            {"countryid","RUS" },
        //            {"orgname","Сибирский государственный университет геосистем и технологий" },
        //            {"rubriccode",idTematic },
        //            {"metrics","1" },
        //            {"show_nopubl","on" },
        //            {"sortorder","0" },
        //            {"order","0" },
        //            {"authorboxid","0" },
        //        };
        //    data.Add("hid280970", "Бугакова Т Ю");

        //    var req = Post("https://elibrary.ru/authors.asp", new FormUrlEncodedContent(data));

        //    var doc = Parser.ParseDocument(req);
        //    if(req.Contains("Не найдено авторов, удовлетворяющих условиям поиска"))
        //    {
        //        MessageBox.Show("Не найдено авторов, удовлетворяющих условиям поиска", "Ошибка");
        //        IsFind = false;
        //        return;
        //    }
        //    IsFind = true;
        //    var table = doc.QuerySelector("table[id=\"restab\"]").Children.First().Children.Skip(3).Where(x => x.TagName == "TR");

        //    educator.Id = Regex.Match(table.First().GetAttribute("id"), @"(\d+)").Groups[1].Value;
        //    educator.FIO = table.First().Children[2].FirstChild.TextContent.Trim().Replace("  ", " ");
        //    educator.LinkOnPublish = $"https://www.elibrary.ru/author_items.asp?authorid={educator.Id}&pubrole=100&show_refs=1&show_option=0";
        //}


        private void SetCity()
        {
            var city = Driver.FindElement(By.Name("town"));
            var towns = city.FindElements(By.TagName("option"));
            towns.First(x => x.Text.Contains("Новосибирск")).Click();
        }
        private void SetCountry()
        {
            var country = Driver.FindElement(By.Name("countryid"));
            var values = country.FindElements(By.TagName("option"));
            values.First(x => x.Text.Contains("Россия")).Click();
        }
        private void SetOrganisation()
        {
            Driver.FindElement(By.XPath("//*[@id=\"show_param\"]/table[3]/tbody/tr[2]/td[2]/div")).Click();
            var s = Driver.ExecuteScript("javascript:organizat_add(548,&quot;Сибирский государственный университет геосистем и технологий&quot;)");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count>0)
            {
                CreateExcel();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(listBox1.SelectedItems.Count > 0)
            {
                if(!listBox2.Items.Contains(listBox1.SelectedItems[0]))
                {
                    listBox2.Items.Add(listBox1.SelectedItems[0]);
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if(listBox2.SelectedItems.Count>0)
            {
                listBox2.Items.Remove(listBox2.SelectedItems[0]);
            }
        }
        //private bool IsFind = false;

        public void SetEducator(IElement rowAuthor, Educator educator)
        {
            var doc = Parser.ParseDocument(rowAuthor.OuterHtml);
            educator.Id = Regex.Match(doc.Body.Children[2].Children[0].Id, @"(\d+)").Groups[1].Value;
            educator.FIO = doc.Body.Children[3].TextContent.Trim();
            educator.LinkOnPublish = $"https://www.elibrary.ru/author_items.asp?authorid={educator.Id}&pubrole=100&show_refs=1&show_option=0";
        }
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (var autor in listBox2.Items)
                {
                    Driver.Navigate().GoToUrl("https://elibrary.ru/authors.asp");
                    Driver.FindElement(By.Id("surname")).Clear();
                    Driver.FindElement(By.Id("surname")).SendKeys(autor.ToString());
                    var currentAutor = Educators.FirstOrDefault(x => x.FIO == autor);
                    if(currentAutor == null)
                    {
                        currentAutor = Educators.First(x => x.FIO.ToLower().Contains(autor.ToString().ToLower().Split(' ')[0]));
                    }
                    SelectElement selected = new SelectElement(Driver.FindElement(By.Id("rubriccode")));
                    if(SelectedTematic!=null)
                    {
                        selected.SelectByValue(SelectedTematic.Value);
                    }

                    Driver.FindElement(By.ClassName("butred")).Click();
                    Task.Delay(1000).Wait();

                    var pageAllAutors = Parser.ParseDocument(Driver.PageSource);

                    var countAuthorsQ = pageAllAutors.QuerySelector("body > table > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(2) > table > tbody > tr:nth-child(3) > td:nth-child(1) > table > tbody > tr > td > div:nth-child(3) > table > tbody > tr > td.redref");
                    if(countAuthorsQ.TextContent.Contains("Не найдено авторов, удовлетворяющих условиям поиска"))
                    {
                        MessageBox.Show("Указанный автор не найден","Ошибка");
                        return;
                    }    
                    var countAuthors = Convert.ToInt32(Regex.Match(countAuthorsQ.TextContent, @"Всего найдено авторов: (\d+) из").Groups[1].Value);
                    if(countAuthors>1)
                    {
                        var allAutors = pageAllAutors.QuerySelectorAll("tr").Where(x=>x.OuterHtml.Contains(currentAutor.FIO)).Skip(5);

                        FormAuthors formAuthors = new FormAuthors(allAutors);
                        if(formAuthors.ShowDialog()==DialogResult.OK)
                        {
                            currentAutor = formAuthors.SelectedEducator;
                            Educators.Add(currentAutor);
                        }
                    }
                    else
                    {
                        var allAutors = pageAllAutors.QuerySelector("#restab > tbody");
                        SetEducator(allAutors.Children[3], currentAutor);
                    }



                    GetLinks(currentAutor);
                    int value = 0;
                    if (currentAutor.Published.Count > 0)
                    {
                        progressBar1.Maximum = currentAutor.Published.Count;
                        foreach (var publish in currentAutor.Published)
                        {
                            try
                            {
                                value++;
                                progressBar1.Value = value;
                                if (publish.Link == "https://www.elibrary.ru/")
                                {
                                    continue;
                                }
                                Driver.Navigate().GoToUrl(publish.Link);
                                CheckCaptcha();
                                var doc = Parser.ParseDocument(Driver.PageSource);
                                var journal = doc.QuerySelector("body > table > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(1) > div:nth-child(4) > table:nth-child(6) > tbody > tr:nth-child(2) > td:nth-child(2)");
                                publish.Journal = journal?.TextContent.Trim();

                                var citat = doc.QuerySelectorAll("td");
                                var citat2 = citat.Where(x => x.TextContent.Contains("Цитирований в РИНЦ"));
                                var citat3 = citat2.Children("font");
                                if (citat3.Count() == 0)
                                {
                                    citat3 = citat2.Children("a");
                                }
                                publish.Citat = citat3.First().TextContent;

                                var authors = doc.QuerySelectorAll("body > table > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(1) > div:nth-child(4) > table:nth-child(2) > tbody > tr > td:nth-child(2)");
                                var authorsChild = authors.Children("div");
                                if (authorsChild.Count() > 1)
                                {
                                    string str = "";
                                    foreach (var child in authorsChild)
                                    {
                                        if (child.Children.Count() > 0)
                                        {
                                            if (child.Children[0].TextContent.ToLower().Contains(currentAutor.FIO.Split(' ')[0].ToLower()))
                                            {
                                                continue;
                                            }
                                            str += child.Children[0].TextContent;
                                        }
                                    }
                                    publish.Autors = str;
                                }
                                var anotation = doc.QuerySelector("div[id=\"abstract1\"]");
                                publish.Anotation = anotation?.TextContent.Trim();

                                var rinc = doc.QuerySelectorAll("td");
                                var rinc2 = rinc.Where(x => x.TextContent.Contains("Входит в РИНЦ"));
                                var rinc3 = rinc2.Children("font");
                                publish.RINC = rinc3.First()?.TextContent;

                                var core_rinc = doc.QuerySelector("#InCoreRisc");
                                publish.Core_RINC = core_rinc?.TextContent.Trim();

                                var recenz = doc.QuerySelectorAll("td");
                                var recenz2 = recenz.Where(x => x.TextContent.Contains("Рецензии"));
                                var recenz3 = recenz2.Children("font");
                                if (recenz3.Any())
                                {
                                    publish.Recenz = recenz3.First().TextContent;
                                }

                                var getCountPage = new Func<string,int>((html) =>
                                {
                                    if (Regex.IsMatch(html, @"Страницы:.?\d+-\d+.*?"))
                                    {
                                        var r = Regex.Match(html, @"Страницы:.?(\d+)-(\d+).*?");
                                        return Convert.ToInt32(r.Groups[2].Value) - Convert.ToInt32(r.Groups[1].Value) + 1;
                                    }
                                    else
                                    {
                                        if (Regex.IsMatch(html, @"Число.страниц:.?\d+"))
                                        {
                                            return Convert.ToInt32(Regex.Match(html, @"Число.страниц:.?(\d+)").Groups[1].Value);
                                        }
                                        else
                                        {
                                            if(Regex.IsMatch(html, @"Страницы:.?\d+"))
                                            {
                                                return Convert.ToInt32(Regex.Match(html, @"Страницы:.?(\d+)").Groups[1].Value);
                                            }
                                            else
                                            {

                                            }
                                        }
                                    }
                                    
                                    
                                    return 1;
                                });

                                publish.CountPage = getCountPage(doc.Body.TextContent);

                                var cit = JS.ExecuteScript("javascript:for_reference()");
                                var forms = Driver.SwitchTo().Frame(Driver.FindElement(By.Id("fancybox-frame")));
                                var citElement = Driver.FindElement(By.Id("ref"));
                                var citText = citElement.Text;
                                publish.Citir = citText;
                            }
                            catch(Exception ex)
                            {

                            }
                        }
                    }
                }
                progressBar1.Value = 0;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"Ошибка");
            }
            finally
            {
                printDataTable();
            }
        }


        public void LogOut()
        {
            var end = JS.ExecuteScript("javascript:end_session()");
            var btn = Driver.FindElement(By.ClassName("butred"));
            if(btn!=null)
            {
                MessageBox.Show("Выполнен выход из учетной записи","Сессия завершена");
            }
        }


        private void printDataTable()
        {
            foreach(var educator in Educators)
            {
                foreach(var publish in educator.Published)
                {
                    dataGridView1.Rows.Add(publish.Id, educator.FIO, publish.Name, publish.Link, publish.Anotation, publish.Citat, publish.Journal, publish.Autors, publish.Status, publish.RINC, publish.Core_RINC, publish.Recenz, publish.Citir);
                }
            }    
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count > 0)
            {
                if (!listBox2.Items.Contains(listBox1.SelectedItems[0]))
                {
                    listBox2.Items.Add(listBox1.SelectedItems[0]);
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox3.Text))
            {
                var arr = new string[] {  $"\r\n{textBox3.Text}" };
                File.AppendAllLines("Преподаватели.txt", arr);
                listBox1.Items.Add(textBox3.Text);
                Educators.Add(new Educator() { FIO = textBox3.Text });
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if(listBox1.SelectedItem!=null)
            {
                listBox1.Items.Remove(listBox1.SelectedItem);
                string str = "";
                foreach(var item in listBox1.Items)
                {
                    str+= item.ToString() + "\r\n";
                }
                str.Trim();
                File.WriteAllText("Преподаватели.txt", str);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            LogOut();
        }
        private Tematic SelectedTematic = null;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedTematic = (sender as System.Windows.Forms.ComboBox).SelectedItem as Tematic;
        }
    }
}
