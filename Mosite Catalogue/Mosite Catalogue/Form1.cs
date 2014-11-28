using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support;
using Microsoft.Office.Interop.Excel;

namespace Mosite_Catalogue
{
    public partial class Form1 : Form
    {
        public int num;
        public int catNum;
        public int prodNum;
        public int entryNum;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            num = 2;
            catNum = 1;
            prodNum = 1;
            entryNum = 1;
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                String OutputFileName = textBox2.Text + " on " + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + " at " + DateTime.Now.Hour + "h " + DateTime.Now.Minute + "m " + DateTime.Now.Second + "s";
                string excel_filename = @"C:\catalogue\" + OutputFileName + ".xlsx";
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                Workbook wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wb.Sheets)
                {
                    if (sh.Name == "Sheet1")
                    {
                        //do something
                        sh.Cells[1, "A"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        sh.Cells[1, "A"].VAlue2 = "<No.>";
                        sh.Cells[1, "B"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        sh.Cells[1, "B"].VAlue2 = "Name";
                        sh.Cells[1, "C"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        sh.Cells[1, "C"].VAlue2 = "Count";
                        sh.Cells[1, "D"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        sh.Cells[1, "D"].VAlue2 = "Href";

                        Microsoft.Office.Interop.Excel.Range A = sh.get_Range("A:A", System.Type.Missing);
                        A.EntireColumn.ColumnWidth = 6;
                        A.EntireColumn.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                        Microsoft.Office.Interop.Excel.Range B = sh.get_Range("B:B", System.Type.Missing);
                        B.EntireColumn.ColumnWidth = 60;
                        B.EntireColumn.WrapText = true;

                        Microsoft.Office.Interop.Excel.Range C = sh.get_Range("C:C", System.Type.Missing);
                        C.EntireColumn.ColumnWidth = 10;
                        C.EntireColumn.WrapText = true;

                        Microsoft.Office.Interop.Excel.Range D = sh.get_Range("D:D", System.Type.Missing);
                        D.EntireColumn.ColumnWidth = 100;
                        D.EntireColumn.WrapText = true;

                        Microsoft.Office.Interop.Excel.Range E = sh.get_Range("E:E", System.Type.Missing);
                        E.EntireColumn.ColumnWidth = 60;
                        E.EntireColumn.WrapText = true;

                        Microsoft.Office.Interop.Excel.Range F = sh.get_Range("F:F", System.Type.Missing);
                        F.EntireColumn.WrapText = true;
                    }
                }

                wb.SaveAs(excel_filename);
                excel.Quit();

                StartReadingCatalogue(textBox1.Text, excel_filename);
            }

            else if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("Please give the mosite url and merchant name");
            }
        }

        public void StartReadingCatalogue(string url, string excel_filename)
        {
            //IWebDriver driver = new FirefoxDriver();
            //driver.Manage().Cookies.DeleteAllCookies();
            //driver.Navigate().GoToUrl(url);
            string PageUrl = url;
            ExtractPageData(PageUrl, excel_filename);
        }

        public int ExtractPageData(string PageUrl, string excel_filename)
        {
            try
            {
                IWebDriver driver = new FirefoxDriver();
                driver.Navigate().GoToUrl(PageUrl);

                // CATEGORY / SUBCATEGORY
                //Checks the presence of category/subacategory elements
                if (IsElementDisplayed(driver, By.ClassName("wrappable")))
                {

                    IList<IWebElement> PageElements = driver.FindElements(By.ClassName("wrappable"));
                    String[] PageElementNames = new String[PageElements.Count];
                    String[] PageElementHref = new String[PageElements.Count];
                    string PageData = "";
                    int i = 0;
                    int j = 0;
                    foreach (IWebElement element in PageElements)
                    {
                        PageData = PageData + " - " + element.Text + " --> " + driver.FindElement(By.LinkText(element.Text)).GetAttribute("href") + element.GetCssValue("a") + "\n";
                        PageElementNames[i++] = element.Text;
                        PageElementHref[j++] = driver.FindElement(By.LinkText(element.Text)).GetAttribute("href") + "";
                        string href = driver.FindElement(By.LinkText(element.Text)).GetAttribute("href") + "";
                        AddRowCategory(element.Text, driver.FindElement(By.LinkText(element.Text)).GetAttribute("href") + "", num, excel_filename, PageElements.Count);
                        num++;
                        ExtractPageData(href, excel_filename);
                    }
                    num++;
                    //MessageBox.Show(PageData);
                    driver.Close();
                    return PageElements.Count;
                }

                // PRODUCT LIST PAGE
                //Checks for the presence of product list element
                else if (IsElementDisplayed(driver, By.ClassName("responsive-list")))
                {
                    IList<IWebElement> ProductPageElements = driver.FindElements(By.ClassName("responsive-list"));
                    String[] ProductPageElementNames = new String[ProductPageElements.Count];
                    String[] ProductPageElementHref = new String[ProductPageElements.Count];
                    string ProductPageData = "";

                    //Here we are capturing the products available on the page individually as a list.
                    IList<IWebElement> ProductsOnPage = driver.FindElements(By.CssSelector("div[class='ui-btn-text']"));

                    int i = 0;
                    int j = 0;
                    foreach (IWebElement element in ProductPageElements)
                    {
                        ProductPageData = ProductPageData + " - " + element.Text + "\n";
                        ProductPageElementNames[i++] = element.Text;
                        AddRowProducts(element.Text, "", num, excel_filename, ProductsOnPage.Count);
                        num++;
                    }
                    num++;
                    //MessageBox.Show(ProductPageData);

                    //In case of extended product pages

                    driver.Close();
                }

                else 
                {
                    AddRowError("Neither categories nor products found on page. Please check internet connectiivity" , excel_filename, num);
                    num++;
                    ExtractPageData(PageUrl, excel_filename);
                }
            }

            catch (Exception ex)
            {
                //MessageBox.Show("Exception Handled:\n" + ex);
                AddRowError(ex.ToString(), excel_filename,num);
                num++;
                ExtractPageData(PageUrl, excel_filename);
            }
            return 0;
        }

        public void AddRowCategory(string ElementName, string href, int row, string excel_filename,int count)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(excel_filename);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wb.Sheets)
            {
                if (sh.Name == "Sheet1")
                {
                    //do something
                    sh.Cells[row, "A"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    sh.Cells[row, "A"].Value2 = entryNum + "  " + "C" + catNum;
                    entryNum++; catNum++;
                    sh.Cells[row, "B"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    sh.Cells[row, "B"].Value2 = ElementName;
                    sh.Cells[row, "D"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    sh.Cells[row, "D"].Value2 = href;
                    sh.Cells[row, "C"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    
                    //Avoided to get the number of categories present on category page.[For the number getting repeated on each category row]
                    //sh.Cells[row, "C"].Value2 = count;
                }
            }
            wb.Save();
            excel.Quit();
        }

        public void AddRowProducts(string ElementName, string href, int row, string excel_filename, int count)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(excel_filename);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wb.Sheets)
            {
                if (sh.Name == "Sheet1")
                {
                    //do something
                    sh.Cells[row, "A"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    sh.Cells[row, "A"].Value2 = entryNum + "  " + "P" + prodNum;
                    entryNum++; prodNum++;
                    sh.Cells[row, "B"].Value2 = ElementName;
                    sh.Cells[row, "C"].Value2 = count;
                }
            }
            wb.Save();
            excel.Quit();
        }

        public void AddRowError(string Error,string excel_filename,int row)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(excel_filename);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wb.Sheets)
            {
                if (sh.Name == "Sheet1")
                {
                    //do something
                    sh.Cells[row, "E"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                    sh.Cells[row, "E"].Value2 = Error;

                    sh.Cells[row, "F"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    sh.Cells[row, "F"].Value2 = DateTime.Now.ToShortTimeString();
                }
            }
            wb.Save();
            excel.Quit();
        }

        //Element Present and Displayed
        public bool IsElementDisplayed(IWebDriver driver, By element)
        {
            if (driver.FindElements(element).Count > 0)
            {
                if (driver.FindElement(element).Displayed)
                    return true;
                else
                    return false;
            }
            else
            {
                return false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
