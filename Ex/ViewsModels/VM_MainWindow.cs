using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Ex.Infrastructure;
using Ex.Models;
using Ex.Views;
using Ex.ViewsModels;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;

namespace Ex.ViewsModels
{
    class VM_MainWindow:OnPropertyChangedClass
    {
        sam1Entities db = new sam1Entities();
        Order _curren_order;
        public Order curren_order { get => _curren_order; set 
            {

                SetProperty(ref _curren_order, value);
                    } }
        Customer _curren_customers = new Customer();
        public Customer curren_customers { get => _curren_customers; set => SetProperty(ref _curren_customers, value); }
        ObservableCollection<Order> _orders = new ObservableCollection<Order>();
        public ObservableCollection<Order> orders { get => new ObservableCollection<Order> (curren_customers.Orders);}
        ObservableCollection<Customer> _customers = new ObservableCollection<Customer>();
        public ObservableCollection<Customer> customers { get => new ObservableCollection<Customer>(db.Customers);}
        ObservableCollection<OrderItem> _orderItem=new ObservableCollection<OrderItem>();
        public ObservableCollection<OrderItem> orderItem { get
            {
                return _orderItem;
            }
            set { SetProperty(ref _orderItem, value); }
        }
        RelayCommand _excel;
        public RelayCommand excel => _excel ?? (_excel = new RelayCommand(
            p=>
            {
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

                ExcelApp.Application.Workbooks.Add(Type.Missing);
                ExcelApp.Columns.ColumnWidth = 15;
                ExcelApp.Cells[1, 1] = $"Заказ №{curren_order.order_num}";
                ExcelApp.Cells[2, 1] = $"Дата заказа";
                ExcelApp.Cells[2, 3] = $"{curren_order.order_date}";
                ExcelApp.Cells[3, 1] = $"Заказчик";
                ExcelApp.Cells[3, 3] = $"{curren_order.Customer.cust_name}";
                ExcelApp.Cells[4, 1] = $"Адрес заказа";
                ExcelApp.Cells[4, 3] = $"{curren_order.Customer.cust_address}";

                ExcelApp.Cells[6, 1] = $"Название продукта";
                ExcelApp.Cells[6, 2] = $"Количество";
                ExcelApp.Cells[6, 3] = $"Цена";
                ExcelApp.Cells[6, 4] = $"Стоимость";
                int i = 7;
                foreach (OrderItem item in curren_order.OrderItems)
                {
                    ExcelApp.Cells[i, 1] = $"{item.Product.prod_name}";
                    ExcelApp.Cells[i, 2] = $"{item.quantity}";
                    ExcelApp.Cells[i, 3] = $"{item.Product.prod_price}";
                    ExcelApp.Cells[i, 4] = $"{item.sum_item}";
                    i++;
                }
                ExcelApp.Cells[i, 1] = $"Стоимость";
                ExcelApp.Cells[i, 2] = $"{curren_order.sum_order}";
                Microsoft.Office.Interop.Excel.Range _excelCells3 = (Microsoft.Office.Interop.Excel.Range)ExcelApp.get_Range("A5", "D5").Cells;
                _excelCells3.Interior.Color = 500  ;

                Microsoft.Office.Interop.Excel.Range _excelCells4 = (Microsoft.Office.Interop.Excel.Range)ExcelApp.get_Range("A5", $"D{i+1}").Cells;
                _excelCells4.Borders.LineStyle = 1;
                ExcelApp.Visible = true; 
            },
            p =>
            (curren_order != null)
            ));
        RelayCommand _pdf;
        public RelayCommand pdf => _pdf ?? (_pdf = new RelayCommand(
            p =>
            {
                
                iTextSharp.text.Document doc = new iTextSharp.text.Document();

                PdfWriter.GetInstance(doc, new FileStream("pdfTables.pdf", FileMode.Create));

                doc.Open();

                BaseFont baseFont = BaseFont.CreateFont("C:/Windows/Fonts/arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
                PdfPTable table = new PdfPTable(4);

                PdfPCell cell = new PdfPCell(new Phrase("Заказ № " + curren_order.cust_id, font));

                cell.Colspan = 4;
                cell.HorizontalAlignment = 1;
                cell.Border = 0;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase("Дата заказа " + curren_order.order_date, font));
                cell.Colspan = 4;
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase("Заказчик " + curren_customers.cust_name, font));
                cell.Colspan = 4;
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase("Адрес заказа " + curren_customers.cust_address, font));
                cell.Colspan = 4;
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase(new Phrase($"Название продукта", font)));
                cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase($"Количество", font)));
                cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase($"Цена", font)));
                cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase($"Стоимость", font)));
                cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
                foreach (OrderItem item in curren_order.OrderItems)
                {
                    cell = new PdfPCell(new Phrase(new Phrase($"{item.Product.prod_name}", font)));
                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase($"{item.quantity}", font)));
                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase($"{item.Product.prod_price}", font)));
                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase($"{item.sum_item}", font)));
                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }
                cell = new PdfPCell(new Phrase("Итого " + curren_order.sum_order, font));
                cell.Colspan = 4;
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);
                doc.Add(table);

                doc.Close();
            },
            p=>
            (curren_order!=null)
            ));
        RelayCommand _json;
        public RelayCommand json => _json ?? (_json = new RelayCommand(
            p =>
            {
                List < Product > products = new List<Product>();
                foreach (OrderItem item in curren_order.OrderItems)
                {
                    products.Add(new Product()
                    {
                        prod_id=item.Product.prod_id,
                        prod_desc=item.Product.prod_desc,
                        prod_name=item.Product.prod_name,
                        prod_price=item.Product.prod_price
                    });
                }
                    string json = JsonConvert.SerializeObject(products);
                File.WriteAllText(@"Products.json", json);
            },
            p=>
            (curren_order!=null)
            ));


    }
}
