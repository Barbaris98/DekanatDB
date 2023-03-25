using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DekanatDB
{
    public partial class FacultyInfo : Form
    {
        readonly MainForm mainForm;

        public FacultyInfo(MainForm owner)
        {
            mainForm = owner;
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                if (mainForm.dataGridView2.SelectedRows.Count > 0)
                {
                    int id = int.Parse(mainForm.dataGridView2.CurrentRow.Cells[0].Value.ToString());

                    //по соответствующему id
                    Faculty faculty = db.Facultys.Include(s => s.Students).FirstOrDefault(x => x.Id == id);

                    //экспорт
                    var path = Path.Combine(Environment.CurrentDirectory, "Export", "export_selected_students.xlsx");

                    var selectedStudents = faculty.Students.ToList();
                    XLWorkbook workBook = new XLWorkbook();
                    
                    var sheet = workBook.Worksheets.Add("Students");
                    var startRow = 2;
                    int startCol = 1;

                    //делаем "шапку таблицы"
                    sheet.Cell("A1").Value = "Id";
                    sheet.Cell("B1").Value = "Фамилия";
                    sheet.Cell("C1").Value = "Имя";
                    sheet.Cell("D1").Value = "Отчество";
                    sheet.Cell("E1").Value = "№ зачётки";
                    sheet.Cell("F1").Value = "Дата рождения";
                    sheet.Cell("G1").Value = "№ Ф";

                    foreach (var item in selectedStudents)
                    {
                        sheet.Cell(startRow, startCol++).Value = item.Id;
                        sheet.Cell(startRow, startCol++).Value = item.LastName;
                        sheet.Cell(startRow, startCol++).Value = item.Name;
                        sheet.Cell(startRow, startCol++).Value = item.MiddleName;
                        sheet.Cell(startRow, startCol++).Value = item.RecordNumber;
                        sheet.Cell(startRow, startCol++).Value = item.DateOfBirth;
                        sheet.Cell(startRow, startCol).Value = item.FacultyId;

                        startCol = 1;
                        startRow++;
                    }

                    // авт. р-р ячеек по их содержимому
                    sheet.Columns(1, 7).AdjustToContents();

                    sheet.Row(1).Height = 25;

                    workBook.SaveAs(path);

                    MessageBox.Show("Отчёт сформирован!");

                }
            }
        }
    }
}
