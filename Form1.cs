using Microsoft.EntityFrameworkCore;
using System.Linq;
using System.Numerics;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.EntityFrameworkCore.ChangeTracking.Internal;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DekanatDB
{
    public partial class MainForm : Form
    {
        public MainForm()
        {

            InitializeComponent();

            //создадим элемент меню
            ToolStripMenuItem shoMore = new ToolStripMenuItem("Подробнее");
            //Добавление элемента меню
            contextMenuStrip1.Items.Add(shoMore);
            //Ассоциируем контекстное меню с гридом
            dataGridView2.ContextMenuStrip = contextMenuStrip1;
            //Установим обработчик события для меню
            shoMore.Click += shoMore_Click;
            

            using (ApplicationContext db = new ApplicationContext())
            {
                db.Students.LoadAsync();
                db.Facultys.LoadAsync();
            }
            
        }


        private void сброситьБДПоумолчаниюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Database.EnsureDeleted();

            }

            //добавление сущностей
            using (ApplicationContext db = new ApplicationContext())
            {
                
                // создание и добавление моделей
                Faculty engineering = new Faculty { NameFaculty = "М - Машиностроительный" };
                Faculty instrumentation = new Faculty { NameFaculty = "П - Приборостроительный" };
                Faculty computerScience = new Faculty { NameFaculty = "ИВТ - Информатика и вычислительная техника" };
                Faculty engineeringAndEconomic = new Faculty { NameFaculty = "ИЭ - Инженерно-экономический" };
                db.Facultys.AddRange(engineering, instrumentation, computerScience, engineeringAndEconomic);


                Student s1 = new Student
                {
                    LastName = "Александров",
                    Name = "Никита",
                    MiddleName = "Сергеевич",
                    RecordNumber = 18051351,
                    DateOfBirth = "08.31.1998",
                    Faculty = engineering
                };
                Student s2 = new Student
                {
                    LastName = "Анкин",
                    Name = "Алексей",
                    MiddleName = "Николаевич",
                    RecordNumber = 18051353,
                    DateOfBirth = "01.06.1998",
                    Faculty = engineering
                };
                Student s3 = new Student
                {
                    LastName = "Башкуртов",
                    Name = "Егор",
                    MiddleName = "Юрьевич",
                    RecordNumber = 18051354,
                    DateOfBirth = "29.05.2000",
                    Faculty = instrumentation
                };
                Student s4 = new Student
                {
                    LastName = "Беляев",
                    Name = "Антон",
                    MiddleName = "Сергеевич",
                    RecordNumber = 18051301,
                    DateOfBirth = "21.07.1999",
                    Faculty = instrumentation
                };
                Student s5 = new Student
                {
                    LastName = "Брик",
                    Name = "Валериий",
                    MiddleName = "Романович",
                    RecordNumber = 18051355,
                    DateOfBirth = "02.04.2000",
                    Faculty = instrumentation
                };
                Student s6 = new Student
                {
                    LastName = "Бурков",
                    Name = "Денис",
                    MiddleName = "Валерьевич",
                    RecordNumber = 18051479,
                    DateOfBirth = "03.04.1998",
                    Faculty = instrumentation
                };
                Student s7 = new Student
                {
                    LastName = "Буров",
                    Name = "Андрей",
                    MiddleName = "Николаевич",
                    RecordNumber = 18051358,
                    DateOfBirth = "17.12.1998",
                    Faculty = engineering
                };
                Student s8 = new Student
                {
                    LastName = "Бутолина",
                    Name = "Диана",
                    MiddleName = "Васильевна",
                    RecordNumber = 18051480,
                    DateOfBirth = "18.04.1998",
                    Faculty = computerScience
                };
                Student s9 = new Student
                {
                    LastName = "Вахрушев",
                    Name = "Артем",
                    MiddleName = "Алексеевич",
                    RecordNumber = 18051303,
                    DateOfBirth = "06.06.1999",
                    Faculty = engineering
                };
                Student s10 = new Student
                {
                    LastName = "Камашев",
                    Name = "Григорий",
                    MiddleName = "Андреевич",
                    RecordNumber = 18051400,
                    DateOfBirth = "04.04.1998",
                    Faculty = computerScience
                };
                Student s11 = new Student
                {
                    LastName = "Касимов",
                    Name = "Амир",
                    MiddleName = "Борисович",
                    RecordNumber = 18051368,
                    DateOfBirth = "01.11.2000",
                    Faculty = engineeringAndEconomic
                };
                Student s12 = new Student
                {
                    LastName = "Кузнецов",
                    Name = "Дмитрий",
                    MiddleName = "Андреевич",
                    RecordNumber = 18051372,
                    DateOfBirth = "21.08.1999",
                    Faculty = computerScience
                };
                Student s13 = new Student
                {
                    LastName = "Новиков",
                    Name = "Максим",
                    MiddleName = "Александрович",
                    RecordNumber = 18051376,
                    DateOfBirth = "18.03.1998",
                    Faculty = computerScience
                };
                Student s14 = new Student
                {
                    LastName = "Павленко",
                    Name = "Владислав",
                    MiddleName = "Александрович",
                    RecordNumber = 18051484,
                    DateOfBirth = "23.02.1998",
                    Faculty = computerScience
                };
                Student s15 = new Student
                {
                    LastName = "Рожков",
                    Name = "Тимофей",
                    MiddleName = "Михайлович",
                    RecordNumber = 18051306,
                    DateOfBirth = "20.04.1999",
                    Faculty = engineering
                };
                Student s16 = new Student
                {
                    LastName = "Санников",
                    Name = "Кирилл",
                    MiddleName = "Юрьевич",
                    RecordNumber = 18051307,
                    DateOfBirth = "18.04.1999",
                    Faculty = computerScience
                };
                Student s17 = new Student
                {
                    LastName = "Фролов",
                    Name = "Егор",
                    MiddleName = "Вадимовиич",
                    RecordNumber = 18051384,
                    DateOfBirth = "15.07.1999",
                    Faculty = engineeringAndEconomic
                };
                Student s18 = new Student
                {
                    LastName = "Яркеева",
                    Name = "Камила",
                    MiddleName = "Сергеевна",
                    RecordNumber = 18051486,
                    DateOfBirth = "22.12.1999",
                    Faculty = engineeringAndEconomic
                };
                Student s19 = new Student
                {
                    LastName = "Андропов",
                    Name = "Иван",
                    MiddleName = "Андреевич",
                    RecordNumber = 18051652,
                    DateOfBirth = "08.05.1998",
                    Faculty = engineering
                };
                Student s20 = new Student
                {
                    LastName = "Стрелков",
                    Name = "Егор",
                    MiddleName = "Иванович",
                    RecordNumber = 18051653,
                    DateOfBirth = "16.09.1999",
                    Faculty = engineeringAndEconomic
                };
                db.Students.AddRange(s1, s2, s3, s4, s5, s6,
                    s7, s8, s9, s10, s11, s12, s13, s14, s15,
                    s16, s17, s18, s19, s20);
                db.SaveChanges();
            }
            MessageBox.Show("БД обновлена по-умолчанию!");
        }


        private void button1_Click(object sender, EventArgs e)//вывод БД по Студентам
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Students.Load();

                dataGridView1.DataSource = db.Students.ToList();

            }
        }

        private void button2_Click(object sender, EventArgs e)//добавить
        {
            StudentsForm studentsForm = new StudentsForm();

            DialogResult result = studentsForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;

            using (ApplicationContext db = new ApplicationContext())
            {
                Student student = new Student();

                //можно не заполнять свойство
                try
                {
                    student.MiddleName = studentsForm.textBox3.Text;
                }
                catch { }
                student.RecordNumber = Convert.ToInt32(studentsForm.textBox4.Text);
                try
                {
                    student.DateOfBirth = studentsForm.textBox5.Text;
                }
                catch { }

                // можно и так...так короче запись
                student.FacultyId = int.Parse(studentsForm.textBox6.Text);

                db.Students.Add(student);
                db.SaveChanges();

                db.Students.Load();
                dataGridView1.DataSource = db.Students.ToList();
            }

            MessageBox.Show("Студент добавлен!");

        }

        private void button3_Click(object sender, EventArgs e)//редактировать
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    int index = dataGridView1.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView1[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;

                    Student student = db.Students.Find(id);
                    StudentsForm studentsForm = new StudentsForm();

                    studentsForm.textBox1.Text = student.LastName;
                    studentsForm.textBox2.Text = student.Name;
                    studentsForm.textBox3.Text = student.MiddleName;
                    studentsForm.textBox4.Text = student.RecordNumber.ToString();
                    studentsForm.textBox5.Text = student.DateOfBirth;
                    studentsForm.textBox6.Text = student.FacultyId.ToString();

                    DialogResult result = studentsForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;

                    student.LastName = studentsForm.textBox1.Text;
                    student.Name = studentsForm.textBox2.Text;
                    student.MiddleName = studentsForm.textBox3.Text;
                    student.RecordNumber = Convert.ToInt32(studentsForm.textBox4.Text);
                    //можно не заполнять свойство
                    try
                    {
                        student.DateOfBirth = studentsForm.textBox5.Text;
                    }
                    catch { }

                    student.FacultyId = int.Parse(studentsForm.textBox6.Text);

                    db.Students.Update(student);
                    db.SaveChanges();

                    db.Students.Load();
                    dataGridView1.DataSource = db.Students.ToList();

                    MessageBox.Show("Объект изменён!");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)//удалить
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    int index = dataGridView1.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView1[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;

                    Student student = db.Students.Find(id);
                    db.Students.Remove(student);
                    db.SaveChanges();

                    MessageBox.Show("Объект удален");
                }
                //автообновление вывода dataGridView1
                db.Students.Load();// вроде не нужен.... но пусть будет,
                                   // ещё раз загрузим в контекст/ обноввим его 

                dataGridView1.DataSource = db.Students.ToList();

            }


        }

        private void button5_Click(object sender, EventArgs e)// вывод БД по факультетам
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Facultys.Load();

                dataGridView2.DataSource = db.Facultys.ToList();

            }

        }

        private void button6_Click(object sender, EventArgs e)// Добавить
        {
            
            using (ApplicationContext db = new ApplicationContext())
            {
                FacultyForm facultyForm = new FacultyForm();

                DialogResult result = facultyForm.ShowDialog(this);
                if (result == DialogResult.Cancel)
                    return;

                Faculty faculty = new Faculty();
                faculty.NameFaculty = facultyForm.textBox1.Text;

                db.Facultys.Add(faculty);
                db.SaveChanges();

                db.Facultys.Load();
                dataGridView2.DataSource = db.Facultys.ToList();
            }

            MessageBox.Show("Факультет добавлен!");

        }

        private void button7_Click(object sender, EventArgs e)// Редактировать
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    int index = dataGridView2.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView2[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;

                    Faculty faculty = db.Facultys.Find(id);
                    FacultyForm facultyForm = new FacultyForm();

                    facultyForm.textBox1.Text = faculty.NameFaculty;
                    
                    DialogResult result = facultyForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;

                    faculty.NameFaculty = facultyForm.textBox1.Text;
                    
                    db.Facultys.Update(faculty);
                    db.SaveChanges();

                    db.Facultys.Load();
                    dataGridView2.DataSource = db.Facultys.ToList();

                }

                MessageBox.Show("Объект изменён!");
            }

        }

        private void button8_Click(object sender, EventArgs e) // удалить
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    int index = dataGridView1.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView1[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;

                    Faculty faculty = db.Facultys.Find(id);
                    db.Facultys.Remove(faculty);
                    db.SaveChanges();

                }

                MessageBox.Show("Объект удален");

                //автообновление вывода dataGridView
                db.Facultys.Load();
                dataGridView2.DataSource = db.Facultys.ToList();

            }
        }

        void shoMore_Click(object sender, EventArgs e)// ПКМ - на строке - Подробнее о Факультете
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    int id = int.Parse(dataGridView2.CurrentRow.Cells[0].Value.ToString());
                    
                    //по соответствующему id
                    Faculty faculty = db.Facultys.Include(s => s.Students).FirstOrDefault(x => x.Id == id);

                    FacultyInfo facultyInfo = new FacultyInfo();//созд и показать экз формы 

                    facultyInfo.textBox1.Text = faculty.NameFaculty;
                    facultyInfo.Show();

                    try
                    {
                        facultyInfo.dataGridView1.DataSource = faculty.Students.ToList();
                    }
                    catch
                    {
                        MessageBox.Show("Ошибка");
                    }

                }
            }
        }

        //Отчёт по Студентам
        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Environment.CurrentDirectory, "Export", "export_students.xlsx");

            using (ApplicationContext db = new ApplicationContext())
            {
                // так всё работает, топорно ,но работает 
                /*
                var products = db.Products.ToList();
                var workBook = new XLWorkbook();

                var sheet = workBook.Worksheets.Add("Products");

                sheet.Cell("A1").InsertTable(products);
                workBook.SaveAs(path);

                //смотри https://github.com/ClosedXML/ClosedXML/wiki
                // примеры есть, но несовсем понял... просто втавил таблицу
                //                                      по первой ячейке
                */

                var students = db.Students.ToList();
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

                foreach (var item in students)
                {
                    sheet.Cell(startRow, startCol++).Value = item.Id;
                    sheet.Cell(startRow, startCol++).Value = item.LastName;
                    sheet.Cell(startRow, startCol++).Value = item.Name;
                    sheet.Cell(startRow, startCol++).Value = item.MiddleName;
                    sheet.Cell(startRow, startCol++).Value = item.RecordNumber;
                    sheet.Cell(startRow, startCol++).Value = item.DateOfBirth;
                    sheet.Cell(startRow, startCol).Value = item.FacultyId;

                    //sheet.Cell(startRow, startCol).Value = item.Faculty.Students.ToString();//!! Исправь этот баг!!
                    // так тоже не работает...Students[1]...
                    //sheet.Cell(startRow, startCol).Value = item.Faculty.Students[1].ToString();
                    
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

        //Отчёт по Факультетам
        private void экспортВExcelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Environment.CurrentDirectory, "Export", "export_facultys.xlsx");

            using (ApplicationContext db = new ApplicationContext())
            {
                var facultys = db.Facultys.ToList();
                XLWorkbook workBook = new XLWorkbook();

                var sheet = workBook.Worksheets.Add("Facultys");
                var startRow = 2;
                int startCol = 1;

                //делаем "шапку таблицы"
                sheet.Cell("A1").Value = "Id";
                sheet.Cell("B1").Value = "Наименование";
                

                foreach (var item in facultys)
                {
                    sheet.Cell(startRow, startCol++).Value = item.Id;
                    sheet.Cell(startRow, startCol).Value = item.NameFaculty;

                    startCol = 1;
                    startRow++;
                }

                // авт. р-р ячеек по их содержимому
                sheet.Columns(1, 2).AdjustToContents();
                // увелич высоту первой строки
                sheet.Row(1).Height = 25;

                workBook.SaveAs(path);

                MessageBox.Show("Отчёт сформирован!");
            }
        }
    }
}