using Microsoft.EntityFrameworkCore;
using System.Linq;
using System.Numerics;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.EntityFrameworkCore.ChangeTracking.Internal;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing;
using ClosedXML.Excel;

namespace DekanatDB
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }



        private void ���������������������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Database.EnsureDeleted();

            }

            //���������� ���������
            using (ApplicationContext db = new ApplicationContext())
            {
                
                // �������� � ���������� �������
                Faculty engineering = new Faculty { NameFaculty = "� - ������������������" };
                Faculty instrumentation = new Faculty { NameFaculty = "� - �������������������" };
                Faculty computerScience = new Faculty { NameFaculty = "��� - ����������� � �������������� �������" };
                Faculty engineeringAndEconomic = new Faculty { NameFaculty = "�� - ���������-�������������" };
                db.Facultys.AddRange(engineering, instrumentation, computerScience, engineeringAndEconomic);


                Student s1 = new Student
                {
                    LastName = "�����������",
                    Name = "������",
                    MiddleName = "���������",
                    RecordNumber = 18051351,
                    DateOfBirth = "08.31.1998",
                    Faculty = engineering
                };
                Student s2 = new Student
                {
                    LastName = "�����",
                    Name = "�������",
                    MiddleName = "����������",
                    RecordNumber = 18051353,
                    DateOfBirth = "01.06.1998",
                    Faculty = engineering
                };
                Student s3 = new Student
                {
                    LastName = "���������",
                    Name = "����",
                    MiddleName = "�������",
                    RecordNumber = 18051354,
                    DateOfBirth = "29.05.2000",
                    Faculty = instrumentation
                };
                Student s4 = new Student
                {
                    LastName = "������",
                    Name = "�����",
                    MiddleName = "���������",
                    RecordNumber = 18051301,
                    DateOfBirth = "21.07.1999",
                    Faculty = instrumentation
                };
                Student s5 = new Student
                {
                    LastName = "����",
                    Name = "��������",
                    MiddleName = "���������",
                    RecordNumber = 18051355,
                    DateOfBirth = "02.04.2000",
                    Faculty = instrumentation
                };
                Student s6 = new Student
                {
                    LastName = "������",
                    Name = "�����",
                    MiddleName = "����������",
                    RecordNumber = 18051479,
                    DateOfBirth = "03.04.1998",
                    Faculty = instrumentation
                };
                Student s7 = new Student
                {
                    LastName = "�����",
                    Name = "������",
                    MiddleName = "����������",
                    RecordNumber = 18051358,
                    DateOfBirth = "17.12.1998",
                    Faculty = engineering
                };
                Student s8 = new Student
                {
                    LastName = "��������",
                    Name = "�����",
                    MiddleName = "����������",
                    RecordNumber = 18051480,
                    DateOfBirth = "18.04.1998",
                    Faculty = computerScience
                };
                Student s9 = new Student
                {
                    LastName = "��������",
                    Name = "�����",
                    MiddleName = "����������",
                    RecordNumber = 18051303,
                    DateOfBirth = "06.06.1999",
                    Faculty = engineering
                };
                Student s10 = new Student
                {
                    LastName = "�������",
                    Name = "��������",
                    MiddleName = "���������",
                    RecordNumber = 18051400,
                    DateOfBirth = "04.04.1998",
                    Faculty = computerScience
                };
                Student s11 = new Student
                {
                    LastName = "�������",
                    Name = "����",
                    MiddleName = "���������",
                    RecordNumber = 18051368,
                    DateOfBirth = "01.11.2000",
                    Faculty = engineeringAndEconomic
                };
                Student s12 = new Student
                {
                    LastName = "��������",
                    Name = "�������",
                    MiddleName = "���������",
                    RecordNumber = 18051372,
                    DateOfBirth = "21.08.1999",
                    Faculty = computerScience
                };
                Student s13 = new Student
                {
                    LastName = "�������",
                    Name = "������",
                    MiddleName = "�������������",
                    RecordNumber = 18051376,
                    DateOfBirth = "18.03.1998",
                    Faculty = computerScience
                };
                Student s14 = new Student
                {
                    LastName = "��������",
                    Name = "���������",
                    MiddleName = "�������������",
                    RecordNumber = 18051484,
                    DateOfBirth = "23.02.1998",
                    Faculty = computerScience
                };
                Student s15 = new Student
                {
                    LastName = "������",
                    Name = "�������",
                    MiddleName = "����������",
                    RecordNumber = 18051306,
                    DateOfBirth = "20.04.1999",
                    Faculty = engineering
                };
                Student s16 = new Student
                {
                    LastName = "��������",
                    Name = "������",
                    MiddleName = "�������",
                    RecordNumber = 18051307,
                    DateOfBirth = "18.04.1999",
                    Faculty = computerScience
                };
                Student s17 = new Student
                {
                    LastName = "������",
                    Name = "����",
                    MiddleName = "����������",
                    RecordNumber = 18051384,
                    DateOfBirth = "15.07.1999",
                    Faculty = engineeringAndEconomic
                };
                Student s18 = new Student
                {
                    LastName = "�������",
                    Name = "������",
                    MiddleName = "���������",
                    RecordNumber = 18051486,
                    DateOfBirth = "22.12.1999",
                    Faculty = engineeringAndEconomic
                };
                Student s19 = new Student
                {
                    LastName = "��������",
                    Name = "����",
                    MiddleName = "���������",
                    RecordNumber = 18051652,
                    DateOfBirth = "08.05.1998",
                    Faculty = engineering
                };
                Student s20 = new Student
                {
                    LastName = "��������",
                    Name = "����",
                    MiddleName = "��������",
                    RecordNumber = 18051653,
                    DateOfBirth = "16.09.1999",
                    Faculty = engineeringAndEconomic
                };
                db.Students.AddRange(s1, s2, s3, s4, s5, s6,
                    s7, s8, s9, s10, s11, s12, s13, s14, s15,
                    s16, s17, s18, s19, s20);
                db.SaveChanges();
            }
            MessageBox.Show("�� ��������� ��-���������!");
        }


        private void button1_Click(object sender, EventArgs e)//����� �� �� ���������
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Students.Load();

                dataGridView1.DataSource = db.Students.ToList();

            }
        }

        private void button2_Click(object sender, EventArgs e)//��������
        {
            StudentsForm studentsForm = new StudentsForm();

            DialogResult result = studentsForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;

            Student student = new Student();
            student.LastName = studentsForm.textBox1.Text;
            student.Name = studentsForm.textBox2.Text;
            student.MiddleName = studentsForm.textBox3.Text;
            student.RecordNumber = Convert.ToInt32(studentsForm.textBox4.Text);
            //����� �� ��������� ��������
            try
            {
                student.DateOfBirth = studentsForm.textBox5.Text;
            }
            catch { }
            //!!������ ���� ��Ȩ� � �������� � �������� ������������ �� ���������
            student.FacultyId = Convert.ToInt32(studentsForm.textBox6.Text);//!!!

            using (ApplicationContext db = new ApplicationContext())
            {
                db.Students.Add(student);
                db.SaveChanges();

                db.Students.Load();
                dataGridView1.DataSource = db.Students.ToList();
            }

            MessageBox.Show("������� ��������!");


        }

        private void button3_Click(object sender, EventArgs e)
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
                    studentsForm.textBox4.Text = student.FacultyId.ToString();


                    DialogResult result = productForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;

                    student.LastName = studentsForm.textBox1.Text;
                    student.Name = studentsForm.textBox2.Text;
                    student.MiddleName = studentsForm.textBox3.Text;
                    student.RecordNumber = Convert.ToInt32(studentsForm.textBox4.Text);
                    //����� �� ��������� ��������
                    try
                    {
                        student.DateOfBirth = studentsForm.textBox5.Text;
                    }
                    catch { }
                    //!!������ ���� ��Ȩ� � �������� � �������� ������������ �� ���������
                    student.FacultyId = Convert.ToInt32(studentsForm.textBox6.Text);//!!!

                    db.Students.Add(student);
                    db.SaveChanges();

                    db.Students.Load();
                    dataGridView1.DataSource = db.Students.ToList();

                    MessageBox.Show("������ ������!");

                }

            }
        }
    }
}