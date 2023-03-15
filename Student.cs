using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace DekanatDB
{
    public class Student
    {
        public int Id { get; set; }
        public string? LastName { get; set; }
        public string? Name { get; set; }
        public string? MiddleName { get; set; }
        public int Age { get; set; }
        public bool FullTimeEducation { get; set; }

        public int FacultyId { get; set; }
        public Faculty? Faculty { get; set; }  // компания пользователя

    }
}
