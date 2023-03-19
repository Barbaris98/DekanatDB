using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace DekanatDB
{
    [Index("RecordNumber", IsUnique = true)] //настроим также уникальность номера зачётной книжки
    public class Student
    {
        public int Id { get; set; }
        public string LastName { get; set; } = null!;
        public string Name { get; set; } = null!;
        public string? MiddleName { get; set; }

        public int RecordNumber { get; set; }
        public string? DateOfBirth { get; set; }
        
        public int FacultyId { get; set; }
        public Faculty? Faculty { get; set; } 

    }
}
