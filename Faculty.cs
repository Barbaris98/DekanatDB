using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DekanatDB
{
    public class Faculty
    {
        public int Id { get; set; }
        public string? NameFaculty { get; set; }
        public List<Student> Students { get; set; } = new();
    }
}
