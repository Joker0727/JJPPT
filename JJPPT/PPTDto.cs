using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JJPPT
{
    public class PPTDto
    {
        public Int64 Id { get; set; }
        public Int64 PPTId { get; set; }
        public string PPTName { get; set; }
        public string PPTUrl { get; set; }
        public Int64 IsDownLoad { get; set; }
    }
}
