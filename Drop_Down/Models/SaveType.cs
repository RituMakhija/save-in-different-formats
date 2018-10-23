using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Drop_Down.Models;
namespace Drop_Down.Models
{
    public class SaveType
    {
        public int saved { get; set; }
        public IEnumerable<tbl_saveFormat> savedFormatlst { get; set; } 
    }
}