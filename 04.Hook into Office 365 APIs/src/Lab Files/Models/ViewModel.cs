using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCResearchTracker.Models
{
    public class ViewModel
    {
        public List<MyFile> Files { get; set; }
        public List<MyContact> Contacts { get; set; }

        public MyProject Project { get; set; }

    }
}