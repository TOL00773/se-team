using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IT_School
{
    public static class StringExtensions
    {
        public static bool Contains(this String str, String substring,
                                    StringComparison comp)
        {
            if (substring == null)
                throw new ArgumentNullException("substring",
                                                "substring cannot be null.");
            else if (!Enum.IsDefined(typeof(StringComparison), comp))
                throw new ArgumentException("comp is not a member of StringComparison",
                                            "comp");

            return str.IndexOf(substring, comp) >= 0;
        }
    }
    public class Selecting
    {
        ObservableCollection<Organization> orgs { get; set; }
        void ByAccName( string keyword, ObservableCollection<Organization> mylist)
        {
            foreach(Organization i in mylist)
            {
                if (i.AccName.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    orgs.Add(i);
                }
            }
        }
    }
}
