using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain.Data;

namespace DYL.EmailIntegration.UI.Helpers
{
    public static class Enumerable
    {
        public static bool ContainsId(this IEnumerable<Email> source, Email value)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            if (string.IsNullOrEmpty(value.Id))
                throw new ApplicationException("Email ID is null or empty");

            return source.Any(element => element.Id == value.Id);
        }
    }
}
