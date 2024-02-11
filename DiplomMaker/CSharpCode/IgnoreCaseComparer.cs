using System.Collections.Generic;

namespace DiplomMaker
{
    /// <summary>
    /// Сравнивает строки, игнорируя регистр.
    /// Технически считает, что все строки в верхнем регистре.
    /// </summary>
    public class IgnoreCaseComparer : IEqualityComparer<string>
    {
        public bool Equals(string x, string y)
        {
            return x.ToUpper().Equals(y.ToUpper());
        }

        public int GetHashCode(string obj)
        {
            return obj.ToUpper().GetHashCode();
        }
    }
}
