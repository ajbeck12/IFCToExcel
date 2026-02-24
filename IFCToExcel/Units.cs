using System;

namespace IfcToExcelWinForms
{
    public static class Units
    {
        public static double MmToIn(double mm) => mm / 25.4;

        // Formats 4.5" as 4"1/2, 2.25" as 2"1/4, 3.0" as 3"
        public static string InToArchitectural(double inches, int denom = 16)
        {
            inches = Math.Round(inches * denom) / denom;

            int whole = (int)Math.Floor(inches);
            double frac = inches - whole;

            if (Math.Abs(frac) < 1e-9)
                return $"{whole}\"";

            int num = (int)Math.Round(frac * denom);
            int g = Gcd(num, denom);
            num /= g;
            int den = denom / g;

            if (whole == 0)
                return $"0\"{num}/{den}";

            return $"{whole}\"{num}/{den}";
        }

        private static int Gcd(int a, int b)
        {
            while (b != 0) { int t = a % b; a = b; b = t; }
            return Math.Abs(a);
        }

        // If equal spacing, output like 3*3" (3 spaces between 4 bolts)
        public static string FormatSpacing(double[] relInchesAscending, double tolIn = 1.0 / 64.0)
        {
            if (relInchesAscending.Length <= 1) return "";

            var s = (double[])relInchesAscending.Clone();
            Array.Sort(s);

            var deltas = new double[s.Length - 1];
            for (int i = 0; i < deltas.Length; i++)
                deltas[i] = s[i + 1] - s[i];

            double first = deltas[0];
            bool allEqual = true;
            for (int i = 1; i < deltas.Length; i++)
                if (Math.Abs(deltas[i] - first) > tolIn) { allEqual = false; break; }

            if (allEqual)
            {
                if (deltas.Length == 1)
                    return InToArchitectural(first);  // single gap: no count prefix
                return $"{deltas.Length}*{InToArchitectural(first)}";
            }

            string[] parts = new string[deltas.Length];
            for (int i = 0; i < deltas.Length; i++)
                parts[i] = InToArchitectural(deltas[i]);

            return string.Join(", ", parts);
        }

        public static string? TryFormatSignedArchitectural(double inches, int denom = 16)
        {
            if (double.IsNaN(inches) || double.IsInfinity(inches)) return null;
            bool neg = inches < 0;
            inches = Math.Abs(inches);
            var s = InToArchitectural(inches, denom);
            return neg ? "-" + s : s;
        }
    }
}