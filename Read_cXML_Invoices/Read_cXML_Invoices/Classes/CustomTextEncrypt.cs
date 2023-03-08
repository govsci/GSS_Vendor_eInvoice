using Microsoft.VisualBasic;

namespace Read_cXML_Invoices.Classes
{
    public static class CustomTextEncrypt
    {
        private static string Base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            + "abcdefghijklmnopqrstuvwxyz"
            + "0123456789"
            + "+/";
        public static string Encode(string val)
        {
            int c1 = 0, c2 = 0, c3 = 0, w1 = 0, w2 = 0, w3 = 0, w4 = 0;
            string strOut = "";

            for (int n = 1; n <= val.Length; n += 3)
            {
                c1 = Strings.Asc(Strings.Mid(val, n, 1));
                c2 = Strings.Asc(Strings.Mid(val, n + 1, 1) + Strings.Chr(0));
                c3 = Strings.Asc(Strings.Mid(val, n + 2, 1) + Strings.Chr(0));

                w1 = (int)(c1 / 4);
                w2 = (c1 & 3) * 16 + (int)(c2 / 16);

                if (val.Length >= (n + 1))
                    w3 = (c2 & 15) * 4 + (int)(c3 / 64);
                else
                    w3 = -1;

                if (val.Length >= (n + 2))
                    w4 = c3 & 63;
                else
                    w4 = -1;

                strOut += mimeencode(w1) + mimeencode(w2) + mimeencode(w3) + mimeencode(w4);
            }

            return strOut;
        }
        private static string mimeencode(int i)
        {
            if (i >= 0)
                return Strings.Mid(Base64Chars, i + 1, 1);
            else
                return "";
        }

        public static string Decode(string val)
        {
            string strOut = "";
            int w1 = 0, w2 = 0, w3 = 0, w4 = 0;
            for (int n = 1; n <= val.Length; n += 4)
            {
                w1 = mimedecode(Strings.Mid(val, n, 1));
                w2 = mimedecode(Strings.Mid(val, n + 1, 1));
                w3 = mimedecode(Strings.Mid(val, n + 2, 1));
                w4 = mimedecode(Strings.Mid(val, n + 3, 1));
                if (w2 >= 0)
                    strOut += Strings.Chr(((w1 * 4 + (int)(w2 / 16)) & 255));
                if (w3 >= 0)
                    strOut += Strings.Chr(((w2 * 16 + (int)(w3 / 4)) & 255));
                if (w4 >= 0)
                    strOut += Strings.Chr(((w3 * 64 + w4) & 255));
            }
            return strOut;
        }
        private static int mimedecode(string str)
        {
            if (str.Length == 0)
                return -1;
            else
                return Strings.InStr(Base64Chars, str) - 1;
        }
    }
}
