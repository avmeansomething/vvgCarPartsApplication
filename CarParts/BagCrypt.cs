using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarParts
{
    class BagCrypt
    {
        char[] characters = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
                                         'a','b','c','d','e','f','g','h','i',
                                         'j','k','l','m','n','o','p','r','s','t','q','u','v','w','x','y','z'};

        int[] w = { 2, 3, 6, 13, 27, 52 };
        int q = 105;
        int r = 31;
        int[] b = new int[6];

        public BagCrypt()
        {
            for (int i = 0; i < w.Length; i++)
            {
                b[i] = w[i] * r % q;
            }
        }

        public string CryptBag(string text)
        {
            string crypted = string.Empty;
            string word = text.ToLower();
            var s = word.ToCharArray();
            string[] nom = new string[s.Length];
            for (int i = 0; i < s.Length; i++)
            {
                for (int j = 0; j < characters.Length; j++)
                {
                    if (s[i] == characters[j])
                        nom[i] += j;
                }
            }

            for (int i = 0; i < nom.Length; i++)
            {
                string a = Convert.ToString(Convert.ToInt32(nom[i]), 2);
                int k = 0;
                while (a.Length < 6)
                {
                    a = a.Insert(0, "0");
                }

                for (int j = 0; j < a.Length; j++)
                {
                    if (a[j] != '0')
                        k += Convert.ToInt32(b[j]);
                }
                crypted += k.ToString() + " ";
            }
            return crypted;
        }

        private string DecryptBag(string text)
        {
            string decrypted = string.Empty;
            string[] s = text.Split(' ');
            int[] a = new int[s.Length - 1];

            for (int i = 0; i < s.Length - 1; i++)
            {
                a[i] = Convert.ToInt32(s[i]) * 61 % q;
            }


            for (int i = 0; i < a.Length; i++)
            {
                string viv = "";
                for (int j = w.Length - 1; j > -1; j--)
                {
                    if (w[j] <= a[i])
                    {
                        a[i] -= w[j];
                        viv = viv.Insert(0, "1");
                    }
                    else viv = viv.Insert(0, "0");
                }
                viv = Convert.ToString(Convert.ToInt32(viv, 2));
                decrypted += Convert.ToString(characters[Convert.ToInt32(viv)]);
            }
            return decrypted;
        }
    }
}
