using System;
using System.Text;

namespace LoopCAD.WPF
{
    public class SnakeCase
    {
        public static string Convert(string text)
        {
            if(text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            if (text.Length < 2)
            {
                return text.ToLowerInvariant();
            }

            var builder = new StringBuilder();
            builder.Append(char.ToLowerInvariant(text[0]));

            for (int i = 1; i < text.Length; ++i)
            {
                char c = text[i];
                if (char.IsUpper(c))
                {
                    builder.Append('_');
                    builder.Append(char.ToLowerInvariant(c));
                }
                else
                {
                    builder.Append(c);
                }
            }

            return builder.ToString();
        }
    }
}
