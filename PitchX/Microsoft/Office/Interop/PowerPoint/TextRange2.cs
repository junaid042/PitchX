using System;

namespace Microsoft.Office.Interop.PowerPoint
{
    internal class TextRange2
    {
        public object Text { get; internal set; }
        public int ParagraphsCount { get; internal set; }

        internal TextRange2 Paragraphs(int i, int v)
        {
            throw new NotImplementedException();
        }
    }
}