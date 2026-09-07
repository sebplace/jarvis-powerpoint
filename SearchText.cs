using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace JarvisPowerPoint
{
    internal sealed class SearchDocument
    {
        public readonly int Id;
        public readonly int Number;
        public readonly string Title;
        public readonly string NormalizedTitle;
        public readonly string NormalizedContent;
        private readonly string fingerprint;

        public SearchDocument(int id, int number, string title, string content)
        {
            Id = id;
            Number = number;
            Title = title ?? string.Empty;
            NormalizedTitle = PowerPointController.NormalizeForSearch(Title);
            NormalizedContent = PowerPointController.NormalizeForSearch(content);
            using (SHA256 hash = SHA256.Create())
                fingerprint = Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(content ?? string.Empty)));
        }

        public bool SameContent(SearchDocument other)
        {
            return other != null && Id == other.Id && Number == other.Number &&
                Title == other.Title && fingerprint == other.fingerprint;
        }

        private SearchDocument(SearchDocument source)
        {
            Id = source.Id;
            Number = source.Number;
            Title = source.Title;
            fingerprint = source.fingerprint;
            NormalizedTitle = NormalizedContent = string.Empty;
        }

        public SearchDocument WithoutText()
        {
            return new SearchDocument(this);
        }
    }

    internal abstract class SearchSlideReader : IDisposable
    {
        public abstract bool ReadNext();
        public abstract SearchDocument Document { get; }
        public abstract void Dispose();
    }

    internal sealed class PowerPointSearchSlideReader : SearchSlideReader
    {
        private PowerPoint.Slide slide;
        private IEnumerator<string> chunks;
        private StringBuilder content = new StringBuilder();
        private readonly int id;
        private readonly int number;
        private string title;
        private SearchDocument document;

        // The caller transfers exactly one slide reference on successful construction.
        public PowerPointSearchSlideReader(PowerPoint.Slide slide)
        {
            id = slide.SlideID;
            number = slide.SlideIndex;
            title = PowerPointController.GetSlideTitle(slide);
            chunks = PowerPointSearchText.Chunks(slide).GetEnumerator();
            this.slide = slide;
        }

        public override SearchDocument Document { get { return document; } }

        public override bool ReadNext()
        {
            if (document != null) return true;
            if (chunks == null) throw new ObjectDisposedException("SearchSlideReader");
            if (chunks.MoveNext())
            {
                content.Append(chunks.Current);
                return false;
            }
            document = new SearchDocument(id, number, title, content.ToString());
            return true;
        }

        public override void Dispose()
        {
            try
            {
                if (chunks != null) chunks.Dispose();
            }
            finally
            {
                chunks = null;
                PowerPointSearchText.Release(slide);
                slide = null;
                content = null;
                title = null;
                document = null;
            }
        }
    }

    // One shared text walk serves both the legacy controller and incremental searches.
    // Iterator finally blocks release nested COM references even when cancelled mid-cell.
    internal static class PowerPointSearchText
    {
        public static string ReadAll(PowerPoint.Slide slide)
        {
            var text = new StringBuilder();
            foreach (string chunk in Chunks(slide)) text.Append(chunk);
            return text.ToString();
        }

        internal static IEnumerable<string> Chunks(PowerPoint.Slide slide)
        {
            PowerPoint.Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                int count = shapes.Count;
                yield return null;
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        foreach (string text in ShapeChunks(shape)) yield return text;
                    }
                    finally { Release(shape); }
                }
            }
            finally { Release(shapes); }
        }

        private static IEnumerable<string> ShapeChunks(PowerPoint.Shape shape)
        {
            yield return null;
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                PowerPoint.GroupShapes children = null;
                try
                {
                    children = shape.GroupItems;
                    int count = children.Count;
                    yield return null;
                    for (int index = 1; index <= count; index++)
                    {
                        PowerPoint.Shape child = null;
                        try
                        {
                            child = (PowerPoint.Shape)children[index];
                            foreach (string text in ShapeChunks(child)) yield return text;
                        }
                        finally { Release(child); }
                    }
                }
                finally { Release(children); }
            }

            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
            {
                PowerPoint.TextFrame frame = null;
                PowerPoint.TextRange range = null;
                string text = null;
                try
                {
                    frame = shape.TextFrame;
                    if (frame.HasText == Office.MsoTriState.msoTrue)
                    {
                        range = frame.TextRange;
                        text = " " + range.Text;
                    }
                }
                finally { Release(range); Release(frame); }
                yield return text;
            }

            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                PowerPoint.Table table = null;
                PowerPoint.Rows rows = null;
                PowerPoint.Columns columns = null;
                try
                {
                    table = shape.Table;
                    rows = table.Rows;
                    columns = table.Columns;
                    int rowCount = rows.Count;
                    int columnCount = columns.Count;
                    yield return null;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int column = 1; column <= columnCount; column++)
                        {
                            PowerPoint.Cell cell = null;
                            PowerPoint.Shape cellShape = null;
                            PowerPoint.TextFrame frame = null;
                            PowerPoint.TextRange range = null;
                            string text = null;
                            try
                            {
                                cell = table.Cell(row, column);
                                cellShape = cell.Shape;
                                frame = cellShape.TextFrame;
                                if (frame.HasText == Office.MsoTriState.msoTrue)
                                {
                                    range = frame.TextRange;
                                    text = " " + range.Text;
                                }
                            }
                            finally { Release(range); Release(frame); Release(cellShape); Release(cell); }
                            yield return text;
                        }
                    }
                }
                finally { Release(columns); Release(rows); Release(table); }
            }

            string alternative = shape.AlternativeText;
            yield return string.IsNullOrWhiteSpace(alternative) ? null : " " + alternative;
        }

        internal static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
