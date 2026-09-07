using System;

namespace JarvisPowerPoint
{
    internal sealed class PresentationSnapshot
    {
        public string SessionKey { get; set; }
        public string PresentationKey { get; set; }
        public string PresentationName { get; set; }
        public int SlideId { get; set; }
        public int SlideNumber { get; set; }
        public int SlideCount { get; set; }
        public string Title { get; set; }
        public bool IsBlack { get; set; }
        public string SavedPath { get; set; }
        public IntPtr WindowHandle { get; set; }
    }

    internal sealed class SlideAlias
    {
        public string Name { get; set; }
        public int SlideId { get; set; }
        public int SlideNumber { get; set; }
        public string Title { get; set; }
    }
}
