using TFlex.Configuration;

namespace TFlexCADCommandsUser
{
    public class Format
    {
        private static bool FormatType(string format)
        {
            var document = TFlex.Application.ActiveDocument;
            var activPage = document.ActivePage;
            bool flag = false;
            document.BeginChanges("Изменить формат на " + format);
            if (format == activPage.Properties.Paper.Format)
            {
                flag = true;
                activPage.Properties.Paper.Landscape = document.ActivePage.Properties.Paper.Landscape == LandscapePaper.Horizontal ? LandscapePaper.Vertical : LandscapePaper.Horizontal;
            }
            activPage.Properties.Paper.Format = format;
            document.EndChanges();
            return flag;
            //document.Regenerate(new RegenerateOptions { Full = true });
        }
        public static bool BCH() { return FormatType("БЧ"); }
        public static bool A4() { return FormatType("A4"); }
        public static bool A3() { return FormatType("A3"); }
        public static bool A2() { return FormatType("A2"); }
        public static bool A1() { return FormatType("A1"); }
        public static bool A0() { return FormatType("A0"); }
        public static bool A4x3() { return FormatType("A4x3"); }
        public static bool A3x3() { return FormatType("A3x3"); }
        public static bool A2x3() { return FormatType("A2x3"); }
        public static bool A1x3() { return FormatType("A1x3"); }
        public static bool A0x2() { return FormatType("A0x2"); }
    }
}