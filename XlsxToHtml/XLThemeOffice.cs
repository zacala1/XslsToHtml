using ClosedXML.Excel;

namespace XlsxToHtml
{
    internal class XLThemeOffice : IXLTheme
    {
        public XLThemeOffice()
        {
            Background1 = XLColor.FromArgb(255, 255, 255);
            Text1 = XLColor.FromArgb(0, 0, 0);
            Background2 = XLColor.FromArgb(231, 230, 230);
            Text2 = XLColor.FromArgb(68, 84, 106);
            Accent1 = XLColor.FromArgb(91, 155, 213);
            Accent2 = XLColor.FromArgb(237, 125, 49);
            Accent3 = XLColor.FromArgb(165, 165, 165);
            Accent4 = XLColor.FromArgb(255, 192, 0);
            Accent5 = XLColor.FromArgb(68, 114, 196);
            Accent6 = XLColor.FromArgb(112, 173, 71);
            Hyperlink = XLColor.FromArgb(5, 99, 193);
            FollowedHyperlink = XLColor.FromArgb(149, 79, 114);
        }

        public XLColor Background1 { get; set; }
        public XLColor Text1 { get; set; }
        public XLColor Background2 { get; set; }
        public XLColor Text2 { get; set; }
        public XLColor Accent1 { get; set; }
        public XLColor Accent2 { get; set; }
        public XLColor Accent3 { get; set; }
        public XLColor Accent4 { get; set; }
        public XLColor Accent5 { get; set; }
        public XLColor Accent6 { get; set; }
        public XLColor Hyperlink { get; set; }
        public XLColor FollowedHyperlink { get; set; }

        public XLColor ResolveThemeColor(XLThemeColor themeColor)
        {
            return themeColor switch
            {
                XLThemeColor.Background1 => this.Background1,
                XLThemeColor.Text1 => this.Text1,
                XLThemeColor.Background2 => this.Background2,
                XLThemeColor.Text2 => this.Text2,
                XLThemeColor.Accent1 => this.Accent1,
                XLThemeColor.Accent2 => this.Accent2,
                XLThemeColor.Accent3 => this.Accent3,
                XLThemeColor.Accent4 => this.Accent4,
                XLThemeColor.Accent5 => this.Accent5,
                XLThemeColor.Accent6 => this.Accent6,
                XLThemeColor.Hyperlink => this.Hyperlink,
                XLThemeColor.FollowedHyperlink => this.FollowedHyperlink,
                _ => null
            };
        }
    }
}