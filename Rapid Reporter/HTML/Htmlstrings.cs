namespace Rapid_Reporter.HTML
{
    public static class Htmlstrings
    {
        public static string HtmlTitle = ": Session Report";
        public static string AHtmlHead;
        public static string BTitleOut;
        public static string CStyle;
        public static string DJavascript;
        public static string EBody;
        public static string GTable;
        public static string JTableEnd;
        public static string MHtmlEnd;

        static Htmlstrings()
        {
            AHtmlHead = "<html>\r\n    <head>\r\n        <meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n        <!--RRExV3.0-->\r\n        <title>";
            BTitleOut = "        </title>";

            CStyle = @"
        <style>
            html * {
                font-family: Verdana !important;
                font-size: 11px;
            }
            .aroundtable { font-family: Verdana; font-size: 11px; }
            H1 { text-align: center; font-family: Verdana; }
            H5 { text-align: center; font-family: Verdana; font-weight: normal; }
            table {
                margin-left: auto;
                margin-right: auto;
                min-width: 700px;
                width: 80%;
                border-collapse: collapse;
                border: 1px solid black;
            }
            table tr img {
                max-width: 350px;
                max-height: 250px;
                resize-scale: showall;
            }
            table tr td { padding: 2px; }

            table tr.Session,
            table tr.Scenario,
            table tr.Environment,
            table tr.Versions {
                font-weight: bold;
                background: #FAFAFA;
            }

            table tr.Bug\\/Issue { background: #FF4D4D; }
            table tr.Follow { background: #5CADFF; }
            table tr.Note,
            table tr.Test,
            table tr.Prerequisite,
            table tr.Summary,
            table tr.Screenshot,
            table tr.PlainText,
            table tr.Success {
                background: #FAFAFA;
            }

            table tr.Success { background: #80FF80; }

            table td.notetype { font-weight: bold; width: 190px; }
            table td.timestamp { font-weight: bold; width: 175px; }
        </style>";

            DJavascript = @"
        <script>
            var savedScrollY = 0;

            function ShowImgEle(eleId, bigImgId, littleImgId) {
                savedScrollY = window.scrollY;
                var sessionTable = document.getElementById('aroundtable');
                sessionTable.style.display = 'none';
                var bigImgDiv = document.getElementById(eleId);
                var bigImg = document.getElementById(bigImgId);
                var littleImg = document.getElementById(littleImgId);
                if (littleImg && bigImg) {
                    bigImg.src = littleImg.src;
                }
                bigImgDiv.style.display = 'block';
                bigImgDiv.style.textDecoration = 'underline';
                window.scrollTo(0, savedScrollY);
            }

            function HideImgEle(eleId) {
                var bigImgDiv = document.getElementById(eleId);
                var sessionTable = document.getElementById('aroundtable');
                sessionTable.style.display = 'inline';
                bigImgDiv.style.display = 'none';
                window.scrollTo(0, savedScrollY);
            }

            function ShowPlaintextNote(eleId) {
                savedScrollY = window.scrollY;
                var ele = document.getElementById(eleId);
                var eletable = document.getElementById('aroundtable');
                eletable.style.display = 'none';
                ele.style.display = 'block';
                window.scrollTo(0, savedScrollY);
            }

            function HidePlaintextNote(eleId) {
                var ele = document.getElementById(eleId);
                var eletable = document.getElementById('aroundtable');
                eletable.style.display = 'inline';
                ele.style.display = 'none';
                window.scrollTo(0, savedScrollY);
            }
        </script>
    </head>";

            EBody = "<body>\r\n        <div id=\"allbody\">\r\n            <h1>";
            GTable = "\r\n            </h1>\r\n            <!--[if IE]><h5>For best results, use Chrome or Firefox.</h5><![endif]-->\r\n            <div id=\"aroundtable\">\r\n                <table border=\"1\">";
            JTableEnd = "\r\n                </table>\r\n            </div>";
            MHtmlEnd = "\r\n        </div>\r\n    </body>\r\n</html>\r\n";
        }
    }
}
