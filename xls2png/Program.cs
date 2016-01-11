using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace xls2png
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if(args.Length != 4 && args.Length != 6)
            {
                //                         0        1           2       3          4      5
                Console.WriteLine("xls2png <in.xls> <sheetname> <range> <out.png> [<unit> <addval>]");
                return;
            }
            // open file
            Console.WriteLine("xls2png: please wait...");

            var eapp = new Excel.Application();
            eapp.Visible = false;

            try
            {
                var wb = eapp.Workbooks.Open(args[0]);

                // read sheet
                Excel.Worksheet sht = wb.Sheets[args[1]];
                sht.Select();

                if (args.Length == 6)
                {
                    // set times, will auto calculate
                    sht.get_Range(args[4]).Value = sht.get_Range(args[4]).Value + int.Parse(args[5]);
                }
                
                // copy pic
                Excel.Range rng = sht.get_Range(args[2]);
                rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                // export pic
                var img = Clipboard.GetImage();
                if (img != null) { img.Save(args[3]); img.Dispose(); }
                else { Console.WriteLine("xls2png: image not copied"); }

                // close file                
                wb.Close(true);
                Console.WriteLine("xls2png: done");
            }
            catch (Exception ec) { Console.WriteLine("xls2png: " + ec.GetType() + ": " + ec.Message); }
            finally { eapp.Quit(); }
        }
    }
}
