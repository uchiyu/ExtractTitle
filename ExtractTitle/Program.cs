//ref http://kevin3sei.blog95.fc2.com/blog-entry-192.html

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.IO;
using Microsoft.Office.Core;

namespace PptReader
{
    public class PPTReader
    {
        static bool SaveImage = false;

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string pptfile = null;
          
            pptfile = "file path";
            if (pptfile != null && !pptfile.Equals(""))
            {
                SearchFile(pptfile);
            }

        }

        private static void SearchFile(string folder)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = null;
            try
            {
                //Fileオブジェクトを作る
                FileInfo target = new FileInfo(folder);

                //PowerPointの新しいインスタンスを作成する
                app = new Microsoft.Office.Interop.PowerPoint.Application();

                //最小化状態で表示する
                app.Visible = MsoTriState.msoTrue;
                app.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;

                GetPowerPointData(
                    app.Presentations.Open(target.FullName, MsoTriState.msoFalse,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                //PowerPointを終了する
                if (app != null)
                    app.Quit();
            }

        }

        private static void GetPowerPointData(Microsoft.Office.Interop.PowerPoint.Presentation presen)
        {
            //ファイル名の出力
            Console.WriteLine("file:" + presen.FullName);

            Microsoft.Office.Interop.PowerPoint.Slides AllSlide = presen.Slides;
            Microsoft.Office.Interop.PowerPoint.Slide slide = AllSlide[1];
            Microsoft.Office.Interop.PowerPoint.Shapes shapes = slide.Shapes;

            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                string title = shape.TextFrame.TextRange.Text;
                title = title.Replace("\r", "");
                //文字列の出力
                Console.WriteLine(title);

            }
            Console.WriteLine();

            if (SaveImage)
            {
                SaveSlideImage(presen);
            }
            presen.Close();
        }

        private static void SaveSlideImage(Microsoft.Office.Interop.PowerPoint.Presentation presen)
        {
            presen.SaveAs(presen.FullName + "_img",
                Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPNG,
                MsoTriState.msoTrue);
        }
    }
}