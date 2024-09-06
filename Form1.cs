using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Fonts;
using System.Diagnostics;
using System.Drawing;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

using System.Drawing.Imaging;

using System.Collections.Generic;
using System.Web;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices.ComTypes;

using Aspose.PSD;
using Aspose.PSD.ImageOptions;

using ImageMagick;
using Aspose.PSD.FileFormats.Ai;

namespace ImageExtensionChangeTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int PsdCounter = 0;
            int PsdErrorCounter = 0;
            int PdfCounter = 0;
            int PdfErrorCounter = 0;

            List<string> PsdCreatePath = new List<string>();
            string PsdCreateText = "";

            string OutPutTxt = "";
            try 
            { 
                string SelectFolder;

                if(textBox1.Text == "")
                {
                    FolderBrowserDialog fbDialog = new FolderBrowserDialog();

                    // ダイアログの説明文を指定する
                    fbDialog.Description = "ダイアログの説明文";

                    // デフォルトのフォルダを指定する
                    fbDialog.SelectedPath = @"C:";

                    // 「新しいフォルダーの作成する」ボタンを表示する
                    fbDialog.ShowNewFolderButton = true;

                    if (fbDialog.ShowDialog() == DialogResult.OK)
                    {
                        SelectFolder = fbDialog.SelectedPath;
                    }
                    else
                    {
                        Console.WriteLine("キャンセルされました");
                        // オブジェクトを破棄する
                        fbDialog.Dispose();
                        return;
                    }
                    fbDialog.Dispose();

                }
                else
                {
                    SelectFolder = textBox1.Text;
                }

                string[] Folders = Directory.GetDirectories(SelectFolder);
                string PdfPsdFlg = "";
                foreach (string path in Folders)
                {
                    try
                    {
                        //files...指定したフォルダの中にあるファイルを全て取り出す
                        string[] files = Directory.GetFiles(path);
                        
                        //その中にあるファイルを画像にする
                        foreach (string filename in files)
                        {
                            try
                            {
                                if (filename.IndexOf(".pdf") > -1)
                                {
                                    continue;
                                    PdfPsdFlg = ".pdf";
                                    PdfCounter++;
                                }
                                else if (filename.IndexOf(".psd") > -1)
                                {
                                    PdfPsdFlg = ".psd";
                                    PsdCounter++;
                                }
                                else if (filename.IndexOf(".ai") > -1)
                                {
                                    PdfPsdFlg = ".ai";
                                    PsdCounter++;
                                }
                                else
                                {
                                    continue;
                                }

                                // 以下の変数にはお好きなファイルパスを入れてください。
                                string pdfPath = filename;
                                string outputPath = filename.Replace(PdfPsdFlg, ".png");
                                if(PdfPsdFlg == ".pdf")
                                {

                                    // PDFをロード
                                    //PdfiumViewer.PdfDocument document = PdfiumViewer.PdfDocument.Load(pdfPath);
                                    using (var document = PdfiumViewer.PdfDocument.Load(pdfPath))
                                    {

                                        //PDFから画像データを作成
                                        var image = document.Render(0, 595, 842, false);
                                        //画像をセーブ
                                        image.Save(outputPath, ImageFormat.Png);
                                        image.Dispose();
                                    }
                                }
                                else if(PdfPsdFlg == ".psd")
                                {
                                    using (MagickImage img = new MagickImage(pdfPath))
                                    {
                                        // PngOptionsクラスのインスタンスを作成します
                                        PngOptions pngOptions = new PngOptions();
                                        // PSDをPNGに変換する
                                        int index = pdfPath.LastIndexOf("\\");
                                        string psdPath = pdfPath.Substring(index, pdfPath.Length - index).Replace(PdfPsdFlg, ".png");
                                        img.Write(@"C:\Users\murakami.k\Desktop\0801_data\切ったデータ\PSD" + psdPath);
                                        img.Dispose();

                                        if (PsdCreatePath.IndexOf(path) == -1)
                                        {
                                            PsdCreatePath.Add(path);
                                            PsdCreateText += "\n" + path;
                                        }

                                        string ProductNo = filename.Substring(filename.LastIndexOf(@"\") + 1, 8);


                                        OutPutTxt += @"UPDATE ToolLists SET ImagePath = '画像まとめ\ツール画像\" + ProductNo + ".png' WHERE ProductNumber = '" + ProductNo + "'\n";

                                        //image.Save(outputPath, pngOptions);
                                    }
                                }
                                else
                                {
                                    using (MagickImage img = new MagickImage(pdfPath))
                                    {
                                        // PngOptionsクラスのインスタンスを作成します
                                        PngOptions pngOptions = new PngOptions();
                                        // PSDをPNGに変換する
                                        int index = pdfPath.LastIndexOf("\\");
                                        string psdPath = pdfPath.Substring(index, pdfPath.Length - index).Replace(PdfPsdFlg, ".png");
                                        img.Write(@"C:\Users\murakami.k\Desktop\0801_data\切ったデータ\PSD" + psdPath);
                                        img.Dispose();

                                        if (PsdCreatePath.IndexOf(path) == -1)
                                        {
                                            PsdCreatePath.Add(path);
                                            PsdCreateText += "\n" + path;
                                        }

                                        string ProductNo = filename.Substring(filename.LastIndexOf(@"\") + 1, 8);


                                        OutPutTxt += @"UPDATE ToolLists SET ImagePath = '画像まとめ\ツール画像\" + ProductNo + ".png' WHERE ProductNumber = '" + ProductNo + "'\n";

                                        //image.Save(outputPath, pngOptions);
                                    }
                                }
                            }
                            catch (Exception error)
                            {
                                if(PdfPsdFlg == ".pdf")
                                {
                                    PdfErrorCounter++;
                                }
                                else if(PdfPsdFlg == ".psd")
                                {
                                    PsdErrorCounter++;
                                }
                            }
                        }
                    }
                    catch(Exception error)
                    {

                    }

                    }
            }
            catch(Exception error)
            {
                MessageBox.Show(error.Message);
                return;
            }
            string OutPutPath = textBox1.Text.Substring(0, textBox1.Text.LastIndexOf("\\"));
            using (StreamWriter sw = new StreamWriter(OutPutPath + "\\Log.txt", false, Encoding.UTF8))
            {
                sw.WriteLine("PDF件数:" + PdfCounter + "\n" + "PDFエラー件数:" + PdfErrorCounter + "\n" + "Psd件数:" + PsdCounter + "\n" + "Psdエラー件数:" + PsdErrorCounter + PsdCreateText + "\n\n" + OutPutTxt);
            }

            MessageBox.Show("PDF件数:" + PdfCounter + " PDFエラー件数:" + PdfErrorCounter + "\n" + "Psd件数:" + PsdCounter + " Psdエラー件数:" + PsdErrorCounter);
            MessageBox.Show("終了しました。");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string SelectFolder;

            if (textBox1.Text == "")
            {
                FolderBrowserDialog fbDialog = new FolderBrowserDialog();

                // ダイアログの説明文を指定する
                fbDialog.Description = "ダイアログの説明文";

                // デフォルトのフォルダを指定する
                fbDialog.SelectedPath = @"C:";

                // 「新しいフォルダーの作成する」ボタンを表示する
                fbDialog.ShowNewFolderButton = true;

                if (fbDialog.ShowDialog() == DialogResult.OK)
                {
                    SelectFolder = fbDialog.SelectedPath;
                }
                else
                {
                    Console.WriteLine("キャンセルされました");
                    // オブジェクトを破棄する
                    fbDialog.Dispose();
                    return;
                }
                fbDialog.Dispose();

            }
            else
            {
                SelectFolder = textBox1.Text;
            }

            string[] Folders = Directory.GetDirectories(SelectFolder);

            
            string SavePath = SelectFolder + "\\" + DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Date.ToString("00")  + DateTime.Now.Hour.ToString("00")  + DateTime.Now.Minute.ToString("00")  + DateTime.Now.Second.ToString("00") + @"\";
            Directory.CreateDirectory(SavePath);

            int DistinctCnt = 0;
            int ImageCounter = 0;
            int NotImageCounter = 0;
            int Counter = 0;

            string OutPutTxt = "";
            foreach (string path in Folders)
            {
                try
                {
                    string[] files = Directory.GetFiles(path + @"\Links");
                    foreach (string filename in files)
                    {
                        Counter++;
                        try
                        {
                            if (filename.IndexOf(".png") == -1 && filename.IndexOf(".jpeg") == -1 && filename.IndexOf(".jpg") == -1 )
                            {
                                NotImageCounter++;
                                continue;
                            }

                            ImageCounter++;

                            string ProductNo = filename.Substring(filename.LastIndexOf(@"\") + 1, 8);

                            File.Copy(filename, SavePath + ProductNo + ".png");

                            OutPutTxt += @"UPDATE ToolLists SET ImagePath = '画像まとめ\ツール画像\" + ProductNo + ".png' WHERE ProductNumber = '" + ProductNo + "'\n";

                        }
                        catch (Exception error)
                        {
                            Console.WriteLine(error);
                            DistinctCnt++;
                        }
                    }
                }
                catch (Exception error)
                {
                    Console.WriteLine(error);
                }

            }

            // テキストファイル出力（新規作成、エンコード指定）
            using (StreamWriter sw = new StreamWriter(SavePath + "Sql.txt", false, Encoding.UTF8))
            {
                sw.WriteLine(OutPutTxt);
            }
            
            using (StreamWriter sw = new StreamWriter(SavePath + "Log.txt", false, Encoding.UTF8))
            {
                sw.WriteLine("重複：" + DistinctCnt + "\nすべてのファイル：" + Counter + "\nその他ファイル：" + NotImageCounter + "\n画像ファイル：" + ImageCounter);
            }

            MessageBox.Show("処理が終了しました。");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbDialog = new FolderBrowserDialog();

            // ダイアログの説明文を指定する
            fbDialog.Description = "ダイアログの説明文";

            // デフォルトのフォルダを指定する
            fbDialog.SelectedPath = @"C:";

            // 「新しいフォルダーの作成する」ボタンを表示する
            fbDialog.ShowNewFolderButton = true;

            if (fbDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fbDialog.SelectedPath;
            }
            else
            {
                Console.WriteLine("キャンセルされました");
            }
            fbDialog.Dispose();
        }
    }
}
