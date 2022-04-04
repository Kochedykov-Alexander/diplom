using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Text.RegularExpressions;

namespace FileUploadDemo.Controllers
{
    public class FileUploadController : Controller
    {

        // GET: FileUpload
        public ActionResult Index()
        {
            var items = GetFiles();
            return View(items);
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file, string gender, string text_1, string text_2, string text_3, string text_4, string text_5)
        {

            if (file != null && file.ContentLength > 0)
                try
                {
                    string path = Path.Combine(Server.MapPath("~/Images"),
                    Path.GetFileName(file.FileName));
                    file.SaveAs(path);
                    ViewBag.Message = "File uploaded successfully"; 
                    Format(path, gender, text_1, text_2, text_3, text_4, text_5);
                    CreateWordDocument(path, path, text_1, text_2, text_3, text_4, text_5);
                    int wordc = Convert.ToInt32(wordcount(path));
                    int annoc = Convert.ToInt32(annopages(path));
                    int pagec = Convert.ToInt32(pagecount(path));
                    if (gender == "Male")
                    {
                        if (pagec > 20)
                        {
                            ViewBag.Message = "У вас " + pagec + " страниц и вы превысили ограничение в 20 страниц!";
                        }
                        else
                        {
                            ViewBag.Message = "Все требования выполнены!";
                        }
                        // Не больше 10 страниц сделать вывод сообщения ViewBag.Message = "Страниц больше 10"
                    }
                    if (gender == "Female")
                    {
                        if (pagec > 10 && annoc > 3000)
                        {
                            ViewBag.Message = "У вас " + pagec + " страниц и вы превысили ограничение в 10 страниц! " +
                                "У вас " + annoc + " символов в аннотации и вы превысили ограничение в 2000 символов!";
                        }
                        else if (pagec > 10)
                        {
                            ViewBag.Message = "У вас " + pagec + " страниц и вы превысили ограничение в 10 страниц";
                        }
                        else if (annoc > 3000)
                        {
                            ViewBag.Message = "У вас " + annoc + " символов в аннотации и вы превысили ог   раничение в 3000 символов";
                        }
                        else
                        {
                            ViewBag.Message = "Все требования выполнены";
                        }
                    }
                    if (gender == "Shemale")
                    {
                        if (wordc > 6000 && annoc > 2000)
                        {
                            ViewBag.Message = "У вас " + wordc + " символов в тексте и вы превысили ограничение в 6000 символов! " +
                                "У вас " + annoc + " символов в аннотации и вы превысили ограничение в 2000 символов!";
                        }
                        else if (wordc > 6000)
                        {
                            ViewBag.Message = "У вас " + wordc + " символов и вы превысили ограничение в 6000 символов!";
                        }
                        else if (annoc > 2000)
                        {
                            ViewBag.Message = "У вас " + annoc + " символов в аннотации и вы превысили ограничение в 2000 символов!";
                        }
                        else
                        {
                            ViewBag.Message = "Все требования выполнены";
                        }                    
                        // создать еще один файл, который будет сохранять string text_1, string text_2, string text_3, string text_4, string text_5
                    }
                    if (gender == "Tamale")
                    {
                        if (pagec > 12)
                        {
                            ViewBag.Message = "У вас " + pagec + " страниц и вы превысили ограничение в 20 страниц!";
                        }
                        else
                        {
                            ViewBag.Message = "Все требования выполнены!";
                        }
                    }
                }
                catch (Exception ex)
                {
                    ViewBag.Message = "ERROR:" + ex.Message.ToString();
                }
            else
            {
                ViewBag.Message = "You have not specified a file.";
            }


            var items = GetFiles();
            return View(items);
        }

        public static string pagecount(string forpagepath)
        {
            DocumentCore dc = DocumentCore.Load(forpagepath);
            dc.CalculateStats();
            return dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages];
        }






        public static string wordcount(string forpagepath)
        {
            DocumentCore dc = DocumentCore.Load(forpagepath);
            dc.CalculateStats();
            return dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Words];
        }

        public int annopages(string forwordcountanno)
        {
            Application wordApp = new Application();
            Document doc = wordApp.Documents.Open(forwordcountanno, ReadOnly: false);
            wordApp.Visible = false;
            Range r1 = doc.Content;
            Range r2 = doc.Content;
            r1.Find.Execute("Аннотация");
            r2.Find.Execute("Ключевые");
            Range chapter = doc.Range(r1.End, r2.Start);
            var chapter_1 = Regex.Replace(chapter.Text, "[-.?!)(,:]", "");
            //string chapter_1 = chapter.Text.Replace("\r", " ").Replace(" ", "").Replace(" ", "");
            string[] chapter_2 = chapter_1.Split(' ');
            int wordc = chapter_2.Count();
            doc.Close(false);
            wordApp.Application.Quit();
            wordApp.Quit();
            return wordc;
        }

        private void Format(string FileFolder, string selectedgender, string text1, string text2, string text3, string text4, string text5)
        {
            string fullpath = "";
            Application wordApp = new Application();
            Document doc = wordApp.Documents.Open(FileFolder, ReadOnly: false);

            wordApp.Visible = false;
            if (selectedgender == "Male")
            {
                wordApp?.Run("style_1");
                doc?.SaveAs(FileFolder);

                doc.Close(false);
                wordApp.Application.Quit();
                wordApp.Quit();
            }
            if (selectedgender == "Female")
            {
                wordApp?.Run("style_2");
                doc?.SaveAs(FileFolder);

                doc.Close(false);
                wordApp.Application.Quit();
                wordApp.Quit();
            }
            if (selectedgender == "Shemale")
            {
                wordApp?.Run("style_3");
                doc?.SaveAs(FileFolder);

                doc.Close(false);
                wordApp.Application.Quit();
                wordApp.Quit();


                string new_path = FileFolder.Replace(".docx", "");
                Application objword = new Application();
                objword.Visible = false;
                objword.WindowState = WdWindowState.wdWindowStateNormal;
                Document objdoc = objword.Documents.Add();
                fullpath = new_path + " (заявление)" + ".docx";
                objdoc.SaveAs2(fullpath);


                objword?.Run("stylehead");
                objdoc?.SaveAs(fullpath);

                objdoc.Close(false);
                objword.Application.Quit();
                objword.Quit();
                CreateSecond(fullpath, fullpath, text1, text2, text3, text4, text5);
                //CreateWordDocument(fullpath, fullpath, text1, text2, text3, text4, text5);
            }
            if (selectedgender == "Tamale")
            {
                wordApp?.Run("style_4");
                doc?.SaveAs(FileFolder);

                doc.Close(false);
                wordApp.Application.Quit();
                wordApp.Quit();
            }
        }

        public FileResult Download(string ImageName)
        {
            var FileVirtualPath = "~/Images/" + ImageName;
            return File(FileVirtualPath, "application/force- download", Path.GetFileName(FileVirtualPath));
        }

        private List<string> GetFiles()
        {
            var dir = new System.IO.DirectoryInfo(Server.MapPath("~/Images"));
            System.IO.FileInfo[] fileNames = dir.GetFiles("*.*");

            List<string> items = new List<string>();
            foreach (var file in fileNames)
            {
                items.Add(file.Name);
            }


            return items;
        }

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object toFindText, object replaceWithText)
        {
            object matchCase = true;

            object matchwholeWord = true;

            object matchwildCards = false;

            object matchSoundLike = false;

            object nmatchAllforms = false;

            object forward = true;

            object format = false;

            object matchKashida = false;

            object matchDiactitics = false;

            object matchAlefHamza = false;

            object matchControl = false;

            object read_only = false;

            object visible = true;

            object replace = -2;

            object wrap = 1;

            wordApp.Selection.Find.Execute(ref toFindText, ref matchCase,
                                            ref matchwholeWord, ref matchwildCards, ref matchSoundLike,

                                            ref nmatchAllforms, ref forward,

                                            ref wrap, ref format, ref replaceWithText,

                                                ref replace, ref matchKashida,

                                            ref matchDiactitics, ref matchAlefHamza,

                                             ref matchControl);
        }


        private void CreateWordDocument(object filename, object SaveAs, string text1, string text2, string text3, string text4, string text5)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;

            Microsoft.Office.Interop.Word.Document myWordDoc = null;

            if (System.IO.File.Exists((string)filename))
            {
                object readOnly = false;

                object isvisible = false;

                wordApp.Visible = false;
                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                     ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();
                this.FindAndReplace(wordApp, "afdsafsdafas", text1);
                this.FindAndReplace(wordApp, "dfshgdfhfgs", text2);
                this.FindAndReplace(wordApp, "afdsagfsgvr", text3);
                this.FindAndReplace(wordApp, "vrvvvrewrcec", text4);
                this.FindAndReplace(wordApp, "rtbytrysd", text5);
                Range r1 = myWordDoc.Content;
                Range r2 = myWordDoc.Content;
                r1.Find.Execute("Ключевые слова");
                r2.Find.Execute("Второй");
                Range chapter = myWordDoc.Range(r1.End, r2.Start);   
                string chapter_1 = chapter.Text.Replace(".", "").Replace("\r", "").Trim();
                string[] chapter_2 = chapter_1.Split(',');
                string chapter_3 = String.Join(",", chapter_2, 0, 6);
                this.FindAndReplace(wordApp, chapter_1, chapter_3);
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
            }
        }


        private void CreateSecond(object filename, object SaveAs, string text1, string text2, string text3, string text4, string text5)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;

            Microsoft.Office.Interop.Word.Document myWordDoc = null;

            if (System.IO.File.Exists((string)filename))
            {
                object readOnly = false;

                object isvisible = false;

                wordApp.Visible = false;
                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                     ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();
                this.FindAndReplace(wordApp, "afdsafsdafas", text1);
                this.FindAndReplace(wordApp, "dfshgdfhfgs", text2);
                this.FindAndReplace(wordApp, "afdsagfsgvr", text3);
                this.FindAndReplace(wordApp, "vrvvvrewrcec", text4);
                this.FindAndReplace(wordApp, "rtbytrysd", text5);
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
            }
        }

    }
}
