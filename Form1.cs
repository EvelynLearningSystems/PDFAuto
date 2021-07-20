using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Documents;
using SD= Spire.Doc.Documents;
using System.Text.RegularExpressions;
using System.IO;
using TextBox = Spire.Doc.Fields.TextBox;
using System.Drawing.Imaging;
using System.Net;

namespace PDFAuto
{
    public partial class Form1 : Form
    {

        Dictionary<string, string> answers = new Dictionary<string, string>();
        Dictionary<int, string> headings = new Dictionary<int, string>();
        int HeadingsCounter = 0;
        string LoadFilePath = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadFiles();
        }

        public void LoadFiles()
        {
            string FILE_NAME = "headings.txt";
            string currLine = "";
            int counter = 0;

            if (System.IO.File.Exists(FILE_NAME) == true)
            {
                System.IO.StreamReader objReader = new System.IO.StreamReader(FILE_NAME);

                while (objReader.Peek() != -1)
                {
                    currLine = objReader.ReadLine();
                    headings.Add(counter, currLine);
                    counter += 1;
                }
                objReader.Dispose();
            }

            HeadingsCounter = counter;
        }
        private void loadDocBtn_Click(object sender, EventArgs e)
        {
            LoadDocFile();

            extractImages();

            imageTags();

            readContents();

            LoadContent();

            CleanContent();

            TableOutput();
            //extractImages();
            //ReadContentFromDoc1();


            MessageBox.Show("Done");
        }

        public void LoadDocFile()
        {
            if (loadDocFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileBox.Text = loadDocFileDialog.FileName;
                LoadFilePath = loadDocFileDialog.FileName;
                //contentBox.AppendText(Path.GetDirectoryName(LoadFilePath) + Environment.NewLine);
                //contentBox.AppendText(Path.GetDirectoryName(Path.GetDirectoryName(LoadFilePath)));

            }
        }

        public void ReadContentFromDoc()
        {
            string LoadFileName = @"C:\Users\C K Bhushan\Desktop\NCERT Books\Exemplar.docx";
            Document doc = new Document();
            doc.LoadFromFile(LoadFileName);

            StringBuilder sb = new StringBuilder();
            foreach (Section section in doc.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    DocumentObject obj = section.Body.ChildObjects[i];
                    if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];
                            if (cobj.DocumentObjectType == DocumentObjectType.Shape)
                            {
                                ShapeObject shape = cobj as ShapeObject;
                                for (int m = 0; m < shape.ChildObjects.Count; m++)
                                {
                                    if (shape.ChildObjects[m].DocumentObjectType == DocumentObjectType.Paragraph)
                                    {
                                        Paragraph para = shape.ChildObjects[m] as Paragraph;
                                        for (int n = 0; n < para.ChildObjects.Count; n++)
                                        {
                                            if (para.ChildObjects[n].DocumentObjectType == DocumentObjectType.TextRange)
                                            {
                                                TextRange range = para.ChildObjects[n] as TextRange;
                                                string text = range.Text;
                                                bool isBold = range.CharacterFormat.Bold;
                                                //bool isItalic = range.CharacterFormat.Italic;
                                                //UnderlineStyle underlineStyle = range.CharacterFormat.UnderlineStyle;
                                                sb.AppendLine(text);
                                                // sb.AppendLine("Is bold: " + isBold);
                                                //sb.AppendLine("Is italic: " + isItalic);
                                                //sb.AppendLine("Underline style: " + underlineStyle.ToString());
                                            }
                                        }
                                        sb.AppendLine();
                                    }
                                }
                            }
                        }
                    }
                    if (obj is Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];
                            if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange range = cobj as TextRange;
                                string text = range.Text;
                                bool isBold = range.CharacterFormat.Bold;
                                //bool isItalic = range.CharacterFormat.Italic;
                                //UnderlineStyle underlineStyle = range.CharacterFormat.UnderlineStyle;
                                sb.AppendLine(text);
                                //other code...
                                // sb.Append((cobj as TextRange).Text);
                            }
                            if (cobj.DocumentObjectType == DocumentObjectType.Picture)
                            {
                                //other code...
                            }
                            if (cobj.DocumentObjectType == DocumentObjectType.ShapeGroup)
                            {
                                ShapeGroup shapeGroup = cobj as ShapeGroup;
                                for (int k = 0; k < shapeGroup.ChildObjects.Count; k++)
                                {
                                    //Console.WriteLine(shapeGroup.ChildObjects[k].DocumentObjectType);
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Picture)
                                    {
                                        //other code...
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Shape)
                                    {
                                        ShapeObject shape = cobj as ShapeObject;
                                        for (int m = 0; m < shape.ChildObjects.Count; m++)
                                        {
                                            if (shape.ChildObjects[m].DocumentObjectType == DocumentObjectType.Paragraph)
                                            {
                                                Paragraph para = shape.ChildObjects[m] as Paragraph;
                                                for (int n = 0; n < para.ChildObjects.Count; n++)
                                                {
                                                    if (para.ChildObjects[n].DocumentObjectType == DocumentObjectType.TextRange)
                                                    {
                                                        TextRange range = para.ChildObjects[n] as TextRange;
                                                        string text = range.Text;
                                                        bool isBold = range.CharacterFormat.Bold;
                                                        //bool isItalic = range.CharacterFormat.Italic;
                                                        //UnderlineStyle underlineStyle = range.CharacterFormat.UnderlineStyle;
                                                        //if(isBold)
                                                        //{
                                                        //    sb.AppendLine("<bold>" + text + "<bold>");
                                                        //}
                                                        //else
                                                        //{
                                                        //    sb.AppendLine(text);
                                                        //}
                                                        sb.AppendLine(text);
                                                        //sb.AppendLine("Is bold: " + isBold);
                                                        //sb.AppendLine("Is italic: " + isItalic);
                                                        //sb.AppendLine("Underline style: " + underlineStyle.ToString());
                                                    }
                                                }
                                                sb.AppendLine();
                                            }
                                        }
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Table)
                                    {
                                        Table table = shapeGroup.ChildObjects[k] as Table;
                                        ExtractTextFromTables(table, sb);
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.TextBox)
                                    {

                                        TextBox textbox = shapeGroup.ChildObjects[k] as TextBox;
                                        foreach (DocumentObject objt in textbox.ChildObjects)
                                        {
                                            Console.WriteLine(objt.DocumentObjectType);
                                            //Extract text from paragraph in TextBox.
                                            if (objt.DocumentObjectType == DocumentObjectType.Paragraph)
                                            {
                                                sb.AppendLine((objt as Paragraph).Text);
                                            }
                                            if (objt.DocumentObjectType == DocumentObjectType.Table)
                                            {
                                                Table table = objt as Table;
                                                ExtractTextFromTables(table, sb);
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        sb.AppendLine();
                    }
                    if (obj is Table)
                    {
                        Table table = obj as Table;
                        ExtractTextFromTables(table, sb);
                    }
                }
            }

            contentBox.Text = sb.ToString();


            doc.Close();
        }

        public void ReadContentFromDoc1()
        {
            string LoadFileName = @"C:\Users\C K Bhushan\Desktop\NCERT Books\Exemplar.docx";
            Document doc = new Document();
            doc.LoadFromFile(LoadFileName);

            // StringBuilder sb = new StringBuilder();
            foreach (Section section in doc.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    DocumentObject obj = section.Body.ChildObjects[i];

                    if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];
                            if (cobj.DocumentObjectType == DocumentObjectType.Shape)
                            {
                                ShapeObject shape = cobj as ShapeObject;
                                for (int m = 0; m < shape.ChildObjects.Count; m++)
                                {
                                    if (shape.ChildObjects[m].DocumentObjectType == DocumentObjectType.Paragraph)
                                    {
                                        Paragraph para = shape.ChildObjects[m] as Paragraph;
                                        for (int n = 0; n < para.ChildObjects.Count; n++)
                                        {
                                            if (para.ChildObjects[n].DocumentObjectType == DocumentObjectType.TextRange)
                                            {
                                                TextRange range = para.ChildObjects[n] as TextRange;
                                                string text = range.Text;
                                                bool isBold = range.CharacterFormat.Bold;
                                                //bool isItalic = range.CharacterFormat.Italic;
                                                //UnderlineStyle underlineStyle = range.CharacterFormat.UnderlineStyle;
                                                // sb.Append(text);
                                                // sb.AppendLine("Is bold: " + isBold);
                                                //sb.AppendLine("Is italic: " + isItalic);
                                                //sb.AppendLine("Underline style: " + underlineStyle.ToString());
                                                contentBox.AppendText(text);
                                            }
                                        }
                                        // sb.AppendLine();
                                    }
                                }
                            }

                            if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange range = cobj as TextRange;
                                string text = range.Text;
                                bool isBold = range.CharacterFormat.Bold;
                                //bool isItalic = range.CharacterFormat.Italic;
                                //UnderlineStyle underlineStyle = range.CharacterFormat.UnderlineStyle;
                                // sb.Append(text);
                                //other code...
                                // sb.Append((cobj as TextRange).Text);
                                contentBox.AppendText(text);

                            }

                            //  sb.AppendLine();


                        }
                        contentBox.AppendText(Environment.NewLine);
                    }

                }
            }
            //contentBox.Text = sb.ToString();


            doc.Close();
        }

        public void imageTags()
        {
            var sImgName = Path.GetFileNameWithoutExtension(LoadFilePath);
            string imagePath = sImgName + "-images/";
            Document doc = new Document();
            doc.LoadFromFile(LoadFilePath);
            int i = 1;
            foreach (Section sec in doc.Sections)
            {
                foreach (Paragraph para in sec.Paragraphs)
                {
                    List<DocumentObject> pictures = new List<DocumentObject>();
                    List<DocumentObject> oleObjects = new List<DocumentObject>();
                    foreach (DocumentObject dobjt in para.ChildObjects)
                    {
                        if (dobjt.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            pictures.Add(dobjt);
                        }
                        //if(dobjt.DocumentObjectType == DocumentObjectType.OleObject)
                        // {
                        //     oleObjects.Add(dobjt);
                        // }
                    }
                    foreach (DocumentObject pic in pictures)
                    {
                        int index = para.ChildObjects.IndexOf(pic);
                        TextRange range = new TextRange(doc);
                        //string imgTextReplace = @" < img src=""images/image001.jpg"" alt=""images""/>";
                        range.Text = string.Format(@"<img>https://repository.evelynlearning.com/imageresource/" + imagePath + sImgName + "-image00" + i.ToString() + @".png</img>");
                        para.ChildObjects.Insert(index, range);
                        para.ChildObjects.Remove(pic);
                        i++;
                    }
                }
            }

            doc.SaveToFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-1.docx", FileFormat.Docx);



        }

        public void readContents()
        {
            Document doc = new Document();

            doc.LoadFromFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-1.docx");

            // doc.LoadFromFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Exemplar.docx");

            Document doc2 = new Document();

            Section s2 = doc2.AddSection();
            s2.PageSetup.PageSize = PageSize.A4;

            doc2.Sections[0].PageSetup.Margins.Top = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Bottom = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Left = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Right = 10.9f;

            Boolean qstatus = false;
            int optstatus = 0;
            Boolean headStatus = false;
            Boolean isBold = false;
            string currLine = "";

            foreach (Section section in doc.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    DocumentObject obj = section.Body.ChildObjects[i];

                    if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;

                        isBold = false;
                        currLine = "";
                        foreach (DocumentObject docobj in paragraph.ChildObjects)
                        {
                            if (docobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange text = docobj as TextRange;

                                if (text.CharacterFormat.Bold)
                                {
                                    isBold = true;
                                }
                                else
                                {
                                    isBold = false;
                                }
                            }
                        }

                        // contentBox.AppendText(paragraph.ListText);
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];

                            if (cobj.DocumentObjectType == DocumentObjectType.Shape)
                            {
                                ShapeObject shape = cobj as ShapeObject;
                                for (int m = 0; m < shape.ChildObjects.Count; m++)
                                {
                                    if (shape.ChildObjects[m].DocumentObjectType == DocumentObjectType.Paragraph)
                                    {
                                        Paragraph para = shape.ChildObjects[m] as Paragraph;
                                        for (int n = 0; n < para.ChildObjects.Count; n++)
                                        {
                                            if (para.ChildObjects[n].DocumentObjectType == DocumentObjectType.TextRange)
                                            {
                                                TextRange range = para.ChildObjects[n] as TextRange;
                                                string text = range.Text;
                                                isBold = range.CharacterFormat.Bold;
                                                currLine += text;
                                                // contentBox.AppendText(para.ListText + " " + text);

                                                // questionsBox.AppendText(text + Environment.NewLine);

                                            }
                                        }
                                    }
                                }
                            }

                            if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange range = cobj as TextRange;
                                string text = range.Text;
                                //bool isBold = range.CharacterFormat.Bold;
                                currLine += text;
                                // contentBox.AppendText(text);

                            }

                            if (cobj.DocumentObjectType == DocumentObjectType.Table)
                            {
                                Table table = cobj as Table;
                                for (int a = 0; a < table.Rows.Count; a++)
                                {
                                    TableRow row = table.Rows[a];
                                    for (int b = 0; b < row.Cells.Count; b++)
                                    {
                                        TableCell cell = row.Cells[b];
                                        foreach (Paragraph para in cell.Paragraphs)
                                        {
                                            //currLine += para.Text;
                                            //  contentBox.AppendText(para.ListText + " " + para.Text);
                                        }
                                    }
                                }
                            }

                            if (cobj.DocumentObjectType == DocumentObjectType.ShapeGroup)
                            {
                                ShapeGroup shapeGroup = cobj as ShapeGroup;
                                for (int k = 0; k < shapeGroup.ChildObjects.Count; k++)
                                {
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Table)
                                    {
                                        Table table = shapeGroup.ChildObjects[k] as Table;
                                        //Table table = cobj as Table;
                                        for (int a = 0; a < table.Rows.Count; a++)
                                        {
                                            TableRow row = table.Rows[a];
                                            for (int b = 0; b < row.Cells.Count; b++)
                                            {
                                                TableCell cell = row.Cells[b];
                                                foreach (Paragraph para in cell.Paragraphs)
                                                {
                                                    //  contentBox.AppendText(para.ListText + " " + para.Text);
                                                }
                                            }
                                        }
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.TextBox)
                                    {
                                        TextBox textbox = shapeGroup.ChildObjects[k] as TextBox;
                                        foreach (DocumentObject objt in textbox.ChildObjects)
                                        {
                                            Console.WriteLine(objt.DocumentObjectType);
                                            //Extract text from paragraph in TextBox.
                                            if (objt.DocumentObjectType == DocumentObjectType.Paragraph)
                                            {
                                                Paragraph para = objt as Paragraph;
                                                //questionsBox.AppendText(para.Text);

                                                isBold = true;

                                                // contentBox.AppendText(para.ListText + para.Text);
                                                currLine += para.Text;
                                            }
                                            if (objt.DocumentObjectType == DocumentObjectType.Table)
                                            {
                                                Table table = objt as Table;
                                                // Table table = cobj as Table;
                                                for (int a = 0; a < table.Rows.Count; a++)
                                                {
                                                    TableRow row = table.Rows[a];
                                                    for (int b = 0; b < row.Cells.Count; b++)
                                                    {
                                                        TableCell cell = row.Cells[b];
                                                        foreach (Paragraph para in cell.Paragraphs)
                                                        {
                                                            //  contentBox.AppendText(para.ListText + para.Text);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        //if(currLine.Trim().Length >1)
                                        //{
                                        //    foreach (KeyValuePair<int, string> entry in headings)
                                        //    {
                                        //        if (currLine.Contains(entry.Value.Trim()))  // .Key must be capitalized
                                        //        {
                                        //            questionsBox.AppendText(currLine + Environment.NewLine);
                                        //        }
                                        //    }
                                        //}

                                    }
                                }
                            }

                        } // outer for loop
                          // contentBox.AppendText(Environment.NewLine);
                        if (isBold)
                        {

                            //headStatus = false;
                            Boolean intStatus = false;
                            foreach (KeyValuePair<int, string> entry in headings)
                            {
                                if (currLine.Contains(entry.Value.Trim()))  // .Key must be capitalized
                                {
                                    intStatus = true;
                                }
                            }
                            if (intStatus == true)
                            {
                                headStatus = true;
                            }
                            else
                            {
                                headStatus = false;
                            }


                            if (currLine.Contains("In questions"))
                            {
                                headStatus = true;
                            }
                            else if (currLine.Contains("EXERCISE"))
                            {
                                headStatus = true;
                            }
                            else if (currLine.Contains("In each of the questions"))
                            {
                                headStatus = true;
                            }
                            else if (currLine.StartsWith("State whether the statements"))
                            {
                                headStatus = true;
                            }
                            // headingsBox.AppendText(currLine + ">>" + headStatus + Environment.NewLine);

                        }

                        if (headStatus == true)
                        {
                            //headingsBox.AppendText(headStatus.ToString() + Environment.NewLine);
                            Paragraph parag = obj as Paragraph;
                            // questionsBox.AppendText(parag.Text + Environment.NewLine);
                            if (!String.IsNullOrEmpty(parag.Text.Trim()))
                            {
                                Paragraph para1 = (Paragraph)parag.Clone();
                                para1.Format.LeftIndent = 30;

                                // para1.Format.ClearFormatting();
                                para1.Format.HorizontalAlignment = SD.HorizontalAlignment.Left;
                                s2.Paragraphs.Add(para1);
                            }

                        }

                    } //if paragraph

                }
            }



            doc2.SaveToFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.docx", FileFormat.Docx);

            //doc2.SaveToFile(@"C: \Users\C K Bhushan\Desktop\NCERT Books\Exemplar- phase-2.docx", FileFormat.Docx);
            doc2.Close();
            doc.Close();
        }

        public void readContents1()
        {
            Document doc = new Document();

            doc.LoadFromFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-1.docx");

            Document doc2 = new Document();

            Section s2 = doc2.AddSection();
            s2.PageSetup.PageSize = PageSize.A4;

            doc2.Sections[0].PageSetup.Margins.Top = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Bottom = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Left = 10.9f;
            doc2.Sections[0].PageSetup.Margins.Right = 10.9f;

            int qstatus = 0;
            int optstatus = 0;
            Boolean headStatus = false;
            Boolean isBold = false;

            foreach (Section sec in doc.Sections)
            {
                foreach (Paragraph par in sec.Paragraphs)
                {
                    //par.Format.HorizontalAlignment
                    if (par.Text.Trim() == "ANSWERS")
                    {
                        break;
                    }

                    isBold = false;
                    foreach (DocumentObject docobj in par.ChildObjects)
                    {
                        if (docobj.DocumentObjectType == DocumentObjectType.TextRange)
                        {
                            TextRange text = docobj as TextRange;

                            if (text.CharacterFormat.Bold)
                            {
                                isBold = true;
                            }
                            else
                            {
                                isBold = false;
                            }
                        }
                    }
                    // headStatus = false;
                    if (isBold)
                    {
                        if (par.Format.HorizontalAlignment == SD.HorizontalAlignment.Left)
                        {
                            if (!Regex.IsMatch(par.Text.Trim(), @"^\d+$")) //&& par.Text.Trim().Any(char.IsLower)
                            {

                                if (par.Text.Contains("In questions"))
                                {
                                    headStatus = true;
                                    allHeadingsBox.AppendText(par.Text + Environment.NewLine);
                                }
                                else if (par.Text.Contains("EXERCISE"))
                                {
                                    headStatus = true;
                                    allHeadingsBox.AppendText(par.Text + Environment.NewLine);
                                }
                                else if (par.Text.Contains("In each of the questions"))
                                {
                                    headStatus = true;
                                    allHeadingsBox.AppendText(par.Text + Environment.NewLine);
                                }
                                else if (par.Text.StartsWith("State whether the statements"))
                                {
                                    headStatus = true;
                                    allHeadingsBox.AppendText(par.Text + Environment.NewLine);
                                }
                                else
                                {
                                    headStatus = false;
                                }


                                // allHeadingsBox.AppendText(par.Text + Environment.NewLine);
                                headingsBox.AppendText(par.Text + ">>" + Environment.NewLine + headStatus + Environment.NewLine);
                            }
                            else
                            {
                                headStatus = false;
                            }
                        }


                    }

                    if (headStatus == true)
                    {
                        if (par.Text.Replace(" ", "").Any(char.IsLetterOrDigit))
                        {

                            // Paragraph NewPara1 = (Paragraph)par.Clone();

                            if (!String.IsNullOrEmpty(par.Text.Trim()))
                            {
                                Paragraph para1 = (Paragraph)par.Clone();
                                para1.Format.LeftIndent = 30;

                                // para1.Format.ClearFormatting();
                                para1.Format.HorizontalAlignment = SD.HorizontalAlignment.Left;
                                s2.Paragraphs.Add(para1);
                            }




                            questionsBox.AppendText(par.ListText + par.Text + Environment.NewLine);
                        }

                    }

                    string style = par.StyleName;


                    contentBox.AppendText(par.ListText + par.Text + Environment.NewLine);




                }
            }

            doc2.SaveToFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.docx", FileFormat.Docx);
            doc2.Close();
            doc.Close();
        }

        public void extractImages()
        {
            int fileCount = 0;
            string imageName;
            //string imageDir = subjectBox.SelectedItem.ToString();


            //int i = 0;
            try
            {
                fileCount = fileCount + 1;

                Document doc = new Document();
                doc.LoadFromFile(LoadFilePath);
                List<DocPicture> DocPictureList = new List<DocPicture>();
                List<DocPicture> mathPictureList = new List<DocPicture>();
                List<DocPicture> image = new List<DocPicture>();
                image.Clear();
                DocPictureList.Clear();
                mathPictureList.Clear();
                var sImgName = Path.GetFileNameWithoutExtension(LoadFilePath);
                string imagePath = Path.GetDirectoryName(LoadFilePath) + "/" + sImgName + "_images/";
                if (!Directory.Exists(imagePath))
                {
                    Directory.CreateDirectory(imagePath);
                }

                //Loop through contents
                foreach (Section section in doc.Sections)
                {
                    foreach (DocumentObject obj in section.Body.ChildObjects)
                    {
                        if (obj is Paragraph)
                        {
                            Paragraph para = obj as Paragraph;
                            foreach (DocumentObject cobj in para.ChildObjects)
                            {
                                //Find DocPicture object and add it in DocPictureList
                                if (cobj is DocPicture)
                                {
                                    DocPicture pic = cobj as DocPicture;
                                    DocPictureList.Add(pic);

                                }
                                //Find DocOleObject object and add it in mathPictureList
                                if (cobj is DocOleObject)
                                {
                                    DocOleObject ole = cobj as DocOleObject;
                                    mathPictureList.Add(ole.OlePicture);
                                }
                            }
                            image = DocPictureList.Except(mathPictureList).ToList();
                            //image = mathPictureList;

                        }
                    }

                    int imageIdx = 1;
                    foreach (DocPicture pic in DocPictureList)
                    {
                        // textBox1.AppendText(pic + Environment.NewLine);
                        //imageName = string.Format(sImgName + "_" + DateTime.Now.ToString("HH_mm_ss") + "_image00{0}.png", imageIdx);

                        imageName = string.Format(sImgName + "_image00{0}.png", imageIdx);
                        
                        pic.Image.Save(imagePath + imageName, System.Drawing.Imaging.ImageFormat.Png);
                        imageIdx += 1;

                    }

                }

                image.Clear();
                DocPictureList.Clear();
                mathPictureList.Clear();
                doc.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // IssueBox1.AppendText("File is not in expected format" + " " + ex + " " + Environment.NewLine);
            }
            // MessageBox.Show("Done Files");

        }


        public static void Image()
        {
            Document doc = new Document();
            doc.LoadFromFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\Images.docx");

            int imageIdx = 1;
            string imageName = "";
            foreach (Section section in doc.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    DocumentObject obj = section.Body.ChildObjects[i];
                    if (obj is Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];
                            if (cobj.DocumentObjectType == DocumentObjectType.Picture)
                            {
                                //other code...
                                DocPicture pic = cobj as DocPicture;

                                imageName = string.Format("Test" + "_image00{0}.png", imageIdx);
                                pic.Image.Save(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\images\" + imageName, System.Drawing.Imaging.ImageFormat.Png);
                                imageIdx += 1;
                            }
                            if (cobj.DocumentObjectType == DocumentObjectType.ShapeGroup)
                            {
                                ShapeGroup shapeGroup = cobj as ShapeGroup;

                                for (int k = 0; k < shapeGroup.ChildObjects.Count; k++)
                                {
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Picture)
                                    {
                                        //other code...
                                        DocPicture pic = shapeGroup.ChildObjects[k] as DocPicture;
                                        imageName = string.Format("Test" + "_image00{0}.png", imageIdx);
                                        pic.Image.Save(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\images\" + imageName, System.Drawing.Imaging.ImageFormat.Png);
                                        imageIdx += 1;

                                    }
                                }
                            }
                        }
                    }
                }
            }
            doc.SaveToFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\out.docx");
        }

        // image end

        static void Table()
        {
            Document document = new Document();
            document.LoadFromFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\Anchored.docx");

            foreach (Section section in document.Sections)
            {
                for (int i = 0; i < section.Body.Tables.Count; i++)
                {
                    Paragraph para = new Paragraph(document);
                    Table table = section.Body.Tables[i] as Table;
                    int index = section.Body.ChildObjects.IndexOf(table);
                    Image image = ConvertTableToImage(table);
                    para.AppendPicture(image);
                    section.Body.ChildObjects.Insert(index, para);
                    section.Body.ChildObjects.Remove(table);
                    i--;
                }
            }

            document.SaveToFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\tableout.docx");
        }


        private static Image ConvertTableToImage(Table obj)
        {
            Document doc = new Document();
            Section section = doc.AddSection();

            section.Body.ChildObjects.Add(obj.Clone());
            Image image = doc.SaveToImages(0, ImageType.Bitmap);
            doc.Close();
            return CutImageWhitePart(image as Bitmap, 1);
        }


        public static Image CutImageWhitePart(Bitmap bmp, int WhiteBarRate)
        {
            int top = 0, left = 0;
            int right = bmp.Width, bottom = bmp.Height;
            Color white = Color.White;

            for (int i = 0; i < bmp.Height; i++)
            {
                bool find = false;
                for (int j = 0; j < bmp.Width; j++)
                {
                    Color c = bmp.GetPixel(j, i);
                    if (IsWhite(c))
                    {
                        top = i;
                        find = true;
                        break;
                    }
                }
                if (find) break;
            }

            for (int i = 0; i < bmp.Width; i++)
            {
                bool find = false;
                for (int j = top; j < bmp.Height; j++)
                {
                    Color c = bmp.GetPixel(i, j);
                    if (IsWhite(c))
                    {
                        left = i;
                        find = true;
                        break;
                    }
                }
                if (find) break; ;
            }

            for (int i = bmp.Height - 1; i >= 0; i--)
            {
                bool find = false;
                for (int j = left; j < bmp.Width; j++)
                {
                    Color c = bmp.GetPixel(j, i);
                    if (IsWhite(c))
                    {
                        bottom = i;
                        find = true;
                        break;
                    }
                }
                if (find) break;
            }

            for (int i = bmp.Width - 1; i >= 0; i--)
            {
                bool find = false;
                for (int j = 0; j < bottom; j++)
                {
                    Color c = bmp.GetPixel(i, j);
                    if (IsWhite(c))
                    {
                        right = i;
                        find = true;
                        break;
                    }
                }
                if (find) break;
            }
            int iWidth = right - left;
            int iHeight = bottom - left;
            int blockWidth = Convert.ToInt32(iWidth * WhiteBarRate / 100);
            bmp = Cut(bmp, left - blockWidth, top - blockWidth, right - left + 2 * blockWidth, bottom - top + 2 * blockWidth);

            return bmp;
        }
        public static Bitmap Cut(Bitmap b, int StartX, int StartY, int iWidth, int iHeight)
        {
            if (b == null)
            {
                return null;
            }
            int w = b.Width;
            int h = b.Height;
            if (StartX >= w || StartY >= h)
            {
                return null;
            }
            if (StartX + iWidth > w)
            {
                iWidth = w - StartX;
            }
            if (StartY + iHeight > h)
            {
                iHeight = h - StartY;
            }
            try
            {
                Bitmap bmpOut = new Bitmap(iWidth, iHeight, PixelFormat.Format24bppRgb);
                Graphics g = Graphics.FromImage(bmpOut);
                g.DrawImage(b, new Rectangle(0, 0, iWidth, iHeight), new Rectangle(StartX, StartY, iWidth, iHeight), GraphicsUnit.Pixel);
                g.Dispose();
                return bmpOut;
            }
            catch
            {
                return null;
            }
        }
        public static bool IsWhite(Color c)
        {
            if (c.R < 245 || c.G < 245 || c.B < 245)
                return true;
            else return false;
        }

        // table end

        public static void Anchored()
        {
            Document doc = new Document();
            doc.LoadFromFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\Anchored.docx");
            StringBuilder sb = new StringBuilder();
            foreach (Section section in doc.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    DocumentObject obj = section.Body.ChildObjects[i];
                    if (obj is Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];
                            if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                //other code...
                                sb.Append((cobj as TextRange).Text);
                            }
                            if (cobj.DocumentObjectType == DocumentObjectType.Picture)
                            {
                                //other code...
                            }
                            if (cobj.DocumentObjectType == DocumentObjectType.ShapeGroup)
                            {
                                ShapeGroup shapeGroup = cobj as ShapeGroup;
                                for (int k = 0; k < shapeGroup.ChildObjects.Count; k++)
                                {
                                    //Console.WriteLine(shapeGroup.ChildObjects[k].DocumentObjectType);
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Picture)
                                    {
                                        //other code...
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Shape)
                                    {
                                        //other code...
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.Table)
                                    {
                                        Table table = shapeGroup.ChildObjects[k] as Table;
                                        ExtractTextFromTables(table, sb);
                                    }
                                    if (shapeGroup.ChildObjects[k].DocumentObjectType == DocumentObjectType.TextBox)
                                    {

                                        TextBox textbox = shapeGroup.ChildObjects[k] as TextBox;
                                        foreach (DocumentObject objt in textbox.ChildObjects)
                                        {
                                            Console.WriteLine(objt.DocumentObjectType);
                                            //Extract text from paragraph in TextBox.
                                            if (objt.DocumentObjectType == DocumentObjectType.Paragraph)
                                            {
                                                sb.AppendLine((objt as Paragraph).Text);
                                            }
                                            if (objt.DocumentObjectType == DocumentObjectType.Table)
                                            {
                                                Table table = objt as Table;
                                                ExtractTextFromTables(table, sb);
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        sb.AppendLine();
                    }
                    if (obj is Table)
                    {
                        Table table = obj as Table;
                        ExtractTextFromTables(table, sb);
                    }
                }
            }
            doc.SaveToFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\out.docx");
            File.WriteAllText(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\out.txt", sb.ToString());
        }

        static void ExtractTextFromTables(Table table, StringBuilder sb)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                TableRow row = table.Rows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    TableCell cell = row.Cells[j];
                    foreach (Paragraph paragraph in cell.Paragraphs)
                    {
                        sb.AppendLine(paragraph.Text);
                    }
                }
            }
        }

        public static void ShapeGroupExtraction()
        {
            Document doc = new Document();
            doc.LoadFromFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Exemplar.docx");
            foreach (Section section in doc.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    DocumentObject obj = section.Body.ChildObjects[i];
                    if (obj is Paragraph)
                    {
                        Paragraph paragraph = obj as Paragraph;
                        for (int j = 0; j < paragraph.ChildObjects.Count; j++)
                        {
                            DocumentObject cobj = paragraph.ChildObjects[j];
                            if (cobj.DocumentObjectType == DocumentObjectType.ShapeGroup)
                            {
                                ShapeGroup shapeGroup = cobj as ShapeGroup;

                                Image image = ConvertShapeGroupToImage(shapeGroup);
                                DocPicture pic = new DocPicture(doc);
                                pic.TextWrappingStyle = shapeGroup.TextWrappingStyle;
                                pic.TextWrappingType = shapeGroup.TextWrappingType;
                                pic.VerticalOrigin = shapeGroup.VerticalOrigin;
                                pic.HorizontalOrigin = shapeGroup.HorizontalOrigin;
                                pic.VerticalPosition = shapeGroup.VerticalPosition;
                                pic.HorizontalPosition = shapeGroup.HorizontalPosition;
                                pic.LoadImage(image);
                                paragraph.ChildObjects.Insert(j, pic);
                                paragraph.ChildObjects.Remove(shapeGroup);
                            }
                        }
                    }
                }
            }
            doc.SaveToFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Issues\image-out1.docx");
        }

        private static Image ConvertShapeGroupToImage(ShapeGroup obj)
        {
            Document doc = new Document();
            Section section = doc.AddSection();
            section.AddParagraph().Items.Add(obj.Clone());
            Image image = doc.SaveToImages(0, ImageType.Bitmap);
            doc.Close();
            return CutImageWhitePart(image as Bitmap, 1);
        }





        private void testButton_Click(object sender, EventArgs e)
        {
            ShowHeadings();

            //RemoveWatermark();

            // Anchored();

            // Table();

            //Image();

            // ShapeGroupExtraction();

            // FTPImageUpload();

            //UploadFileToFTP();

            ExtractAnswers();

            // dispKeyVal();

            LoadContent();

            CleanContent();

            TableOutput();

            //InsertPageNumber();


            MessageBox.Show("Done");
        }

        public void ShowHeadings()
        {

            // MessageBox.Show(DateTime.Now.ToString("HH_mm_ss"));
            using (var tw = new StreamWriter("headings.txt", true))
            {
                for (int i = 0; i < NewHeadingsBox.Lines.Count(); i++)
                {
                    if (NewHeadingsBox.Lines[i].ToString().Length > 1)
                    {
                        tw.WriteLine(NewHeadingsBox.Lines[i].ToString());
                        headings.Add(HeadingsCounter++, NewHeadingsBox.Lines[i].ToString());
                    }

                }
            }

            //foreach (KeyValuePair<int, string> p in headings)
            //{
            //    contentBox.AppendText(p.Value + Environment.NewLine);

            //}
        }

        public void RemoveWatermark()
        {
            Document doc = new Document();
            doc.LoadFromFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Class VI\Mathematics\Exemplar\Exemplar VI AIO.docx", FileFormat.Docx);

            Section section = doc.Sections[0];
            //This is necessary
            section.PageSetup.DifferentFirstPageHeaderFooter = true;
            section.HeadersFooters.FirstPageHeader.ChildObjects.Clear();
            section.HeadersFooters.Header.ChildObjects.Clear();
            doc.SaveToFile("output.docx", FileFormat.Docx);
            System.Diagnostics.Process.Start("output.docx");

            //  doc.SaveToFile(@"C:\Users\C K Bhushan\Desktop\NCERT Books\Class VI\Mathematics\Exemplar\Exemplar VI Header NW.docx", FileFormat.Docx);

        }


        public void dispKeyVal()
        {
            foreach (var item in answers)
            {
                headingsBox.AppendText(item.Key + "--" + item.Value + Environment.NewLine);
                // MessageBox.Show(item.Key + "   " + item.Value);
            }
        }

        public void InsertPageNumber()
        {

            // Document doc = new Document();
            // doc.LoadFromFile(@"D:\Test\Exemplar.docx");

            Document document = new Document();
            document.LoadFromFile(@"D:\Test\Exemplar.docx");
            Section section = document.Sections[0];

            section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = true;

            Paragraph P1 = section.HeadersFooters.EvenFooter.AddParagraph();
            TextRange EF = P1.AppendText("Even Footer Demo from E-iceblue Using Spire.Doc");
            EF.CharacterFormat.FontName = "Calibri";
            EF.CharacterFormat.FontSize = 20;
            EF.CharacterFormat.TextColor = Color.Green;
            EF.CharacterFormat.Bold = true;
            P1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;


            Paragraph P2 = section.HeadersFooters.OddFooter.AddParagraph();
            TextRange OF = P2.AppendText("Odd Footer Demo");
            P2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            OF.CharacterFormat.FontName = "Calibri";
            OF.CharacterFormat.FontSize = 20;
            OF.CharacterFormat.Bold = true;
            OF.CharacterFormat.TextColor = Color.Blue;

            Paragraph P3 = section.HeadersFooters.OddHeader.AddParagraph();
            TextRange OH = P3.AppendText("Odd Header Demo");
            P3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            OH.CharacterFormat.FontName = "Calibri";
            OH.CharacterFormat.FontSize = 20;
            OH.CharacterFormat.Bold = true;
            OH.CharacterFormat.TextColor = Color.Blue;

            Paragraph P4 = section.HeadersFooters.EvenHeader.AddParagraph();
            TextRange EH = P4.AppendText("Even Header Demo from E-iceblue Using Spire.Doc");
            P4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            EH.CharacterFormat.FontName = "Calibri";
            EH.CharacterFormat.FontSize = 20;
            EH.CharacterFormat.Bold = true;
            EH.CharacterFormat.TextColor = Color.Green;

            document.SaveToFile(@"D:\Test\Exemplar-output.docx", FileFormat.Docx);

            document.Close();

        }

        public void TableOutput()
        {
            Document document = new Document();
            Section section = document.AddSection();
            string correctOptions = "";
            string marksIfTrue = "1";
            string marksIfFalse = "0";
            string isMandatory = "FALSE";
            string questionType = "WORD";
            string createdBy = "PDF Auto";
            string isHidden = "FALSE";
            string currQuestionType = string.Empty;
            int questionNumber = 1;
            string courses = "CBSE";
            string subjects = "Mathematics";
            string tags = "CBSE, Mathematics, Class VI";
            string correctAnswer = "";
            string globalExplanation = "";
            string currentLine = "";
            int tableCount = -1;
            int rowCount = 1;
            Boolean questionStatus = false;
            int lower = 1;
            int upper = 1;
            Boolean tokenstatus = false;
            string skey = "";

            string FILE_NAME = Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-3.txt";


            if (System.IO.File.Exists(FILE_NAME) == true)
            {
                System.IO.StreamReader objReader = new System.IO.StreamReader(FILE_NAME);

                while (objReader.Peek() != -1)
                {
                    currentLine = objReader.ReadLine();

                    if (currentLine.Trim().Length < 1)
                    {
                        continue;
                    }

                    string first1 = currentLine.Split('.').First();

                    if (first1.All(char.IsDigit))
                    {
                        try
                        {
                            questionNumber = Convert.ToInt32(first1);
                        }
                        catch (Exception ex)
                        {

                            // Console.WriteLine(ex.Message);
                        }

                    }

                    if (currentLine.StartsWith("[") && currentLine.Contains("UNIT"))
                    {
                        skey = currentLine.Replace("[", "").Replace("]", "");
                        if (answers.ContainsKey(skey))
                        {
                            correctAnswer = answers[skey];
                        }
                        continue;
                    }

                    if (currentLine.Contains("[MCQ]"))
                    {
                        questionType = "MCQ";
                        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
                        {
                            lower = numValue;
                        }
                        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
                        {
                            upper = numValue2;
                        }
                        continue;

                    }
                    else if (currentLine.Contains("[TF]"))
                    {
                        questionType = "TF";
                        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
                        {
                            lower = numValue;
                        }
                        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
                        {
                            upper = numValue2;
                        }
                        continue;
                    }
                    else if (currentLine.Contains("[FIB]"))
                    {
                        questionType = "FIB";

                        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
                        {
                            lower = numValue;
                        }
                        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
                        {
                            upper = numValue2;
                        }
                        continue;
                    }


                    if (questionNumber >= lower && questionNumber <= upper)
                    {
                        currQuestionType = questionType;
                    }
                    else
                    {
                        currQuestionType = "WORD";
                    }

                    //string first = currentLine.Split('.').First();
                    //if(first.All(char.IsDigit))
                    //{
                    //    questionStatus = true;
                    //    continue;
                    //}
                    if (currentLine.Contains("<question>"))
                    {
                        questionStatus = true;
                        tokenstatus = true;
                        continue;
                    }

                    if (questionStatus == true)
                    {
                        section.AddParagraph();
                        Table table = section.AddTable(true);

                        // table.ResetCells(2, 2);
                        rowCount = 1;
                        table.ResetCells(11, 2);
                        table.Rows[2].Cells[1].SplitCell(2, 1);
                        table[1, 0].AddParagraph().AppendText("Question Type");
                        table[1, 1].AddParagraph().AppendText(currQuestionType);
                        table[2, 0].AddParagraph().AppendText("Marks");
                        table[2, 1].AddParagraph().AppendText(marksIfTrue);
                        table[2, 2].AddParagraph().AppendText(marksIfFalse);
                        table[3, 0].AddParagraph().AppendText("isMandatory");
                        table[3, 1].AddParagraph().AppendText(isMandatory);
                        table[4, 0].AddParagraph().AppendText("isHidden");
                        table[4, 1].AddParagraph().AppendText(isHidden);
                        table[5, 0].AddParagraph().AppendText("correctAnswer");
                        table[5, 1].AddParagraph().AppendText(correctAnswer);
                        table[6, 0].AddParagraph().AppendText("globalExplanation");
                        table[6, 1].AddParagraph().AppendText(globalExplanation);
                        table[7, 0].AddParagraph().AppendText("createdBy");
                        table[7, 1].AddParagraph().AppendText(createdBy);
                        table[8, 0].AddParagraph().AppendText("subjects");
                        table[8, 1].AddParagraph().AppendText(subjects);
                        table[9, 0].AddParagraph().AppendText("courses");
                        table[9, 1].AddParagraph().AppendText(courses);
                        table[10, 0].AddParagraph().AppendText("tags");
                        table[10, 1].AddParagraph().AppendText(tags);
                        //table[10, 0].AddParagraph().AppendText();
                        //table[10, 1].AddParagraph().AppendText();

                        //table.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                        questionStatus = false;
                        tableCount += 1;
                    }

                    if (tableCount >= 0)
                    {
                        Table ctable = document.Sections[0].Tables[tableCount] as Spire.Doc.Table;
                        ctable.ApplyHorizontalMerge(0, 0, 1);
                        //ctable.Rows[6].Cells[1].SplitCell(2, 1);
                        ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                        if (currQuestionType == "MCQ")
                        {
                            if (currentLine.StartsWith("(a)") || currentLine.StartsWith("(A)"))// && (questionType !="SA" || questionType != "LA" || questionType != "FILEUPLOAD" || questionType != "FILEUPLOAD-D"))
                            {
                                currentLine = currentLine.Replace("(a)", "").Trim();
                                currentLine = currentLine.Replace("(A)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(2, 1);
                                ctable.Rows[rowCount].Cells[2].SplitCell(2, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);

                                if (correctAnswer.Trim() == "(A)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);


                            }
                            else if (currentLine.StartsWith("(b)") || currentLine.StartsWith("(B)"))
                            {
                                currentLine = currentLine.Replace("(b)", "").Trim();
                                currentLine = currentLine.Replace("(B)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(B)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(c)") || currentLine.StartsWith("(C)"))
                            {
                                currentLine = currentLine.Replace("(c)", "").Trim();
                                currentLine = currentLine.Replace("(C)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(C)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(d)") || currentLine.StartsWith("(D)"))
                            {
                                currentLine = currentLine.Replace("(d)", "").Trim();
                                currentLine = currentLine.Replace("(D)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(D)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(e)") || currentLine.StartsWith("(E)"))
                            {
                                currentLine = currentLine.Replace("(e)", "").Trim();
                                currentLine = currentLine.Replace("(E)", "").Trim();
                                TableRow row = ctable.AddRow();
                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(E)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else if (currentLine.StartsWith("(f)") || currentLine.StartsWith("(F)"))
                            {
                                currentLine = currentLine.Replace("(f)", "").Trim();
                                currentLine = currentLine.Replace("(F)", "").Trim();
                                TableRow row = ctable.AddRow();

                                rowCount += 1;
                                ctable.Rows.Insert(rowCount, row);
                                ctable.Rows[rowCount].Cells[1].SplitCell(2, 1);
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
                                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
                                if (correctAnswer.Trim() == "(F)")
                                {
                                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                                }
                                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

                            }
                            else
                            {
                                string first = currentLine.Split('.').First();
                                if (first.All(char.IsDigit))
                                {
                                    currentLine = currentLine.Replace(first + ".", "");
                                }
                                ctable[0, 0].AddParagraph().AppendText(currentLine);

                            }


                        }
                        else if (currQuestionType == "TF")
                        {
                            TableRow row = ctable.AddRow();
                            rowCount += 1;
                            ctable.Rows.Insert(rowCount, row);
                            ctable[rowCount, 0].AddParagraph().AppendText("True");
                            if (correctAnswer == "T")
                            {
                                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                            }

                            TableRow row2 = ctable.AddRow();
                            rowCount += 1;
                            ctable.Rows.Insert(rowCount, row2);
                            ctable[rowCount, 0].AddParagraph().AppendText("False");
                            if (correctAnswer == "F")
                            {
                                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
                            }

                            string first = currentLine.Split('.').First();
                            if (first.All(char.IsDigit))
                            {
                                currentLine = currentLine.Replace(first + ".", "");
                            }
                            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
                            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
                            ctable[0, 0].AddParagraph().AppendText(currentLine);


                        }
                        else if (currQuestionType == "FIB")
                        {
                            string[] tokens;

                            if (correctAnswer.Contains(";"))
                            {
                                tokens = correctAnswer.Split(';');
                            }
                            else if (correctAnswer.Contains("(a)"))
                            {
                                string answer = correctAnswer.Trim().Replace("(a)", "").Replace("(b)", "-").Replace("(c)", "-").Replace("(d)", "-").Replace("(e)", "-");
                                tokens = answer.Split('-');
                            }
                            else
                            {
                                if (correctAnswer.Contains(", "))
                                {
                                    var ans = correctAnswer.Replace(", ", "-");
                                    tokens = ans.Split('-');
                                }
                                else
                                {
                                    tokens = new string[] { correctAnswer };
                                }

                            }

                            if (tokenstatus == true)
                            {
                                int tok = 1;
                                foreach (string item in tokens)
                                {
                                    TableRow row = ctable.AddRow();
                                    rowCount += 1;
                                    ctable.Rows.Insert(rowCount, row);
                                    ctable[rowCount, 0].AddParagraph().AppendText("token" + tok.ToString());
                                    ctable[rowCount, 1].AddParagraph().AppendText(item);
                                    tok += 1;
                                }
                                tokenstatus = false;
                            }

                            string first = currentLine.Split('.').First();
                            if (first.All(char.IsDigit))
                            {
                                currentLine = currentLine.Replace(first + ".", "");
                            }

                            if (currentLine.Trim().Length > 0)
                            {
                                if (!currentLine.Contains("token"))
                                {

                                    if (currentLine.EndsWith("."))
                                    {
                                        currentLine = currentLine.TrimEnd('.') + " [token1].";
                                    }
                                    else
                                    {
                                        currentLine = currentLine + " [token1].";
                                    }
                                }
                                ctable[0, 0].AddParagraph().AppendText(currentLine);
                            }

                            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
                            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');


                        }
                        else
                        {
                            string first = currentLine.Split('.').First();
                            if (first.All(char.IsDigit))
                            {
                                currentLine = currentLine.Replace(first + ".", "");
                            }
                            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
                            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
                            ctable[0, 0].AddParagraph().AppendText(currentLine);
                        }

                    }


                }
                objReader.Dispose();
            } 
            


            //for (int i = 0; i < contentBox.Lines.Count() - 1; i++)
            //{
            //    currentLine = contentBox.Lines[i].ToString();
               
            //    if(currentLine.Trim().Length <1)
            //    {
            //        continue;
            //    }

            //    string first1 = currentLine.Split('.').First();

            //    if (first1.All(char.IsDigit))
            //    {
            //        try
            //        {
            //            questionNumber = Convert.ToInt32(first1);
            //        }
            //        catch (Exception ex)
            //        {

            //           // Console.WriteLine(ex.Message);
            //        }
                    
            //    }

            //    if(currentLine.StartsWith("[") && currentLine.Contains("UNIT"))
            //    {
            //        skey = currentLine.Replace("[", "").Replace("]", "");
            //        if (answers.ContainsKey(skey))
            //        {
            //            correctAnswer = answers[skey];
            //        }
            //        continue;
            //    }

            //    if (currentLine.Contains("[MCQ]"))
            //    {
            //        questionType = "MCQ";
            //        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
            //        {
            //            lower = numValue;
            //        }
            //        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
            //        {
            //            upper = numValue2;
            //        }
            //        continue;

            //    }
            //    else if (currentLine.Contains("[TF]"))
            //    {
            //        questionType = "TF";
            //        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
            //        {
            //            lower = numValue;
            //        }
            //        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
            //        {
            //            upper = numValue2;
            //        }
            //        continue;
            //    }
            //    else if (currentLine.Contains("[FIB]"))
            //    {
            //        questionType = "FIB";

            //        if (Int32.TryParse(currentLine.Split('-')[1], out int numValue))
            //        {
            //            lower = numValue;
            //        }
            //        if (Int32.TryParse(currentLine.Split('-')[2], out int numValue2))
            //        {
            //            upper = numValue2;
            //        }
            //        continue;
            //    }
               

            //    if(questionNumber >= lower && questionNumber <= upper)
            //    {
            //        currQuestionType = questionType;
            //    }
            //    else
            //    {
            //        currQuestionType = "WORD";
            //    }

            //    //string first = currentLine.Split('.').First();
            //    //if(first.All(char.IsDigit))
            //    //{
            //    //    questionStatus = true;
            //    //    continue;
            //    //}
            //    if (currentLine.Contains("<question>"))
            //    {
            //        questionStatus = true;
            //        tokenstatus = true;
            //        continue;
            //    }

            //    if (questionStatus == true)
            //    {
            //        section.AddParagraph();
            //        Table table = section.AddTable(true);
                    
            //        // table.ResetCells(2, 2);
            //        rowCount = 1;
            //        table.ResetCells(11, 2);
            //        table.Rows[2].Cells[1].SplitCell(2, 1);
            //        table[1, 0].AddParagraph().AppendText("Question Type");
            //        table[1, 1].AddParagraph().AppendText(currQuestionType);
            //        table[2, 0].AddParagraph().AppendText("Marks");
            //        table[2, 1].AddParagraph().AppendText(marksIfTrue);
            //        table[2, 2].AddParagraph().AppendText(marksIfFalse);
            //        table[3, 0].AddParagraph().AppendText("isMandatory");
            //        table[3, 1].AddParagraph().AppendText(isMandatory);
            //        table[4, 0].AddParagraph().AppendText("isHidden");
            //        table[4, 1].AddParagraph().AppendText(isHidden);
            //        table[5, 0].AddParagraph().AppendText("correctAnswer");
            //        table[5, 1].AddParagraph().AppendText(correctAnswer);
            //        table[6, 0].AddParagraph().AppendText("globalExplanation");
            //        table[6, 1].AddParagraph().AppendText(globalExplanation);
            //        table[7, 0].AddParagraph().AppendText("createdBy");
            //        table[7, 1].AddParagraph().AppendText(createdBy);
            //        table[8, 0].AddParagraph().AppendText("subjects");
            //        table[8, 1].AddParagraph().AppendText(subjects);
            //        table[9, 0].AddParagraph().AppendText("courses");
            //        table[9, 1].AddParagraph().AppendText(courses);
            //        table[10, 0].AddParagraph().AppendText("tags");
            //        table[10, 1].AddParagraph().AppendText(tags);
            //        //table[10, 0].AddParagraph().AppendText();
            //        //table[10, 1].AddParagraph().AppendText();

            //        //table.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //        questionStatus = false;
            //        tableCount += 1;
            //    }

            //    if (tableCount >= 0)
            //    {
            //        Table ctable = document.Sections[0].Tables[tableCount] as Spire.Doc.Table;
            //        ctable.ApplyHorizontalMerge(0, 0, 1);
            //        //ctable.Rows[6].Cells[1].SplitCell(2, 1);
            //        ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //        if(currQuestionType == "MCQ")
            //        {
            //            if (currentLine.StartsWith("(a)") || currentLine.StartsWith("(A)"))// && (questionType !="SA" || questionType != "LA" || questionType != "FILEUPLOAD" || questionType != "FILEUPLOAD-D"))
            //            {
            //                currentLine = currentLine.Replace("(a)", "").Trim();
            //                currentLine = currentLine.Replace("(A)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(2,1);
            //                ctable.Rows[rowCount].Cells[2].SplitCell(2, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);

            //                if(correctAnswer.Trim() == "(A)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);


            //            }
            //            else if (currentLine.StartsWith("(b)") || currentLine.StartsWith("(B)"))
            //            {
            //                currentLine = currentLine.Replace("(b)", "").Trim();
            //                currentLine = currentLine.Replace("(B)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(B)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(c)") || currentLine.StartsWith("(C)"))
            //            {
            //                currentLine = currentLine.Replace("(c)", "").Trim();
            //                currentLine = currentLine.Replace("(C)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(C)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(d)") || currentLine.StartsWith("(D)"))
            //            {
            //                currentLine = currentLine.Replace("(d)", "").Trim();
            //                currentLine = currentLine.Replace("(D)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(D)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(e)") || currentLine.StartsWith("(E)"))
            //            {
            //                currentLine = currentLine.Replace("(e)", "").Trim();
            //                currentLine = currentLine.Replace("(E)", "").Trim();
            //                TableRow row = ctable.AddRow();
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(3, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(E)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else if (currentLine.StartsWith("(f)") || currentLine.StartsWith("(F)"))
            //            {
            //                currentLine = currentLine.Replace("(f)", "").Trim();
            //                currentLine = currentLine.Replace("(F)", "").Trim();
            //                TableRow row = ctable.AddRow();
                           
            //                rowCount += 1;
            //                ctable.Rows.Insert(rowCount, row);
            //                ctable.Rows[rowCount].Cells[1].SplitCell(2, 1);
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
            //                ctable[rowCount, 0].AddParagraph().AppendText(currentLine);
            //                if (correctAnswer.Trim() == "(F)")
            //                {
            //                    ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //                }
            //                ctable.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //            }
            //            else
            //            {
            //                string first = currentLine.Split('.').First();
            //                if (first.All(char.IsDigit))
            //                {
            //                    currentLine = currentLine.Replace(first + ".", "");
            //                }
            //                ctable[0, 0].AddParagraph().AppendText(currentLine);

            //            }
                        

            //        }
            //        else if(currQuestionType == "TF")
            //        {
            //            TableRow row = ctable.AddRow();
            //            rowCount += 1;
            //            ctable.Rows.Insert(rowCount, row);
            //            ctable[rowCount, 0].AddParagraph().AppendText("True");
            //            if (correctAnswer == "T")
            //            {
            //                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //            }

            //            TableRow row2 = ctable.AddRow();
            //            rowCount += 1;
            //            ctable.Rows.Insert(rowCount, row2);
            //            ctable[rowCount, 0].AddParagraph().AppendText("False");
            //            if (correctAnswer == "F")
            //            {
            //                ctable[rowCount, 1].AddParagraph().AppendText("TRUE");
            //            }

            //            string first = currentLine.Split('.').First();
            //            if (first.All(char.IsDigit))
            //            {
            //                currentLine = currentLine.Replace(first + ".", "");
            //            }
            //            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
            //            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
            //            ctable[0, 0].AddParagraph().AppendText(currentLine);


            //        }
            //        else if(currQuestionType == "FIB")
            //        {
            //            string[] tokens;

            //            if(correctAnswer.Contains(";"))
            //            {
            //                tokens = correctAnswer.Split(';');
            //            }
            //            else if(correctAnswer.Contains("(a)"))
            //            {
            //                string answer = correctAnswer.Trim().Replace("(a)", "").Replace("(b)", "-").Replace("(c)", "-").Replace("(d)", "-").Replace("(e)", "-");
            //                tokens = answer.Split('-');
            //            }
            //            else
            //            {
            //                if(correctAnswer.Contains(", "))
            //                {
            //                    var ans = correctAnswer.Replace(", ", "-");
            //                    tokens = ans.Split('-');
            //                }
            //                else
            //                {
            //                    tokens =  new string[] { correctAnswer };
            //                }
                            
            //            }
     
            //            if (tokenstatus == true)
            //            {
            //                int tok = 1;
            //                foreach (string item in tokens)
            //                {
            //                    TableRow row = ctable.AddRow();
            //                    rowCount += 1;
            //                    ctable.Rows.Insert(rowCount, row);
            //                    ctable[rowCount, 0].AddParagraph().AppendText("token" + tok.ToString());
            //                    ctable[rowCount, 1].AddParagraph().AppendText(item);
            //                    tok += 1;
            //                }
            //                tokenstatus = false;
            //            }
                        
            //            string first = currentLine.Split('.').First();
            //            if (first.All(char.IsDigit))
            //            {
            //                currentLine = currentLine.Replace(first + ".", "");
            //            }

            //            if (currentLine.Trim().Length > 0)
            //            {
            //                if (!currentLine.Contains("token"))
            //                {

            //                    if (currentLine.EndsWith("."))
            //                    {
            //                        currentLine = currentLine.TrimEnd('.') + " [token1].";
            //                    }
            //                    else
            //                    {
            //                        currentLine = currentLine + " [token1].";
            //                    }
            //                }
            //                ctable[0, 0].AddParagraph().AppendText(currentLine);
            //            }
                       
            //            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
            //            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
                        

            //        }
            //        else 
            //        {
            //            string first = currentLine.Split('.').First();
            //            if (first.All(char.IsDigit))
            //            {
            //                currentLine = currentLine.Replace(first + ".", "");
            //            }
            //            //currentLine = currentLine.Replace(questionNumber + ".", "").Trim();
            //            //currentLine = currentLine.Replace(questionNumber, "").Trim().TrimStart('.');
            //            ctable[0, 0].AddParagraph().AppendText(currentLine);
            //        }

            //    }


            //}


            document.SaveToFile(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-final.docx", FileFormat.Docx);
            document.Close();

        }
        public void CleanContent()
        {
            //contentBox.Clear();
            string isMandatory = "";
            string questionType = "";
            string questionNumber = "";
            string currentLine = "";
            int tableCount = -1;
            Boolean questionStatus = false;

            string currentQuestionNo;
            int currentQN = 1;
            int qn = 1;
            int currUnit = 1;
            Boolean unitstatus = true;
            int tokenCount = 1;
            string currLine = "";

            string FILE_NAME = Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.txt";

            using (StreamWriter Wr = new StreamWriter(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-3.txt"))
            {

            if (System.IO.File.Exists(FILE_NAME) == true)
            {
                System.IO.StreamReader objReader = new System.IO.StreamReader(FILE_NAME);

                while (objReader.Peek() != -1)
                {
                    currentLine = objReader.ReadLine();

                    currentQuestionNo = currentLine.Split('.')[0];
                    if (currentQuestionNo.Trim().All(char.IsDigit))
                    {
                        try
                        {
                            currentQN = Convert.ToInt32(currentQuestionNo);

                            if (currentQN > 400 || currentQN < 1)
                            {
                                continue;
                            }

                            if (currentQN >= qn)
                            {
                                qn = currentQN;
                                // unitstatus = false;
                            }
                            else
                            {
                                qn = currentQN;
                                currUnit += 1;
                                // unitstatus = true;


                            }

                        }
                        catch (Exception)
                        {

                            continue;
                        }


                            Wr.Write(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "-" + currentQN + "]" + Environment.NewLine);
                        //contentBox.AppendText(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "-" + currentQN + "]" + Environment.NewLine);
                        tokenCount = 1;
                    }

                    string pattern = @"[_]{2,}";
                    Regex regex = new Regex(pattern);
                    foreach (Match ItemMatch in regex.Matches(currentLine))
                    {
                        currentLine = currentLine.Replace(ItemMatch.Value, "_");

                    }


                    int freq = currentLine.Count(f => (f == '_'));

                    for (int fi = 1; fi <= freq; fi++)
                    {
                        string repStr = " [token" + tokenCount.ToString() + "] ";
                        var regex_replace = new Regex(Regex.Escape("_"));

                        currentLine = regex_replace.Replace(currentLine, repStr, 1);
                        tokenCount += 1;
                    }

                    //string first = currentLine.Split('.').First();
                    //if(first.All(char.IsDigit))
                    //{
                    //    textBox1.AppendText(first + Environment.NewLine);
                    //}
                    string first = currentLine.Split('.').First();
                    if (first.All(char.IsDigit))
                    {
                        currentLine = "<question>\\n" + currentLine;
                    }

                    currentLine = currentLine.Replace("(a)", "\\n(a)").Replace("(b)", "\\n(b)").Replace("(c)", "\\n(c)").Replace("(d)", "\\n(d)").Replace("(e)", "\\n(e)").Replace("(f)", "\\n(f)");
                    currentLine = currentLine.Replace("(A)", "\\n(A)").Replace("(B)", "\\n(B)").Replace("(C)", "\\n(C)").Replace("(D)", "\\n(D)").Replace("(E)", "\\n(E)").Replace("(F)", "\\n(F)");

                    var result = currentLine.Split(new string[] { "\\n" }, StringSplitOptions.None);
                    foreach (string s in result)
                    {
                       // contentBox.AppendText(s + Environment.NewLine);
                            Wr.Write(s + Environment.NewLine);
                    }






                }
                objReader.Dispose();
            }

        }


            //for (int i = 0; i < headingsBox.Lines.Count() - 1; i++)
            //{
            //    currentLine = headingsBox.Lines[i].ToString();

            //    //if (unitstatus == true)
            //    //{
            //    //    contentBox.AppendText(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "]" + Environment.NewLine);
            //    //}

            //    currentQuestionNo = currentLine.Split('.')[0];
            //    if (currentQuestionNo.Trim().All(char.IsDigit))
            //    {
            //        try
            //        {
            //            currentQN = Convert.ToInt32(currentQuestionNo);

            //            if (currentQN > 400 || currentQN < 1)
            //            {
            //                continue;
            //            }

            //            if (currentQN >= qn)
            //            {
            //                qn = currentQN;
            //               // unitstatus = false;
            //            }
            //            else
            //            {
            //                qn = currentQN;
            //                currUnit += 1;
            //               // unitstatus = true;
                            
                            
            //            }

            //        }
            //        catch (Exception)
            //        {

            //            continue;
            //        }

                   

            //        contentBox.AppendText(Environment.NewLine + "[UNIT" + "-" + currUnit.ToString() + "-" + currentQN + "]" + Environment.NewLine);
            //        tokenCount = 1;
            //    }

            //    string pattern = @"[_]{2,}";
            //    Regex regex = new Regex(pattern);
            //    foreach (Match ItemMatch in regex.Matches(currentLine))
            //    {
            //        currentLine = currentLine.Replace(ItemMatch.Value, "_" );
                    
            //    }

               
            //    int freq = currentLine.Count(f => (f == '_'));

            //    for(int fi =1; fi <= freq; fi++)
            //    {
            //        string repStr = " [token" + tokenCount.ToString() + "] ";
            //        var regex_replace = new Regex(Regex.Escape("_"));

            //        currentLine = regex_replace.Replace(currentLine, repStr, 1);
            //        tokenCount += 1;
            //    }

            //    //string first = currentLine.Split('.').First();
            //    //if(first.All(char.IsDigit))
            //    //{
            //    //    textBox1.AppendText(first + Environment.NewLine);
            //    //}
            //    string first = currentLine.Split('.').First();
            //    if (first.All(char.IsDigit))
            //    {
            //        currentLine = "<question>\\n" + currentLine;
            //    }

            //    currentLine = currentLine.Replace("(a)", "\\n(a)").Replace("(b)", "\\n(b)").Replace("(c)", "\\n(c)").Replace("(d)", "\\n(d)").Replace("(e)", "\\n(e)").Replace("(f)", "\\n(f)");
            //    currentLine = currentLine.Replace("(A)", "\\n(A)").Replace("(B)", "\\n(B)").Replace("(C)", "\\n(C)").Replace("(D)", "\\n(D)").Replace("(E)", "\\n(E)").Replace("(F)", "\\n(F)");

            //    var result = currentLine.Split(new string[] { "\\n" }, StringSplitOptions.None);
            //    foreach (string s in result)
            //    {
            //        contentBox.AppendText(s + Environment.NewLine);
            //    }

            //}


        }


        public void LoadContent()
        {
            string fileName1 = Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.docx";
            //string fileName1 = @"C:\Users\C K Bhushan\Desktop\NCERT Books\Tools\Input\Questions11.docx";


            Boolean isBold = false;

            Document doc = new Document();
            doc.LoadFromFile(fileName1);
            using (StreamWriter Wr = new StreamWriter(Path.GetDirectoryName(LoadFilePath) + "/" + Path.GetFileNameWithoutExtension(LoadFilePath) + "-phase-2.txt"))
            { 
            foreach (Section sec in doc.Sections)
            {
                foreach (Paragraph para in sec.Paragraphs)
                {

                    isBold = false;
                    foreach (DocumentObject docobj in para.ChildObjects)
                    {
                        if (docobj.DocumentObjectType == DocumentObjectType.TextRange)
                        {
                            TextRange text = docobj as TextRange;

                            if (text.CharacterFormat.Bold)
                            {
                                isBold = true;
                            }
                            else
                            {
                                isBold = false;
                            }
                            if (text.CharacterFormat.UnderlineStyle == UnderlineStyle.Single)
                            {
                                text.Text = "_";
                            }
                        }
                    }

                    if (isBold)
                    {
                       // allHeadingsBox.AppendText(para.Text + Environment.NewLine);

                        if (para.Text.Contains("out of the four options") || para.Text.Contains("out of four options") || para.Text.Contains("only one of the four options") || para.Text.Contains("out of the given four options"))
                        {
                           // headingsBox.AppendText("[MCQ]");
                            Wr.Write("[MCQ]");
                            string input = para.Text;
                            // Split on one or more non-digit characters.
                            string[] numbers = Regex.Split(input, @"\D+");
                            foreach (string value in numbers)
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    //int i = int.Parse(value);
                                   // headingsBox.AppendText("-" + value);
                                    Wr.Write("-" + value);

                                }
                            }
                                Wr.WriteLine();
                            //headingsBox.AppendText(Environment.NewLine);

                        }
                        else if (para.Text.Contains("true or false") || para.Text.Contains("true (T) or false (F)") || para.Text.Contains("T or F") || para.Text.Contains("(T) or (F)"))
                        {
                            //headingsBox.AppendText("[TF]");
                                Wr.Write("[TF]");
                            string input = para.Text;
                            // Split on one or more non-digit characters.
                            string[] numbers = Regex.Split(input, @"\D+");
                            foreach (string value in numbers)
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    //int i = int.Parse(value);
                                   // headingsBox.AppendText("-" + value);
                                        Wr.Write("-" + value);

                                }
                            }
                           // headingsBox.AppendText(Environment.NewLine);
                                Wr.WriteLine();
                        }
                        else if (para.Text.Contains("fill in the blanks"))
                        {
                            headingsBox.AppendText("[FIB]");
                            string input = para.Text;
                            // Split on one or more non-digit characters.
                            string[] numbers = Regex.Split(input, @"\D+");
                            foreach (string value in numbers)
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    //int i = int.Parse(value);
                                   // headingsBox.AppendText("-" + value);
                                        Wr.Write("-" + value);

                                    }
                            }
                                //headingsBox.AppendText(Environment.NewLine);
                                Wr.WriteLine();
                        }
                        //else
                        //{
                        //    headingsBox.AppendText("[WORD]" + Environment.NewLine);
                        //}

                    }
                    else
                    {
                            // headingsBox.AppendText(para.ListText + para.Text + Environment.NewLine);
                            Wr.Write(para.ListText+para.Text);
                            Wr.WriteLine();
                    }

                }
            }
                Wr.Close();
            }
           
        
        }

        public void ExtractAnswers()
        {
            string filename = @"C:\Users\C K Bhushan\Desktop\NCERT Books\Class VI\Mathematics\Exemplar\answers.docx";
            Document doc = new Document();
            doc.LoadFromFile(filename);
            //textBox1.AppendText(Path.GetFileName(filename) + Environment.NewLine);

            //string latexfile = Path.GetFileName(filename).Replace(".docx", ".txt");
            int imgIndex = 1;

            contentBox.Clear();

            foreach (Section section in doc.Sections)
            {
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    //if it is table
                    if (obj is Table)
                    {
                        Table table = obj as Table;
                        foreach (TableRow row in table.Rows)
                        {
                            foreach (TableCell cell in row.Cells)
                            {

                                foreach (DocumentObject nobj in cell.ChildObjects)
                                {
                                    if (nobj is Table)
                                    {
                                        Table ntable = nobj as Table;
                                        foreach (TableRow nrow in ntable.Rows)
                                        {
                                            foreach (TableCell ncell in nrow.Cells)
                                            {
                                                //traverse paragraphs in table cells
                                                foreach (Paragraph para in ncell.Paragraphs)
                                                {
                                                    //contentBox.AppendText(para.ListText +  para.Text + Environment.NewLine);
                                                    if (para.ListText.Length > 0)
                                                    {
                                                       contentBox.AppendText(Environment.NewLine + para.ListText);
                                                    }
                                                    // + para.Text + Environment.NewLine);

                                                    for (int i = 0; i < para.ChildObjects.Count; i++)
                                                    {
                                                        DocumentObject cobj = para.ChildObjects[i];

                                                        if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                                                        {
                                                            TextRange text = cobj as TextRange;

                                                            if (text.Text.Trim().Length > 0)
                                                            {
                                                                if (text.CharacterFormat.Bold)
                                                                {
                                                                   contentBox.AppendText(Environment.NewLine);
                                                                }
                                                            }

                                                            contentBox.AppendText(text.Text);

                                                        }

                                                       // contentBox.AppendText(cobj.GetType() + Environment.NewLine);

                                                    }

                                                   // contentBox.AppendText(Environment.NewLine);

                                                }
                                            }
                                        }
                                    }

                                    if (nobj is Paragraph)
                                    {
                                        Paragraph para = nobj as Paragraph;

                                        //contentBox.AppendText(para.ListText +  para.Text + Environment.NewLine);
                                        if (para.ListText.Length > 0)
                                        {
                                           contentBox.AppendText(Environment.NewLine + para.ListText);
                                        }
                                        // + para.Text + Environment.NewLine);

                                        for (int i = 0; i < para.ChildObjects.Count; i++)
                                        {
                                            DocumentObject cobj = para.ChildObjects[i];

                                            if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                                            {
                                                TextRange text = cobj as TextRange;

                                                if (text.Text.Trim().Length > 0)
                                                {
                                                    if (text.CharacterFormat.Bold)
                                                    {
                                                        contentBox.AppendText(Environment.NewLine);
                                                    }
                                                }

                                                contentBox.AppendText(text.Text);

                                            }
                                           // contentBox.AppendText(cobj.GetType() + Environment.NewLine);

                                        }
                                       // contentBox.AppendText(Environment.NewLine);

                                    }

                                }

                            }
                        }
                    }
                    if (obj is Paragraph)
                    {
                        Paragraph para = obj as Paragraph;

                        if(para.ListText.Length>0)
                        {
                           contentBox.AppendText(Environment.NewLine + para.ListText);
                        }
                        // + para.Text + Environment.NewLine);

                        for(int i=0; i< para.ChildObjects.Count; i++)
                        {
                            DocumentObject cobj = para.ChildObjects[i];

                            if (cobj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange text = cobj as TextRange;

                                if(text.Text.Trim().Length>0)
                                {
                                    if (text.CharacterFormat.Bold)
                                    {
                                       contentBox.AppendText(Environment.NewLine);
                                    }
                                }

                                contentBox.AppendText(text.Text);
                                
                            }

                        }

                        //contentBox.AppendText(Environment.NewLine);

                    }
                }
            }
            //doc.SaveToFile(fpath + @"\" + Path.GetFileName(filename), FileFormat.Docx);

            RefineAnswers();
        }

        public void RefineAnswers()
        {
            questionsBox.Clear();

            int cunit = 1;
            string currentLine;
            string currentQuestion;
            int currentQN=1;
            int qn=1;

            for (int i = 0; i < contentBox.Lines.Count() - 1; i++)
            {
                currentLine = contentBox.Lines[i].ToString();

                currentQuestion = currentLine.Split('.')[0];
                if(currentQuestion.Trim().All(char.IsDigit))
                {
                    try
                    {
                        currentQN = Convert.ToInt32(currentQuestion);

                        if(currentQN >400)
                        {
                            continue;
                        }
                        
                        if(currentQN>=qn)
                        {
                            qn = currentQN;
                        }
                        else
                        {
                            qn = currentQN;
                            cunit += 1;
                        }

                    }
                    catch (Exception)
                    {

                        continue;
                    }
                    
                    questionsBox.AppendText(Environment.NewLine + "UNIT"+ "-" +cunit.ToString()+ "-"+ currentQN +"#" );
                    questionsBox.AppendText(currentLine.Replace(currentQN.ToString() + ".", ""));
                    string key = "UNIT" + "-" + cunit.ToString() + "-" + currentQN.ToString();
                    string val = currentLine.Replace(currentQN.ToString() + ".", "");
                    answers.Add(key, val);
                }
                else
                {
                    questionsBox.AppendText(currentLine);
                    //questionsBox.AppendText(currentLine.Replace(currentQN.ToString() + ".",""));
                }

                //currentQuestion = Convert.ToInt32(currentLine.Split('.')[0]);

                

            }
        }

        public static void FTPImageUpload()
        {
            string ftpUsername = "dev";
            string ftpPassword = "dev@123";
            //string localFile = "";
            string port = ":5456";
            string ipaddr = "203.92.41.138";

            //using (var client = new WebClient())
            //{
            //    client.Credentials = new NetworkCredential(ftpUsername, ftpPassword);
            //    client.UploadFile("ftp://host/path.zip", WebRequestMethods.Ftp.UploadFile, localFile);
            //}
            string filePath = @"D:\dd.txt";

            using (WebClient client = new WebClient())
            {
                client.Credentials = new NetworkCredential(ftpUsername, ftpPassword);

                Uri address = new Uri("ftp://dev@203.92.41.138:5456/upload/images");
               // byte[] rarResp = client.UploadFile(address, filePath);
                byte[] rawResponse = client.UploadFile(address, filePath);
                string response = System.Text.Encoding.ASCII.GetString(rawResponse);

                // check response
                MessageBox.Show(response);
            }


        }

        private void UploadFileToFTP()
        {
            FtpWebRequest ftpReq = (FtpWebRequest)WebRequest.Create("ftp://148.66.128.113/home/evelynlearning/public_html/guidelines.evelynlearning.com/images/chandan_bhushan.jpg");

            ftpReq.UseBinary = true;
            ftpReq.Method = WebRequestMethods.Ftp.UploadFile;
            ftpReq.Credentials = new NetworkCredential("evelynlearning", "U^s9y(4vB^[p");

            byte[] b = File.ReadAllBytes(@"D:\chandan_bhushan.jpg");
            ftpReq.ContentLength = b.Length;
            using (Stream s = ftpReq.GetRequestStream())
            {
                s.Write(b, 0, b.Length);
            }

            FtpWebResponse ftpResp = (FtpWebResponse)ftpReq.GetResponse();

            if (ftpResp != null)
            {
                if (ftpResp.StatusDescription.StartsWith("226"))
                {
                    Console.WriteLine("File Uploaded.");
                }
            }
        }

        // anchored end
    }
}
