using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Aspose.Words;
using System.Drawing;
using Aspose.Words.Tables;
using System.Data;
using Aspose.Words.Drawing;

namespace AsposeTest
{
    [TestClass]
    public class WordsTest
    {
        private string dir = @"F:/AsposeDemoDoc/";
        [TestMethod]
        public void IsEmpty()
        {
            Document doc = new Document();
            //验证创建的新对象是否包含子节点，Document对象至少包含一个子节点
            Assert.AreEqual(true, doc.HasChildNodes);
            doc.Save(dir + "1、Empty.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void Write_HelloWord()
        {
            Document doc = new Document();
            DocumentBuilder docBuilder = new DocumentBuilder(doc);
            docBuilder.Write("Hello Word");
            doc.Save(dir + "2、Write_HelloWord.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void WriteLn_HelloWord()
        {
            Document doc = new Document();
            DocumentBuilder docBuilder = new DocumentBuilder(doc);
            docBuilder.Writeln("Hello Word");
            doc.Save(dir + "3、WriteLn_HelloWord.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocStyle_Font()
        {
            Document doc = new Document();
            DocumentBuilder docBuilder = new DocumentBuilder(doc);
            //粗体
            docBuilder.Font.Bold = true;
            //文字颜色
            docBuilder.Font.Color = System.Drawing.Color.Red;
            //文字大小
            docBuilder.Font.Size = 20f;
            //下划线
            docBuilder.Font.Underline = Underline.Single;
            //字体
            docBuilder.Font.Name = "仿宋";

            //显示文字
            docBuilder.Writeln("Hello Word");
            docBuilder.Writeln("粗体、红色、20f、下划线、仿宋");

            //空行
            docBuilder.Writeln();
            //空行
            docBuilder.InsertParagraph();

            //修改文字样式
            //粗体
            docBuilder.Font.Bold = false;
            //文字颜色
            docBuilder.Font.Color = System.Drawing.Color.Black;
            //文字大小
            docBuilder.Font.Size = 14f;
            //下划线
            docBuilder.Font.Underline = Underline.None;
            //字体
            docBuilder.Font.Name = "华文彩云";

            docBuilder.Writeln("不是粗体、黑色、14f、无下划线、华文彩云");

            doc.Save(dir + "4、DocStyle_Font.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocStyle_Paragrapht()
        {
            Document doc = new Document();
            DocumentBuilder docBuilder = new DocumentBuilder(doc);
            ////行间距1.5倍
            docBuilder.ParagraphFormat.LineSpacing = 18f;
            ////左对齐
            docBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            ////首行缩进值
            docBuilder.ParagraphFormat.FirstLineIndent = 22f;

            docBuilder.Writeln("Aspose于2002年3月在澳大利亚悉尼创建。公司网站于2002年10月对外发布。Aspose 一直致力于成为全球最大的.Net 组件提供商，为全球.NET 程序员提供最丰富的选择。数十个国家的数千机构选择了Aspose的产品，这包括微软、IBM、普华永道、安永、杜邦、希尔顿酒店、读者文摘、美洲银行、波音、西门子等等。");
            docBuilder.Writeln("Aspose.Words是一款先进的类库，通过它可以直接在各个应用程序中执行各种文档处理任务。Aspose.Words支持DOC，OOXML，RTF，HTML，OpenDocument, PDF, XPS, EPUB和其他格式。使用Aspose.Words，您可以生成，更改，转换，渲染和打印文档而不使用Microsoft Word。");
            docBuilder.Writeln("Aspose.Cells是一个广受赞誉的电子表格组件，支持所有Excel格式类型的操作，用户无需依靠Microsoft Excel也可为其应用程序嵌入读写和处理Excel数据表格的功能。Aspose.Cells可以导入和导出每一个具体的数据，表格和格式，在各个层面导入图像，应用复杂的计算公式，并将Excel的数据保存为各种格式等等---完成所有的这一切功能都无需使用Microsoft Excel 和Microsoft Office Automation。");

            //空行
            docBuilder.Writeln();
            //空行
            docBuilder.InsertParagraph();

            ////行间距
            docBuilder.ParagraphFormat.LineSpacing = 30f;
            ////左对齐
            docBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            ////首行缩进值
            docBuilder.ParagraphFormat.FirstLineIndent = 0f;

            docBuilder.Writeln("Aspose于2002年3月在澳大利亚悉尼创建。公司网站于2002年10月对外发布。Aspose 一直致力于成为全球最大的.Net 组件提供商，为全球.NET 程序员提供最丰富的选择。数十个国家的数千机构选择了Aspose的产品，这包括微软、IBM、普华永道、安永、杜邦、希尔顿酒店、读者文摘、美洲银行、波音、西门子等等。");
            docBuilder.Writeln("Aspose.Words是一款先进的类库，通过它可以直接在各个应用程序中执行各种文档处理任务。Aspose.Words支持DOC，OOXML，RTF，HTML，OpenDocument, PDF, XPS, EPUB和其他格式。使用Aspose.Words，您可以生成，更改，转换，渲染和打印文档而不使用Microsoft Word。");
            docBuilder.Writeln("Aspose.Cells是一个广受赞誉的电子表格组件，支持所有Excel格式类型的操作，用户无需依靠Microsoft Excel也可为其应用程序嵌入读写和处理Excel数据表格的功能。Aspose.Cells可以导入和导出每一个具体的数据，表格和格式，在各个层面导入图像，应用复杂的计算公式，并将Excel的数据保存为各种格式等等---完成所有的这一切功能都无需使用Microsoft Excel 和Microsoft Office Automation。");


            doc.Save(dir + "5、DocStyle_Paragrapht.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_Link()
        {
            Document doc = new Document();
            DocumentBuilder docBuilder = new DocumentBuilder(doc);

            docBuilder.Write("国内最流行的搜索引擎是：");
            //设置超链接样式
            docBuilder.Font.Color = Color.Blue;
            docBuilder.Font.Underline = Underline.Single;
            docBuilder.InsertHyperlink("百度", "http://www.baidu.com", false);

            //清除设置的样式，使用默认样式
            docBuilder.Font.ClearFormatting();

            doc.Save(dir + "6、DocContent_Link.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_Catalog()
        {
            // Use a blank document
            Document doc = new Document();

            // Create a document builder to insert content with into document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Start the actual document content on the second page.
            builder.InsertBreak(BreakType.PageBreak);

            // Build a document with complex structure by applying different heading styles thus creating TOC entries.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 2");
            builder.Writeln("Heading 3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 3.1.1");
            builder.Writeln("Heading 3.1.2");
            builder.Writeln("Heading 3.1.3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.2");
            builder.Writeln("Heading 3.3");

            builder.ParagraphFormat.ClearFormatting();

            // Call the method below to update the TOC.
            doc.UpdateFields();

            doc.Save(dir + "7、DocContent_Catalog.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_Image()
        {
            // This creates a builder and also an empty document inside the builder.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a raster image.
            System.Drawing.Image rasterImage = System.Drawing.Image.FromFile(@"F:\AsposeDemo\1.jpg");
            try
            {
                builder.Write("Raster image: ");
                var shape = builder.InsertImage(rasterImage);
                //修改图片的宽高
                //shape.Width = 400f;
            }
            finally
            {
                rasterImage.Dispose();
            }

            builder.Document.Save(dir + "8、DocContent_Image.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_Template()
        {
            string tmplateDir = @"F:\AsposeDemo\template.docx";
            //通过文件路径加载模板
            Document doc = new Document(tmplateDir);
            DocumentBuilder docBuilder = new DocumentBuilder(doc);

            string[] filedArray = { "Title", "Date", "Year", "Tag", "Content" };
            string[] mergeValue = { "标题", DateTime.Now.ToString("dd/MM"), DateTime.Now.Year.ToString(), "Template", "使用Aspose.Words读取文档模板、合并文档" };

            //为合并域填充值
            doc.MailMerge.Execute(filedArray, mergeValue);

            //跳转到书签，继续插入内容
            docBuilder.MoveToBookmark("Document_Content");
            docBuilder.Writeln("这个位置是书签Document_Content");
            docBuilder.Writeln("这个位置是书签Document_Content");
            docBuilder.Writeln("这个位置是书签Document_Content");
            docBuilder.Writeln("这个位置是书签Document_Content");
            docBuilder.Writeln("这个位置是书签Document_Content");

            doc.Save(dir + "9、DocContent_Template.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_Table()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Insert a cell
            builder.InsertCell();
            // 固定列宽
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("This is row 1 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Write("This is row 1 cell 2");

            builder.EndRow();

            // Insert a cell
            builder.InsertCell();

            // Apply new row formatting，设置行高度
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            //自下而上
            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Writeln("This is row 2 cell 1");

            // Insert a cell
            builder.InsertCell();
            //自上而下
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();

            builder.EndTable();

            doc.Save(dir + "10、DocContent_Table.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_TableCellMerge()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();

            ////4行
            for (int i = 1; i <= 4; i++)
            {
                //4列
                for (int j = 1; j <= 4; j++)
                {
                    builder.InsertCell();
                    builder.Write("第" + i + "行，第" + j + "列");
                }

                builder.EndRow();
            }

            builder.EndTable();

            builder.Writeln();
            builder.Writeln();

            builder.StartTable();

            ////4行
            for (int i = 1; i <= 4; i++)
            {
                //4列
                for (int j = 1; j <= 4; j++)
                {
                    builder.InsertCell();
                    //第一行横向合并
                    if (i == 1)
                    {
                        builder.CellFormat.VerticalMerge = CellMerge.None;
                        if (j == 1)
                        {
                            builder.CellFormat.HorizontalMerge = CellMerge.First;
                            builder.Write("第" + i + "行已横向合并");
                        }
                        else
                        {
                            //横向合并
                            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                        }
                    }
                    else
                    {
                        builder.CellFormat.HorizontalMerge = CellMerge.None;
                        //第2/3/4行的第一列垂直合并
                        if (j == 1)
                        {
                            if (i == 2)
                            {
                                builder.Write("第2行第1列");
                                builder.CellFormat.VerticalMerge = CellMerge.First;
                            }
                            else
                            {
                                //垂直合并
                                builder.CellFormat.VerticalMerge = CellMerge.Previous;
                            }
                        }
                        else
                        {
                            if ((i == 3 || i == 4) && (j == 2 || j == 3))
                            {
                                if (i == 3 && j == 2)
                                {
                                    builder.CellFormat.VerticalMerge = CellMerge.First;
                                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                                    builder.Write("第3行第2列");
                                }
                                if (i == 4 && j == 2)
                                {
                                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                                    builder.CellFormat.VerticalMerge = CellMerge.Previous;
                                }
                                if ((i == 3 || i == 4) && j == 3)
                                {
                                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                                }
                            }
                            else
                            {
                                builder.CellFormat.VerticalMerge = CellMerge.None;
                                builder.Write("第" + i + "行，第" + j + "列");
                            }
                        }
                    }
                }

                builder.EndRow();
            }

            builder.EndTable();


            doc.Save(dir + "11、DocContent_TableCellMerge.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_TableMailMerge()
        {
            string tmplateDir = @"F:\AsposeDemo\table.docx";
            Document doc = new Document(tmplateDir);
            DocumentBuilder docBuilder = new DocumentBuilder(doc);

            DataTable dt = new DataTable("Dt");
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");
            dt.Columns.Add("Sex");

            dt.Rows.Add(new object[] { 001, "张三", "男" });
            dt.Rows.Add(new object[] { 002, "李四", "男" });

            doc.MailMerge.ExecuteWithRegions(dt);


            doc.Save(dir + "12、DocContent_TableMailMerge.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_ManyTableMailMerge()
        {
            string tmplateDir = @"F:\AsposeDemo\table1.docx";
            Document doc = new Document(tmplateDir);
            DocumentBuilder docBuilder = new DocumentBuilder(doc);

            DataTable dt = new DataTable("Dt");
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name");
            dt.Columns.Add("Sex");

            dt.Rows.Add(new object[] { 001, "张三", "男" });
            dt.Rows.Add(new object[] { 002, "李四", "男" });

            DataTable xmDt = new DataTable("XM");
            xmDt.Columns.Add("UserId", typeof(int));
            xmDt.Columns.Add("X");
            xmDt.Columns.Add("M");

            xmDt.Rows.Add(new object[] { 001, "张", "三" });
            xmDt.Rows.Add(new object[] { 001, "张", "三" });
            xmDt.Rows.Add(new object[] { 001, "张", "三" });
            xmDt.Rows.Add(new object[] { 002, "李", "四" });
            xmDt.Rows.Add(new object[] { 002, "李", "四" });
            xmDt.Rows.Add(new object[] { 002, "李", "四" });

            DataSet ds = new DataSet();
            //将两个表插入到DataSet中
            ds.Tables.Add(dt);
            ds.Tables.Add(xmDt);
            //设置父子表的关系，主表Dt的Id关联从表XM的UserId
            ds.Relations.Add(dt.Columns["Id"], xmDt.Columns["UserId"]);

            doc.MailMerge.ExecuteWithRegions(ds);

            doc.Save(dir + "13、DocContent_ManyTableMailMerge.docx", SaveFormat.Docx);
        }

        [TestMethod]
        public void DocContent_InsertDocument()
        {
            string tmplateDir = @"F:\AsposeDemo\table1.docx";
            Document doc = new Document(tmplateDir);
            DocumentBuilder docBuilder = new DocumentBuilder(doc);

            DataTable dt = new DataTable("Dt");
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name");
            dt.Columns.Add("Sex");

            dt.Rows.Add(new object[] { 001, "张三", "男" });
            dt.Rows.Add(new object[] { 002, "李四", "男" });

            DataTable xmDt = new DataTable("XM");
            xmDt.Columns.Add("UserId", typeof(int));
            xmDt.Columns.Add("X");
            xmDt.Columns.Add("M");

            xmDt.Rows.Add(new object[] { 001, "张", "三" });
            xmDt.Rows.Add(new object[] { 001, "张", "三" });
            xmDt.Rows.Add(new object[] { 001, "张", "三" });
            xmDt.Rows.Add(new object[] { 002, "李", "四" });
            xmDt.Rows.Add(new object[] { 002, "李", "四" });
            xmDt.Rows.Add(new object[] { 002, "李", "四" });

            DataSet ds = new DataSet();
            //将两个表插入到DataSet中
            ds.Tables.Add(dt);
            ds.Tables.Add(xmDt);
            //设置父子表的关系，主表Dt的Id关联从表XM的UserId
            ds.Relations.Add(dt.Columns["Id"], xmDt.Columns["UserId"]);

            doc.MailMerge.ExecuteWithRegions(ds);


            //创建新文档
            Document newDoc = new Document(tmplateDir);
            DocumentBuilder newBuilder = new DocumentBuilder(newDoc);
            newBuilder.Writeln("插入表格，表格来源于从模板动态填充数据。");
            newBuilder.InsertParagraph();

            newDoc.AppendDocument(doc, ImportFormatMode.UseDestinationStyles);

            newDoc.Save(dir + "14、DocContent_InsertDocument.docx", SaveFormat.Docx);
        }
    }
}
