using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Aspose.Words;
using OfficeOpenXml;


namespace AutoDocument
{
    public partial class MainWindow : Window
    {
        private void AutoDocument_Click(object sender, RoutedEventArgs e)
        {
            string workSheetName = "Sheet1";
            string dataFile = @"Data\名单.xlsx";

            var nameList = new List<string>();    //学生姓名列表

            //1、从Excel中逐行读取原始数据（读到空行为止），并保存为List变量
            if (!File.Exists(dataFile))
            {
                Debug.Print("文件不存在");
            }
            var file = new FileInfo(dataFile);

            int currRow = 1;
            try
            {
                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets[workSheetName];
                    //当前行不为空则读入数据
                    while(!string.IsNullOrWhiteSpace(worksheet.Cells[currRow, 1].Value?.ToString() ?? string.Empty))
                    {
                        nameList.Add((worksheet.Cells[currRow, 1].Value?.ToString() ?? string.Empty).Trim());
                        currRow++;
                    }
                }

            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message.ToString());
            }

            //2、List变量写入Word文档
            string templateFile = @"Templates\光荣榜模板.docx";
            string outputFile = @"DocumentOut\光荣榜.docx";

            //从UI中提取数据
            int oneLineCounts = Convert.ToInt32(OneLineCounts.Text);int whileSpaceCounts = Convert.ToInt32(WhileSpaceCounts.Text);

            Document doc;
            doc = new Document(templateFile);
            var builder = new DocumentBuilder(doc);
            for(int i=0;i<nameList.Count;i++)
            {
                if (i % oneLineCounts == 0)
                {
                    builder.Writeln();
                }
                builder.Write(nameList[i]);
                for(int j=0;j<whileSpaceCounts;j++)
                {
                    builder.Write(" ");
                }    

            }
            doc.Save(outputFile, SaveFormat.Docx);
            MessageBox.Show("成功生成文档！");

        }

        private void OpenDocument_Click(object sender, RoutedEventArgs e)
        {
            string documentFile = @"DocumentOut\光荣榜.docx"; ;
            if (File.Exists(documentFile))
            {
                Process.Start(documentFile);
            }
            else
            {
                MessageBox.Show($"请先生成报告。");
            }
        }
    }
}
