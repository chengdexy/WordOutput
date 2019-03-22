using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using WordOutput.DAL;
using WordOutput.Models;
using Aspose.Words;
using Aspose.Words.Tables;

namespace WordOutput {
    class Program {
        static void Main(string[] args) {
            //程序流程开始
            MainContext db = new MainContext();
            //1. 获取运行位置路径
            Console.Write("请输入文件夹路径->");
            string appPath= Console.ReadLine();
            if (!Directory.Exists(appPath))
            {
                Console.WriteLine("路径不存在.");
                Console.ReadKey();
            }
            else
            {
                //2. 获得待导入文件列表(文件名,文件状态:未导入,导入成功,导入失败)
                List<string> fileList = new List<string>();
                DirectoryInfo folder = new DirectoryInfo(appPath);
                foreach (FileInfo file in folder.GetFiles("*.doc"))
                {
                    fileList.Add(file.FullName);
                }
                //3. 对列表中每个文件执行
                //   获取所需属性并赋值给DocModel对象
                var n = 1;
                fileList.ForEach(file => {
                    Console.WriteLine($"正在进行第{n}个, 共{fileList.Count()}个");
                    DocModel doc = GetDocModel(file);
                    db.DocModels.Add(doc);
                    db.SaveChanges();
                    n++;
                });
                //   存入数据库
                //4. 检查是否存入成功
                //5. 处理未成功文件
                //程序流程结束
                Console.WriteLine("全部完成, 按任意键退出....");
                Console.ReadKey();
            }
        }

        private static DocModel GetDocModel(string file) {
            Document doc = new Document(file);
            var table = doc.GetChildNodes(NodeType.Table, true)[0] as Table;
            string test = table.Rows[2].Cells[0].GetText().Replace("\a", "").Replace("\r", "").Trim();
            int n = 0;
            if (test == "家庭住址") {
                //多出来一行
                n = 1;
            }
            DocModel model = new DocModel {
                Name = table.Rows[0].Cells[1].GetText().Replace("\r", ",").Replace("\a", ";"),
                Gender = table.Rows[0].Cells[3].GetText().Replace("\r", ",").Replace("\a", ";"),
                Age = table.Rows[0].Cells[5].GetText().Replace("\r", ",").Replace("\a", ";"),
                SickYears = table.Rows[0].Cells[7].GetText().Replace("\r", ",").Replace("\a", ";"),
                PhoneNumbers = table.Rows[0].Cells[9].GetText().Replace("\r", ",").Replace("\a", ";"),
                Address = table.Rows[n + 1].Cells[1].GetText().Replace("\r", ",").Replace("\a", ";"),
                VipNumber = table.Rows[n + 1].Cells[3].GetText().Replace("\r", ",").Replace("\a", ";"),
                KindOfSick = table.Rows[n + 2].Cells[1].GetText().Replace("\r", ",").Replace("\a", ";"),
                UsingDrugs = table.Rows[n + 3].Cells[1].GetText().Replace("\r", ",").Replace("\a", ";"),
                CurSymptoms = table.Rows[n + 4].Cells[1].GetText().Replace("\r", ",").Replace("\a", ";"),
                Records = table.Rows[n + 5].Cells[1].GetText().Replace("\r", ",").Replace("\a", ";") + "\n$$\n" +
                table.Rows[n + 5].Cells[2].GetText().Replace("\r", ",").Replace("\a", ";")
            };
            return model;
        }
    }
}
