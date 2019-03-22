using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using WordOutput.DAL;
using WordOutput.Models;
using Aspose.Words;
using Aspose.Words.Tables;

namespace WordOutput
{
    class Program
    {
        static void Main(string[] args)
        {
            //程序流程开始
            MainContext db = new MainContext();
            //0. 检测数据库是否已经包含数据
            int count = db.DocModels.Count();
            if (count != 0)
            {
                Console.WriteLine($"数据库中已经包含导出的数据({count}条),请直接使用,或清空后再运行此程序.");
                Console.WriteLine(@"立刻清空数据库?(y/n)");
                string choose = Console.ReadLine();
                if (choose == "y")
                {
                    //清空数据库
                    db.DocModels.RemoveRange(db.DocModels.ToList());
                    db.SaveChanges();
                }
            }
            else
            {
                //1. 获取运行位置路径
                string appPath;
                if (args.Length > 0)
                {
                    appPath = args[0];
                }
                else
                {
                    appPath = @"D:\Shared\测试数据";
                }
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
                fileList.ForEach(file =>
                {
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

        private static DocModel GetDocModel(string file)
        {
            Document doc = new Document(file);
            var table = doc.GetChildNodes(NodeType.Table, true)[0] as Table;
            string test = table.Rows[2].Cells[0].GetText().Replace("\a", "").Replace("\r", "").Trim();
            int n = 0;
            if (test == "家庭住址")
            {
                //多出来一行
                n = 1;
            }
            //TODO：这里需要进一步优化，应直接输出最终结果，不应是混乱的文本。
            var name = table.Rows[0].Cells[1].GetText().Replace("\a", "").Replace("\r", "");
            var genderForTest = table.Rows[0].Cells[3].GetText().Replace("\a", "").Replace("\r", "");
            var gender = string.IsNullOrWhiteSpace(genderForTest) ? "男" : genderForTest;
            var ageForTest = table.Rows[0].Cells[5].GetText().Replace("\a", "").Replace("\r", "");
            var age = string.IsNullOrWhiteSpace(ageForTest) ? "0" : ageForTest;
            var sickYears = table.Rows[0].Cells[7].GetText().Replace("\a", "").Replace("\r", "");
            var phoneForTest = table.Rows[0].Cells[9].GetText().Replace("\a", "").Replace("\r", " ");
            var phoneNumbers = string.IsNullOrWhiteSpace(phoneForTest) ? "(未留电话)" : phoneForTest;
            var addressForTest = table.Rows[n + 1].Cells[1].GetText().Replace("\a", "").Replace("\r", "");
            var address = string.IsNullOrWhiteSpace(addressForTest) ? "(未留地址)" : addressForTest;
            var vipForTest = table.Rows[n + 1].Cells[3].GetText().Replace("\a", "").Replace("\r", "");
            var vipNumber = string.IsNullOrWhiteSpace(vipForTest) ? "000000" : string.Format("{0:D6}", Convert.ToInt32(vipForTest));
            var kindOfSick = table.Rows[n + 2].Cells[1].GetText().Replace("\a", "").Replace("\r", "");
            var usingDrugs = table.Rows[n + 3].Cells[1].GetText().Replace("\a", "").Replace("\r", "");
            var curSymptoms = table.Rows[n + 4].Cells[1].GetText().Replace("\a", "").Replace("\r", "");
            var records = table.Rows[n + 5].Cells[1].GetText().Replace("\a", "").Replace("\r", "\n").Replace(".", "-") +
                table.Rows[n + 5].Cells[2].GetText().Replace("\a", "").Replace("\r", "\n").Replace("．", "-").Replace(".", "-");
            DocModel model = new DocModel
            {
                Name = name,
                Gender = gender,
                Age = age,
                SickYears = sickYears,
                PhoneNumbers = phoneNumbers,
                Address = address,
                VipNumber = vipNumber,
                KindOfSick = kindOfSick,
                UsingDrugs = usingDrugs,
                CurSymptoms = curSymptoms,
                Records = records
            };
            return model;
        }
    }
}
