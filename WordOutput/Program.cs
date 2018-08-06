using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordOutput.DAL;

namespace WordOutput {
    class Program {
        static void Main(string[] args) {
            //程序流程开始
            MainContext db = new MainContext();
            //1. 获取运行位置路径
            //2. 获得待导入文件列表(文件名,文件状态:未导入,导入成功,导入失败)
            //3. 对列表中每个文件执行
            //   获取所需属性并赋值给DocModel对象
            //   存入数据库
            //4. 检查是否存入成功
            //5. 处理未成功文件
            //程序流程结束
        }
    }
}
