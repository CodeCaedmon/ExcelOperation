using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace Caedmon
{
    class ExcelOperation
    {
        public ExcelOperation()
        { }

        ///方法一：采用OleDB读取EXCEL文件： 
        public DataSet ExcelToDS(string Path)
        {
            // 连接字符串
            // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=2'";
            //Microsoft.ACE.OLEDB.12.0  不能用在X64 为平台上,需要修改配置管理X86
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + Path + ";" + "Extended Properties='Excel 12.0;HDR=Yes;IMEX=2'";
            //  provider：表示提供程序名称
            //  Data Source：这里填写Excel文件的路径
            //  Extended Properties：设置Excel的特殊属性
            //  Extended Properties 取值：
            //  Excel 8.0 针对Excel2000及以上版本，Excel5.0 针对Excel97。
            //  HDR = Yes 表示第一行包含列名,在计算行数时就不包含第一行
            //  IMEX 0:导入模式,1:导出模式: 2混合模式

            OleDbConnection conn = new OleDbConnection(strConn);
            //  在数据访问中首先必须建立到数据库的物理连接。OLEDB.NET Data Provider 使用OleDbConnection类的对象标识与一个数据库的物理连接。
            //  OleDbConnection类的常用属性及其说明:
            //  ———————————————————————————
            //  属性说明
            //  ConnectionString 获取或设置用于打开数据库的字符串
            //  ConnectionTimeOut 获取在尝试建立连接时终止尝试并生成错误之前所等待的时间
            //  Database 获取当前数据库或连接打开后要使用的数据库名称
            //  DataSource 获取数据源的服务器名或文件名
            //  Provider 获取在连接字符串的“Provider = ” 子句中指定的OLEDB提供程序的名称
            //  State 获取连接的当前状态
            //  ———————————————————————————————————————  
            //  Connecting 连接对象正在与数据源连接
            //  Executing  连接对象正在执行命令
            //  Fetching   连接对象正在检索数据
            //  Open       连接对象处于打开状态
            //————————————————————————————————————————
            //  OleDbConnection类的常用方法如下表所示：
            //————————————————————————————————————————
            //  Open  使用ConnectionString所指定的属性设置打开数据库连接
            //  Close 关闭与数据库的连接，这是关闭任何打开连接的首选方法
            //  CreateCommand  创建并返回一个与OleDbConnection关联的OleDbCommand对象
            //  ChangeDatabase 为打开的OleDbConnection更改当前数据库
            //————————————————————————————————————————

            conn.Open();
            //  OleDbDataAdapter 充当 DataSet 和数据源之间的桥梁，用于检索和保存数据。
            //  OleDbDataAdapter 通过以下方法提供这个桥接器：
            //  使用 Fill 将数据从数据源加载到 DataSet中，并使用 Update 将 DataSet 中所作的更改发回数据源。
            string strExcel = "select * from [sheet1$]";
            DataSet ds = new DataSet();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strExcel, strConn);            
            myCommand.Fill(ds, "table1");  //"table1 ?"
            return ds; 
        }

        public DataTable ExcelToDT(string Path)
        {
            // 连接字符串
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=2'";
            //  provider：表示提供程序名称
            //  Data Source：这里填写Excel文件的路径
            //  Extended Properties：设置Excel的特殊属性
            //  Extended Properties 取值：
            //  Excel 8.0 针对Excel2000及以上版本，Excel5.0 针对Excel97。
            //  HDR = Yes 表示第一行包含列名,在计算行数时就不包含第一行
            //  IMEX 0:导入模式,1:导出模式: 2混合模式

            OleDbConnection conn = new OleDbConnection(strConn);
            //  在数据访问中首先必须建立到数据库的物理连接。OLEDB.NET Data Provider 使用OleDbConnection类的对象标识与一个数据库的物理连接。
            //  OleDbConnection类的常用属性及其说明:
            //  ———————————————————————————
            //  属性说明
            //  ConnectionString 获取或设置用于打开数据库的字符串
            //  ConnectionTimeOut 获取在尝试建立连接时终止尝试并生成错误之前所等待的时间
            //  Database 获取当前数据库或连接打开后要使用的数据库名称
            //  DataSource 获取数据源的服务器名或文件名
            //  Provider 获取在连接字符串的“Provider = ” 子句中指定的OLEDB提供程序的名称
            //  State 获取连接的当前状态
            //  ———————————————————————————————————————  
            //  Connecting 连接对象正在与数据源连接
            //  Executing  连接对象正在执行命令
            //  Fetching   连接对象正在检索数据
            //  Open       连接对象处于打开状态
            //————————————————————————————————————————
            //  OleDbConnection类的常用方法如下表所示：
            //————————————————————————————————————————
            //  Open  使用ConnectionString所指定的属性设置打开数据库连接
            //  Close 关闭与数据库的连接，这是关闭任何打开连接的首选方法
            //  CreateCommand  创建并返回一个与OleDbConnection关联的OleDbCommand对象
            //  ChangeDatabase 为打开的OleDbConnection更改当前数据库
            //————————————————————————————————————————

            conn.Open();
            //  OleDbDataAdapter 充当 DataSet 和数据源之间的桥梁，用于检索和保存数据。
            //  OleDbDataAdapter 通过以下方法提供这个桥接器：
            //  使用 Fill 将数据从数据源加载到 DataSet中，并使用 Update 将 DataSet 中所作的更改发回数据源。
            string strExcel = "select * from [sheet1$]";   // sheet1 为源Excel中的默认表名
            DataTable dt = new DataTable();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strExcel, strConn);
            myCommand.Fill(dt, "table1");     //"table1 为dt的表名"
            return dt;
        }

        // IMEX表示是否强制转换为文本
        // 特别注意
        //Extended Properties = 'Excel 8.0;HDR=yes;IMEX=1'
        //A： HDR(HeaDer Row)设置
        //若指定值为Yes，代表 Excel 档中的工作表第一行是栏位名称
        //若指定值為 No，代表 Excel 档中的工作表第一行就是資料了，沒有栏位名称
        //B：IMEX(IMport EXport mode )设置
        //IMEX 有三种模式，各自引起的读写行为也不同，容後再述：
        //0 is Export mode
        //1 is Import mode
        //2 is Linked mode(full update capabilities)
        //我这里特别要说明的就是 IMEX 参数了，因为不同的模式代表著不同的读写行为：
        //当 IMEX = 0 时为“汇出模式”，这个模式开启的 Excel 档案只能用来做“写入”用途。
        //当 IMEX = 1 时为“汇入模式”，这个模式开启的 Excel 档案只能用来做“读取”用途。
        //当 IMEX = 2 时为“连結模式”，这个模式开启的 Excel 档案可同时支援“读取”与“写入”用途。

    }
}
