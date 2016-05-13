using System;
using System.Collections.Generic;
using System.Data;//添加DataSet对象的命名空间
using System.Text;

namespace EvaluatingSystem
{
    using System;
    public class MyDLL
    {
        #region  设置DLL文件中静态变量
        private static DataSet DataCom = new DataSet();
        #endregion

        #region  设置DLL文件中静态方法
        /// <summary>
        /// 接收传递的DataSet对象
        /// </summary>
        /// <param Dset="DataSet">DataSet对象</param>
        public static void TakeOver(DataSet Dset)
        {
            DataCom = Dset;//获取DataSet对象
            Frm_PrintSet FrmPring = new Frm_PrintSet();//实例化打印窗体
            FrmPring.Show();//显示打印窗体
        }

        /// <summary>
        /// 发送传递的DataSet对象
        /// </summary>
        public static DataSet SendOut()
        {
            return DataCom;
        }
        #endregion
    }
}

