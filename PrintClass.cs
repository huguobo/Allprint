using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;
using System.Windows.Forms;

namespace EvaluatingSystem
{
    public class PrintClass
    {
        #region  全局变量
        private DataGridView datagrid;
        private PrintDocument printdocument;
        private PageSetupDialog pagesetupdialog;
        private PrintPreviewDialog printpreviewdialog;
        private string title = "";//标题

        int currentpageindex = 0;//当前页的编号
        int rowcount = 0;//数据的行数
        int pagecount = 0;//打印页数
        int titlesize = 16;//标题的大小
        int xoffset = 0;//标题位置
        public int x = 0;//z绘画时的x轴位置
        public int PrintPageHeight = 1169;//打印的默认高度
        public int PrintPageWidth = 827;//打印的默认宽度
        public bool iseverypageprinttitle = false;//是否每一页都要打印标题
        public int headerheight = 30;//标题高度
        public int topmargin = 60; //顶边距 
        public int celltopmargin = 6;//单元格顶边距 
        public int cellleftmargin = 4;//单元格左边距 
        public int pagerowcount = 7;//每页行数 
        public int rowgap = 23;//行高 
        public int colgap = 5;//每列间隔 
        public int leftmargin = 50;//左边距 
        public Font titlefont = new Font("arial", 14);//标题字体 
        public Font font = new Font("arial", 10);//正文字体

        public Font footerfont = new Font("arial", 8);//页脚显示页数的字体
        public Font uplinefont = new Font("arial", 9, FontStyle.Bold);//当header分两行显示的时候，上行显示的字体
        public Font underlinefont = new Font("arial", 8);//当header分两行显示的时候，下行显示的字体
        public Brush brush = new SolidBrush(Color.Black);//画刷
        public Font headerfont = new Font("arial", 9, FontStyle.Bold);//列名标题字体
        public Brush brushHeaderFont = new SolidBrush(Color.Black);//列名字体画刷
        public Brush brushHeaderBack = new SolidBrush(Color.White);//列名背景画刷
        public Font Cellfont = new Font("arial", 9);//单元格字体
        public Brush brushCellFont = new SolidBrush(Color.Black);//单元格字体画刷
        public Brush brushCellBack = new SolidBrush(Color.White);//单元格背景画刷
        public bool isautopagerowcount = true;//是否自动计算行数
        public int buttommargin = 80;//底边距 
        public bool needprintpageindex = true;//是否打印页脚页数 
        public Color LineColor = Color.Black;//线的颜色
        public bool LineUP = true;//上边线
        public bool LineLeft = true;//左边线
        public bool LineUnit = true;//单元格的线
        public bool Boundary = false;//下边线

        public int AlignmentSgin = -1;//对齐方式的标识
        public int HAlignment = 0;//标题对齐方式
        public int UAlignment = 0;//单元格对齐方式
        public int LeftAlignment = 50;//左对齐的边界
        public int RightAlignment = 50;//右对齐的边界
        public int CenterAlignment = 0;//居中对齐边界
        public bool PageAspect = false;//打印的方向
        public static bool PageScape = false;//打印方向
        public int PageSheet = 0;
        #endregion

        #region  打印信息的初始化
        /// <summary>
        /// 打印信息的初始化
        /// </summary>
        /// <param datagrid="DataGridView">打印数据</param>
        /// <param title="string">打印标题</param>
        /// <param titlesize="int">标题大小</param>
        /// <param PageS="int">纸张大小</param>
        /// <param lendscape="bool">是否横向打印</param>
        /// <returns>返回DataSet对象</returns>
        public PrintClass(DataGridView datagrid, string title, int titlesize, int PageS, bool lendscape)
        {
            this.title = title;//设置标题的名称
            this.titlesize = titlesize;//设置标题的大小
            this.datagrid = datagrid;//获取打印数所据
            this.PageSheet = PageS;//纸张大小
            printdocument = new PrintDocument();//实例化PrintDocument类
            pagesetupdialog = new PageSetupDialog();//实例化PageSetupDialog类
            pagesetupdialog.Document = printdocument;//获取当前页的设置
            printpreviewdialog = new PrintPreviewDialog();//实例化PrintPreviewDialog类
            printpreviewdialog.Document = printdocument;//获取预览文档的信息
            printpreviewdialog.FormBorderStyle = FormBorderStyle.Fixed3D;//设置窗体的边框样式

            //横向打印的设置
            if (PageSheet >= 0)
            {
                if (lendscape == true)
                {
                    printdocument.DefaultPageSettings.Landscape = lendscape;//横向打印
                }
                else
                {
                    printdocument.DefaultPageSettings.Landscape = lendscape;//纵向打印
                }
            }
            pagesetupdialog.Document = printdocument;
            //MessageBox.Show(printdocument.DefaultPageSettings.Landscape.ToString());
            printdocument.PrintPage += new PrintPageEventHandler(this.printdocument_printpage);//事件的重载
        }
        #endregion

        #region  纸张大小的设置
        /// <summary>
        ///  纸张大小的设置
        /// </summary>
        /// <param n="int">纸张大小的编号</param>
        /// <returns>返回string对象</returns>
        public string Page_Size(int n)
        {
            string pageN = "";//纸张的名称
            switch (n)
            {
                case 1: { pageN = "A5"; PrintPageWidth = 583; PrintPageHeight = 827; break; }
                case 2: { pageN = "A6"; PrintPageWidth = 413; PrintPageHeight = 583; break; }
                case 3: { pageN = "B5(ISO)"; PrintPageWidth = 693; PrintPageHeight = 984; break; }
                case 4: { pageN = "B5(JIS)"; PrintPageWidth = 717; PrintPageHeight = 1012; break; }
                case 5: { pageN = "Double Post Card"; PrintPageWidth = 583; PrintPageHeight = 787; break; }
                case 6: { pageN = "Envelope #10"; PrintPageWidth = 412; PrintPageHeight = 950; break; }
                case 7: { pageN = "Envelope B5"; PrintPageWidth = 693; PrintPageHeight = 984; break; }
                case 8: { pageN = "Envelope C5"; PrintPageWidth = 638; PrintPageHeight = 902; break; }
                case 9: { pageN = "Envelope DL"; PrintPageWidth = 433; PrintPageHeight = 866; break; }
                case 10: { pageN = "Envelope Monarch"; PrintPageWidth = 387; PrintPageHeight = 750; break; }
                case 11: { pageN = "ExeCutive"; PrintPageWidth = 725; PrintPageHeight = 1015; break; }
                case 12: { pageN = "Legal"; PrintPageWidth = 850; PrintPageHeight = 1400; break; }
                case 13: { pageN = "Letter"; PrintPageWidth = 850; PrintPageHeight = 1100; break; }
                case 14: { pageN = "Post Card"; PrintPageWidth = 394; PrintPageHeight = 583; break; }
                case 15: { pageN = "16K"; PrintPageWidth = 775; PrintPageHeight = 1075; break; }
                case 16: { pageN = "8.5x13"; PrintPageWidth = 850; PrintPageHeight = 1300; break; }
            }
            return pageN;//返回纸张的名
        }
        #endregion

        #region  页边距的设置
        /// <summary>
        ///  页边距的设置
        /// </summary>
        /// <param SetUp1="string[]">边距信息</param>
        public void PrintSetUp(string[] SetUp1)
        {
            if (SetUp1[0] == "true")
            {
                topmargin = Int32.Parse(SetUp1[1]);//顶边距
                leftmargin = Int32.Parse(SetUp1[2]);//左边距
                buttommargin = Int32.Parse(SetUp1[3]);//底边距
                AlignmentSgin = -1;//设置对齐方式的标识
            }
        }
        #endregion

        #region  文字的位置
        /// <summary>
        ///  文字的位置
        /// </summary>
        /// <param CellW="int">单元格的宽度</param>
        /// <param StrW="int">文字的宽度</param>
        /// <param colW="int">单元格的左边距</param>
        /// <param Ali="int">对齐方式</param>
        /// <returns>返回int对象</returns>
        private int Alignment_Mode(int CellW, int StrW, int colW, int Ali)
        {
            int ALiW = 0;
            switch (Ali)
            {
                case 0://左对齐
                    {
                        ALiW = colW; //设置文字的左端位置
                        break;
                    }
                case 1://局中
                    {
                        ALiW = (int)((CellW - StrW) / 2);//设置文字的左端位置
                        break;
                    }
                case 2://右对齐
                    {
                        ALiW = CellW - StrW - colW;//设置文字的左端位置
                        break;
                    }
            }
            return ALiW;
        }
        #endregion

        #region  页的打印事件
        /// <summary>
        ///  页的打印事件(主要用于绘制打印报表)
        /// </summary>
        private void printdocument_printpage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            PrintPageWidth = e.PageBounds.Width;//获取打印线张的宽度
            PrintPageHeight = e.PageBounds.Height;//获取打印线张的高度
            if (this.isautopagerowcount)//自动计算页的行数
                pagerowcount = (int)((PrintPageHeight - this.topmargin - titlesize - this.headerfont.Height - this.headerheight - this.buttommargin) / this.rowgap);//获取每页的行数
            pagecount = (int)(rowcount / pagerowcount);//获取打印多少页
            pagesetupdialog.AllowOrientation = true;//启动打印页面对话框的方向部分
            if (rowcount % pagerowcount > 0)//如果数据的行数大于页的行数
                pagecount++;//页数加1
            int colcount = 0;//记录数据的列数
            int y = topmargin;//获取表格的顶边距
            string cellvalue = "";//记录文本信息（列标题和单元格的文本信息）
            int startrow = currentpageindex * pagerowcount;//设置打印的初始页数
            int endrow = startrow + this.pagerowcount < rowcount ? startrow + pagerowcount : rowcount;//设置打印的最大页数
            int currentpagerowcount = endrow - startrow;//获取打印页数
            colcount = datagrid.ColumnCount;//获取打印数据的列数
            x = leftmargin;//获取表格的左边距

            //设置绘置背景颜色的点
            Point headup, headdown;
            //获取报表的宽度
            int cwidth = 0;
            for (int j = 0; j < colcount; j++)//循环数据的列数
            {
                if (datagrid.Columns[j].Width > 0)//如果列的宽大于０
                {
                    cwidth += datagrid.Columns[j].Width + colgap;//累加每列的宽度
                }
            }

            if (AlignmentSgin >= -1)//设置对齐方式的边界位置
            {
                int tn = (int)(e.Graphics.MeasureString(this.title, this.titlefont).Width);//获取标题的宽度
                switch (AlignmentSgin)//对齐方式
                {
                    case 0://左对齐
                        {
                            x = LeftAlignment;//获取左对齐的默认边界
                            leftmargin = x;//设置左边距
                            if (tn > cwidth)//如果标题的宽度大于表格的宽度
                                xoffset = leftmargin;//标题的左边距为表格的左边距
                            else
                                xoffset = (int)(PrintPageWidth - (PrintPageWidth - 50) + (cwidth - tn) / 2);//使标题局中
                            break;
                        }
                    case 1://局中
                        {
                            x = (PrintPageWidth - cwidth) / 2;//设置表格的局中位置
                            leftmargin = x;//设置左边距
                            xoffset = (int)((PrintPageWidth - tn) / 2);//标题相对于表格局中
                            break;
                        }
                    case 2://右对齐
                        {
                            x = PrintPageWidth - cwidth - RightAlignment;//设置表格右对齐的左边距位置
                            leftmargin = x;//设置左边距
                            if (tn > cwidth)//如果标题的宽度大于表格的宽度
                                xoffset = (int)(PrintPageWidth - tn);//使标题的右边距与表格的右边距相同
                            else
                                xoffset = (int)(PrintPageWidth - 50 - cwidth + (cwidth - tn) / 2);//标题相对于表格局中
                            break;
                        }
                    case -1://标题的默认状态
                        {
                            if (tn > cwidth)//如果标题的宽度大于表格的宽度
                            {
                                if ((tn - cwidth) / 2 < leftmargin)//标题在表格上局中，左边的超出的部分小于左边距
                                    xoffset = (int)(leftmargin - (tn - cwidth) / 2);//使标题相对于表格局中
                                else
                                    xoffset = leftmargin;//使标题的左边距与表格的左边距相同
                            }
                            else
                                xoffset = (int)(x + (cwidth - tn) / 2);//标题相对于表格局中
                            break;
                        }
                }
            }

            //绘置标题
            if (this.currentpageindex == 0 || this.iseverypageprinttitle)//当前页为首页，并且每页都要打印标题
            {
                e.Graphics.DrawString(this.title, titlefont, brush, xoffset, y);//绘制标题
                y += titlesize;//获取标题底端位置
            }
            y += rowgap;//设置表格的上边线的位置

            //绘制标题栏的背景颜色
            headup = new Point(x, y);//设置左上角位置
            headdown = new Point(cwidth, headerheight);//设置右下角位置
            drawrectangle(brushHeaderBack, headup, headdown, e.Graphics);//填充矩形框
            //绘制单元格的背景颜色
            headup = new Point(x, y + headerheight);//设置左上角位置
            headdown = new Point(cwidth, (endrow - startrow) * rowgap);//设置右下角位置
            drawrectangle(brushCellBack, headup, headdown, e.Graphics);//填充矩形框
            //画出打印表格最左边的竖线
            if (LineLeft == true)//如果是左边线
                drawline(new Point(x, y), new Point(x, y + currentpagerowcount * rowgap + this.headerheight), e.Graphics, 0);//画线
            //设置标题栏中的文字及坚线
            for (int j = 0; j < colcount; j++)//遍历列数据
            {
                int colwidth = datagrid.Columns[j].Width;//获取列的宽度
                if (colwidth > 0)//如果列的宽度大于0
                {
                    cellvalue = datagrid.Columns[j].HeaderText;//获取列标题
                    //绘制标题栏文字
                    int Ha = Alignment_Mode(datagrid.Columns[j].Width, (int)(e.Graphics.MeasureString(cellvalue, Cellfont).Width), cellleftmargin, HAlignment);//设置列标题的位置
                    e.Graphics.DrawString(cellvalue, headerfont, brushHeaderFont, x + Ha, y + celltopmargin);//绘制列标题
                    x += colwidth + colgap;//横向，下一个单元格的位置
                    //右侧坚线
                    if (LineUnit == true)//如果是左边单元格的线
                        drawline(new Point(x, y), new Point(x, y + currentpagerowcount * rowgap + this.headerheight), e.Graphics, 0);//画线
                    if ((LineLeft == true) && (j == (colcount - 1)))//如果是最右边的线
                        drawline(new Point(x, y), new Point(x, y + currentpagerowcount * rowgap + this.headerheight), e.Graphics, 0);//画线
                    int nnp = y + currentpagerowcount * rowgap + this.headerheight;//下一行线的位置
                }
            }
            int rightbound = x;
            //列标题上边的线
            if (LineUP == true)
                drawline(new Point(leftmargin, y), new Point(rightbound, y), e.Graphics, 0); //绘制最上面的线 
            headup = new Point(leftmargin, y);
            y += this.headerheight;//设置下一个线的位置
            //打印所有的行信息
            for (int i = startrow; i < endrow; i++) //对行进行循环
            {
                x = leftmargin;//获取线的Ｘ坐标点
                for (int j = 0; j < colcount; j++)//对列进行循环
                {
                    if (datagrid.Columns[j].Width > 0)//如果列的宽度大于0
                    {
                        cellvalue = datagrid.Rows[i].Cells[j].Value.ToString();//获取单元格的值
                        //绘制单元格中的信息
                        int Ua = Alignment_Mode(datagrid.Columns[j].Width, (int)(e.Graphics.MeasureString(cellvalue, Cellfont).Width), cellleftmargin, UAlignment);
                        e.Graphics.DrawString(cellvalue, Cellfont, brushHeaderFont, x + Ua, y + celltopmargin);//绘制单元格信息
                        x += datagrid.Columns[j].Width + colgap;//单元格信息的X坐标
                        y = y + rowgap * (cellvalue.Split(new char[] { '\r', '\n' }).Length - 1);//单元格信息的Y坐标
                    }
                }
                //单元格上边的线
                if (LineUnit == true)
                    drawline(new Point(leftmargin, y), new Point(rightbound, y), e.Graphics, 0);
                if (Boundary == true && i == startrow)//绘制分割线
                    drawline(new Point(leftmargin, y), new Point(rightbound, y), e.Graphics, 1);
                y += rowgap;//设置下行的位置
            }
            //表格最下面的边线
            if (LineUP == true)
                drawline(new Point(leftmargin, y), new Point(rightbound, y), e.Graphics, 0);//绘制最下面的线
            currentpageindex++;//下一页的页码
            if (this.needprintpageindex)//如果显示页脚
                e.Graphics.DrawString("共 " + pagecount.ToString() + " 页   第 " + this.currentpageindex.ToString() + " 页", this.footerfont, brush, PrintPageWidth - 200, (int)(PrintPageHeight - this.buttommargin / 2 - this.footerfont.Height));//绘制页脚信息
            if (currentpageindex < pagecount)//如果当前页不是最后一页
            {
                e.HasMorePages = true;//打印副页
            }
            else
            {
                e.HasMorePages = false;//不打印副页
                this.currentpageindex = 0;//当前打印的页编号设为0
            }
        }
        #endregion

        #region  绘制边线
        /// <summary>
        ///  绘制边线
        /// </summary>
        /// <param sp="Point">左上角的坐标</param>
        /// <param ep="Point">右下角的坐标</param>
        /// <param gp="Graphics">Graphics类</param>
        /// <param n="int">标识</param>
        private void drawline(Point sp, Point ep, Graphics gp, int n)
        {
            int w = 1;//设置线的宽度
            if (n == 1)//如果是分割线
                w = 2;//设置线宽为2
            Pen pen = new Pen(LineColor, w);//设置画笔样式
            gp.DrawLine(pen, sp, ep);//绘制线
        }
        #endregion

        #region  绘制填充的矩形框
        /// <summary>
        ///  绘制填充的矩形框
        /// </summary>
        /// <param ColorB="Brush">画刷颜色</param>
        /// <param P1="Point">左上角的坐标</param>
        /// <param P2="Point">右下角的坐标</param>
        /// <param gp="Graphics">Graphics类</param>
        private void drawrectangle(Brush ColorB, Point P1, Point P2, Graphics gp)
        {
            gp.FillRectangle(ColorB, P1.X, P1.Y, P2.X, P2.Y);//填充一个矩形框
        }
        #endregion

        #region 显示打印预览窗体
        /// <summary>
        ///  显示打印预览窗体
        /// </summary>
        public void print()
        {

            rowcount = 0;//记录数据的行数
            string paperName = Page_Size(PageSheet);//获取当前纸张的大小
            PageSettings storePageSetting = new PageSettings();//实列化一个对PageSettings对象
            foreach (PaperSize ps in printdocument.PrinterSettings.PaperSizes)//查找当前设置纸张
                if (paperName == ps.PaperName)//如果找到当前纸张的名称
                {
                    storePageSetting.PaperSize = ps;//获取当前纸张的信息
                }
            //--printdocument.DefaultPageSettings.Landscape = PageScape;//设置横向打印



            //if (PageScape)//如果是横向打印
            //{
            //    storePageSetting.PaperSize = new PaperSize("Custom", PrintPageHeight, PrintPageWidth);//设置打印页的大小
            //}
            //else
            //{
            //    storePageSetting.PaperSize = new PaperSize("Custom", PrintPageWidth, PrintPageHeight);//设置打印页的大小
            //}




            //--pagesetupdialog.Document = printdocument;
            //printdocument.DefaultPageSettings = storePageSetting;//对打印机进行设置
            if (datagrid.DataSource.GetType().ToString() == "System.Data.DataTable")//判断数据类型
            {
                rowcount = ((DataTable)datagrid.DataSource).Rows.Count;//获取数据的行数
            }
            else if (datagrid.DataSource.GetType().ToString() == "System.Collections.ArrayList")//判断数据类型
            {
                rowcount = ((ArrayList)datagrid.DataSource).Count;//获取数据的行数
            }
            try
            {
                printdocument.DefaultPageSettings.Landscape = PageScape;//设置横向打印
                pagesetupdialog.Document = printdocument;

                printpreviewdialog.ShowDialog();//显示打印预览窗体
            }
            catch (Exception e)
            {
                throw new Exception("printer error." + e.Message);
            }
        }
        #endregion
    }
}

