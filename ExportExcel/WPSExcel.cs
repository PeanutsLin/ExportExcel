using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;

namespace ExportExcel
{
    class WPSExcel
    {
        private Excel.Application application;
        private Excel.Workbooks workBooks;
        private Excel.Workbook workBook;
        private Excel.Worksheet workSheet;
        private Excel.Range range;

        private object oMissing = System.Reflection.Missing.Value;

        /// <summary>
        /// 构造
        /// </summary>
        public  WPSExcel()
        {
            application = new Excel.Application();
            workBooks = application.Workbooks;      
        }

        /// <summary>
        /// 创建表格
        /// </summary>
        /// <param name="_name"></param>
        public void CreateExcel(string _name)
        {
            System.IO.StreamReader sr;

            if (!System.IO.File.Exists(_name))
            {
                sr = new System.IO.StreamReader(System.IO.File.Create(_name));
            }      
            else
            {
                sr = new System.IO.StreamReader(_name);
            }
            using(sr)
            {
                sr.Dispose();
            }
        }

       /// <summary>
        /// 打开表格,并选择工作表
       /// </summary>
       /// <param name="_name"></param>
       /// <param name="_workSheet"></param>
       /// <returns></returns>
        public bool OpenExcel(string _name, int _workSheet)
        {
            if (!System.IO.File.Exists(_name))
            {
                throw new Exception("文件不存在");
            }

            workBook = workBooks.Open(_name, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing); 
                       
            if (workBook.Worksheets.Count < _workSheet)
            {
                throw new Exception("工作表不存在");
            }
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(_workSheet);
            range = (Excel.Range)workSheet.UsedRange;

            return true;
        }

        /// <summary>
        /// 关闭文件和相关资源
        /// </summary>
        /// <param name="_name"></param>
        public void CloseExcel()
        {
            if (range == null) return;

            workBooks.Close();
            application.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
        }

        /// <summary>
        /// 读取指定位置的内容
        /// </summary>
        /// <param name="_row"></param>
        /// <param name="_line"></param>
        public string Read(int _row, int _line)
        {
            if (range == null)
            {
                throw new Exception("打开表格失败");
            }
            string data = range.get_Item(_row, _line).Text;

            return data;         
        }

        /// <summary>
        /// 向指定位置写入
        /// </summary>
        /// <param name="_row"></param>
        /// <param name="_line"></param>
        public void Write(int _row, int _line, string _text)
        {
            ((Excel.Range)range.get_Item(_row, _line)).Value = _text;
        }
    }
}
