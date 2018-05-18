using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace TestProjectWeb.Utility
{
    /// <summary>
    /// 执行Excel VBA宏
    /// 注意：在IIS正确运行，需要进行DCOM的权限设置，具体步骤：
    ///  dcomcnfg.exe -> 组件服务 -> 计算机 -> 我的电脑 -> DCOM配置 ->Microsoft Excel 97-2003 Worksheet - >属性 -> 安全
    ///     -> 三个配置全部”自定义“，增加对于相应用户的权限。根据IIS应用程序池标识
    ///     如果是ApplicationPoolIdentity，查找用户 IIS AppPool\[应用程序池名称]
    ///     如果是NetworkService/LocalService，查找对应的用户组
    ///     如果是Local System，不需要以上配置
    /// </summary>
    public class ExcelMacroUtil
    {
        /// <summary>
        /// 执行Excel中的宏
        /// </summary>
        /// <param name="excelFilePath">Excel文件路径</param>
        /// <param name="macroName">Excel宏名称</param>
        /// <param name="parameters">Excel宏参数组</param>
        /// <param name="rtnValue">Excel宏返回值</param>
        /// <param name="isShowExcel">执行时是否打开并显示Excel</param>
        public void RunExcelMacro(string excelFilePath, string macroName, object[] parameters, out object rtnValue, bool isShowExcel)
        {
            #region 初始化对象
            Excel.ApplicationClass oExcel = null;
            Excel.Workbooks oBooks = null;
            Excel._Workbook oBook = null;
            #endregion

            try
            {
                #region 检查参数

                // 检查文件是否存在
                if (!File.Exists(excelFilePath))
                {
                    throw new System.Exception(excelFilePath + " 文件不存在");
                }

                // 检查是否输入宏名称
                if (string.IsNullOrEmpty(macroName))
                {
                    throw new System.Exception("请输入宏的名称");
                }

                #endregion

                #region 调用宏处理

                // 缺省参数对象
                object oMissing = System.Reflection.Missing.Value;

                // 根据参数组是否为空，准备参数组对象
                object[] paraObjects;

                if (parameters == null)
                {
                    paraObjects = new object[] { macroName };
                }
                else
                {
                    // 宏参数组长度
                    int paraLength = parameters.Length;

                    paraObjects = new object[paraLength + 1];

                    paraObjects[0] = macroName;
                    for (int i = 0; i < paraLength; i++)
                    {
                        paraObjects[i + 1] = parameters[i];
                    }
                }

                // 创建Excel对象
                oExcel = new Excel.ApplicationClass();

                // 判断是否要求执行时Excel可见
                if (isShowExcel)
                {
                    // 使创建的对象可见
                    oExcel.Visible = true;
                }

                // 获取Workbooks对象
                oBooks = oExcel.Workbooks;

                // 打开指定的Excel文件，赋值Workbook对象
                oBook = oBooks.Open(
                    excelFilePath,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing,
                    oMissing
                );

                // 执行Excel中的宏
                rtnValue = this.RunMacro(oExcel, paraObjects);

                // 保存更改
                oBook.Save();

                // 退出Workbook
                oBook.Close(false, oMissing, oMissing);

                #endregion
            }
            catch (Exception ex)
            {
                #region 异常处理

                throw ex;

                #endregion
            }
            finally
            {
                #region 释放对象

                // 释放Workbook对象
                if (oBook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
                    oBook = null;
                }

                // 释放Workbooks对象
                if (oBooks != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
                    oBooks = null;
                }

                // 关闭Excel，并释放Excel对象
                if (oExcel != null)
                {
                    oExcel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                    oExcel = null;
                }

                // 调用垃圾回收
                GC.Collect();

                #endregion
            }
        }

        /// <summary>
        /// 执行宏
        /// </summary>
        /// <param name="oApp">Excel对象</param>
        /// <param name="oRunArgs">参数（第一个参数为指定宏名称，后面为指定宏的参数值）</param>
        /// <returns>宏返回值</returns>
        private object RunMacro(object oApp, object[] oRunArgs)
        {
            try
            {
                // 声明一个返回对象
                object objRtn;

                // 反射方式执行宏
                objRtn = oApp.GetType().InvokeMember(
                                                        "Run",
                                                        System.Reflection.BindingFlags.Default |
                                                        System.Reflection.BindingFlags.InvokeMethod,
                                                        null,
                                                        oApp,
                                                        oRunArgs
                                                     );

                // 返回值
                return objRtn;
            }
            catch (Exception ex)
            {
                // 如果有底层异常，抛出底层异常
                if (ex.InnerException.Message.ToString().Length > 0)
                {
                    throw ex.InnerException;
                }
                else
                {
                    throw ex;
                }
            }
        }
    }
}
 