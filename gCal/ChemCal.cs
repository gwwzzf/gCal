using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelDna.IntelliSense;
using ExcelDna.Integration;


namespace Github_gCal
{

    public class AddIn : IExcelAddIn
    {

        public void AutoOpen()
        {

            ///加载插件
            IntelliSenseServer.Install();

        }
        public void AutoClose()
        {
            ///卸载插件
            IntelliSenseServer.Uninstall();
        }
    }



    public class ChemCal
    {

        /// <summary>
        /// 化学元素符号计算
        /// </summary>
        /// <param name="n">输入化学元素序号数字</param>
        /// <returns></returns>
        [ExcelFunction(Description = "化学元素符号", IsMacroType = true, Category = "化学式计算")]
        public static string gCal_ChemSymbol_No([ExcelArgument(Name = "元素序号", Description = "输入数字")] string n)
        {

            DataBaseG data = new DataBaseG();
            //Regex.IsMatch(n, @"^\d+$") 判断n是否为数字
            if (Regex.IsMatch(n, @"^\d+$") && Convert.ToInt16(n) >= 1 && Convert.ToInt16(n) <= data.Chem_symbols.Length)
            {
                return data.Chem_symbols[Convert.ToInt16(n) - 1];
            }
            return "";
        }
        /// <summary>
        /// 化学元素名称
        /// </summary>
        /// <param name="str">化学元素符号</param>
        /// <returns></returns>
        [ExcelFunction(Description = "化学元素名称", IsMacroType = true, Category = "化学式计算")]
        public static string gCal_ChemName_Symbol([ExcelArgument(Name = "化学元素符号", Description = "输入化学元素字母,注意大小写")] string str)
        {

            DataBaseG data = new DataBaseG();
            //Regex.IsMatch(str, @"^[A-Za-z]+$")  判断str是否为字母，在上述封装的方法中，正则表达式[A-Za-z]表示匹配英文大写字母A到Z，以及英文小写字母a到z，加号+表示匹配一个到多个。
            if (Regex.IsMatch(str, @"^[A-Za-z]+$") && str.Length >= 1 && str.Length <= 2)
            {
                for (int i = 0; i < data.Chem_symbols.Length; i++)
                {
                    if (data.Chem_symbols[i] == str)
                    {
                        return data.Chem_name[i];
                        // break;
                    }
                    //continue;
                }
            }
            return "";
        }
        /// <summary>
        /// 化学元素相对原子质量
        /// </summary>
        /// <param name="str">化学元素符号</param>
        /// <returns></returns>
        [ExcelFunction(Description = "化学元素相对原子质量", IsMacroType = true, Category = "化学式计算")]
        public static double gCal_ChemAR_Symbol([ExcelArgument(Name = "化学元素符号", Description = "输入化学元素字母,注意大小写")] string str)
        {

            DataBaseG data = new DataBaseG();
            //Regex.IsMatch(str, @"^[A-Za-z]+$")  判断str是否为字母，在上述封装的方法中，正则表达式[A-Za-z]表示匹配英文大写字母A到Z，以及英文小写字母a到z，加号+表示匹配一个到多个。
            if (Regex.IsMatch(str, @"^[A-Za-z]+$") && str.Length >= 1 && str.Length <= 2)
            {
                for (int i = 0; i < data.Chem_symbols.Length; i++)
                {
                    if (data.Chem_symbols[i] == str)
                    {
                        return data.Chem_AR[i];
                        //break;
                    }
                    // continue;

                }
            }
            return 0;
        }
        /// <summary>
        /// 无括号化学式相对原子质量计算
        /// </summary>
        /// <param name="str">化学分子式</param>
        /// <returns></returns>

        private static double Simple_MR_Formula([ExcelArgument(Name = "化学分子式", Description = "输入化学分子式,注意大小写和括号的书写方式")] string str)
        {

            DataBaseG data = new DataBaseG();
            //Regex.IsMatch(str, @"^[A-Za-z]+$")  判断str是否为字母，在上述封装的方法中，正则表达式[A-Za-z]表示匹配英文大写字母A到Z，以及英文小写字母a到z，加号+表示匹配一个到多个。
            if (string.IsNullOrEmpty(str) == false) //string.IsNullOrEmpty(str)  既可以判断string.Empty 又可以判断出null
            {

                string str1 = str.Substring(0, 1);//获取首字母
                for (int i = 1; i < str.Length; i++) // 在化学式大写字母左侧增加 - ，作为分割符号
                {
                    if (Regex.IsMatch(str.Substring(i, 1), @"^[A-Z]+$"))
                    {
                        str1 = str1 + "-" + str.Substring(i, 1);
                    }
                    else
                    {
                        str1 = str1 + str.Substring(i, 1);
                    }
                }
                // Debug.WriteLine(str1);
                String[] arrT1 = str1.Split('-');//以 -  分割字符串
                double mr = 0, check = 0;
                for (int j = 0; j < arrT1.Length; j++)
                {
                    if (arrT1[j].Length == 1)
                    {
                        check = gCal_ChemAR_Symbol(arrT1[j]);
                        mr = mr + check;

                    }
                    else if (arrT1[j].Length == 2)
                    {
                        if (Regex.IsMatch(arrT1[j].Substring(1, 1), @"^\d+$"))//判断第二位是否为数字
                        {
                            check = gCal_ChemAR_Symbol(arrT1[j].Substring(0, 1));
                            mr = mr + check * Convert.ToInt32(arrT1[j].Substring(1, 1));

                        }
                        else
                        {
                            check = gCal_ChemAR_Symbol(arrT1[j].Substring(0, 2));
                            mr = mr + check;

                        }

                    }
                    else if (arrT1[j].Length == 3)
                    {
                        if (Regex.IsMatch(arrT1[j].Substring(1, 2), @"^\d+$"))//判断第二、三位是否为数字
                        {
                            check = gCal_ChemAR_Symbol(arrT1[j].Substring(0, 1));
                            mr = mr + check * Convert.ToInt16(arrT1[j].Substring(1, 2));

                        }
                        else if (Regex.IsMatch(arrT1[j].Substring(2, 1), @"^\d+$"))//判断第三位是否为数字
                        {
                            check = gCal_ChemAR_Symbol(arrT1[j].Substring(0, 2));
                            mr = mr + check * Convert.ToInt16(arrT1[j].Substring(2, 1));

                        }
                        else { mr = 0; }

                    }
                    else if (arrT1[j].Length == 4)
                    {
                        if (Regex.IsMatch(arrT1[j].Substring(1, 3), @"^\d+$"))//判断第二、三、四位是否为数字
                        {

                            check = gCal_ChemAR_Symbol(arrT1[j].Substring(0, 1));
                            mr = mr + check * Convert.ToInt16(arrT1[j].Substring(1, 3));

                        }
                        else if (Regex.IsMatch(arrT1[j].Substring(2, 2), @"^\d+$"))//判断第三、四位是否为数字
                        {
                            check = gCal_ChemAR_Symbol(arrT1[j].Substring(0, 2));
                            mr = mr + check * Convert.ToInt16(arrT1[j].Substring(2, 2));

                        }
                        else { mr = 0; }

                    }
                    else if (arrT1[j].Length > 4)
                    { mr = 0; }


                    if (check == 0) { mr = 0; break; }

                }

                return mr;
            }
            return 0;
        }
        /// <summary>
        /// 化学式相对分子质量计算
        /// </summary>
        /// <param name="str">化学式</param>
        /// <returns></returns>
        [ExcelFunction(Description = "化学式相对分子质量计算", IsMacroType = true, Category = "化学式计算")]
        public static Double gCal_ChemMR_Formula([ExcelArgument(Name = "化学式", Description = "输入化学式,需注意大小写和括号的书写方式")] string str)
        {

            DataBaseG data = new DataBaseG();
            string str1 = "", Fstr = "", Sstr = "";
            int n1 = 0, n2 = 0, n3 = 0;
            Double mr = 0;

            if (string.IsNullOrEmpty(str) == false) //string.IsNullOrEmpty(str)  既可以判断string.Empty 又可以判断出null
            {
                str1 = str.Replace("（", "(");
                str1 = str1.Replace("）", ")");//两个中文括号计算，采用replace替换

                if (str1.Contains("("))
                {
                    n1 = str1.IndexOf("(");
                    n2 = str1.IndexOf(")");
                    if (n2 + 1 == str1.Length)
                    {
                        goto Exit;
                    }

                    if (n2 + 3 > str1.Length) //如果）在末尾，只能算括号后一个数字
                    {

                        if (Regex.IsMatch(str1.Substring(n2 + 1, 1), @"^\d+$"))
                        {
                            n3 = Convert.ToInt16(str1.Substring(n2 + 1, 1));
                            Fstr = str1.Substring(n1, n2 - n1 + 2);
                        }

                        else { goto Exit; }

                    }
                    else
                    {

                        if (Regex.IsMatch(str1.Substring(n2 + 1, 2), @"^\d+$")) //如果）右边两位是数字，说明括号内数字是两位数
                        {
                            n3 = Convert.ToInt16(str1.Substring(n2 + 1, 2));
                            Fstr = str1.Substring(n1, n2 - n1 + 3);
                        }
                        else if (Regex.IsMatch(str1.Substring(n2 + 1, 1), @"^\d+$")) //如果）右边一位是数字，说明括号内数字是两位数
                        {
                            n3 = Convert.ToInt16(str1.Substring(n2 + 1, 1));
                            Fstr = str1.Substring(n1, n2 - n1 + 2);
                        }
                        else { goto Exit; }
                    }
                    Sstr = str1.Replace(Fstr, "");
                    Fstr = str1.Substring(n1 + 1, n2 - n1 - 1);

                    mr = Simple_MR_Formula(Sstr) + n3 * Simple_MR_Formula(Fstr);
                    if (Simple_MR_Formula(Sstr) == 0 || Simple_MR_Formula(Fstr) == 0)
                    {
                        mr = 0;
                    }

                    return mr;
                }
                else
                {
                    mr = Simple_MR_Formula(str);
                    return mr;
                }

            }
        Exit:
            return 0;
        }



    }
}
