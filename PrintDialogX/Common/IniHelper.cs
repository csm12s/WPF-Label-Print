using IniParser;
using IniParser.Model;
using System;

namespace PrintDialogX.Common
{
    /// <summary>
    /// 
    /// </summary>
    public class IniHelper
    {
        #region Read and Write
        /// <summary>
        /// Read
        /// </summary>
        /// <param name="sectionName"></param>
        /// <param name="key"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static string GetIniData(string sectionName, string key, string filePath)
        {
            string value = "";
            try
            {
                FileIniDataParser iniParser = new FileIniDataParser();
                IniParser.Model.IniData iniData = iniParser.ReadFile(filePath);
                value = iniData[sectionName][key];
            }
            catch (Exception ex)
            {
                var errMsg = "Ini file read error, ";

                if(ex.Message.Contains("Unknown file format. Couldn't parse the line"))
                {
                    errMsg = "如果Ini文件看上去没问题，但是报错内容类似: " +
                            "Unknown file format. Couldn't parse the line: '﻿﻿[SectionName]'" +
                            "可能是Ini文件中有空格，可以用WinMerge检测出来, ";
                }

                throw new Exception(errMsg + ex.Message);
            }

            return value;
        }

        /// <summary>
        /// Write
        /// </summary>
        /// <param name="scetionName"></param>
        /// <param name="key"></param>
        /// <param name="saveStr"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool SaveIniData(string scetionName, string key,
            string saveStr, string filePath)
        {
            bool success = false;
            try
            {
                FileIniDataParser iniParser = new FileIniDataParser();
                IniData iniData = iniParser.ReadFile(filePath);

                iniData[scetionName][key] = saveStr;
                iniParser.WriteFile(filePath, iniData);

                success = true;
            }
            catch (Exception)
            {
                throw;
            }
            return success;
        }

        #endregion

    }
}


