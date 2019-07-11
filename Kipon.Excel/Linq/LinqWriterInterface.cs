﻿using System;

namespace Kipon.Excel.Linq
{
    public static class LinqWriterInterface
    {
        #region write to excel
        /// <summary>
        /// Transform information in T into an excel sheet, and write the excel data to the parsend excel stream
        /// 
        /// Based on reflection, this method will transform the data into excel, after a "best guess" approach
        /// using the duck type pattern.
        /// </summary>
        /// <typeparam name="T">T represents the type of data, see the full documentation on how you can interact by the streaming through your type</typeparam>
        /// <param name="data">an instance of T</param>
        /// <returns></returns>
        public static void ToExcel<T>(this T data, System.IO.Stream excel)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Transform information in T into excel sheet, and return the openxml document as byte array 
        /// This method is using an im memory strategy so be carefull about size
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static byte[] ToExcel<T>(this T data) 
        {
            using (var mem = new System.IO.MemoryStream())
            {
                LinqWriterInterface.ToExcel<T>(data, mem);
                return mem.ToArray();
            }
        }


        #endregion
    }
}
