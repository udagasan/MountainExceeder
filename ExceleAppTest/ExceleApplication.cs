using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExceleApplication;
using System.Collections.Generic;

namespace ExceleAppTest
{
    [TestClass]
    public class ExceleApplication
    {
    
        [TestMethod]
        public void ReadExistingExcel()
        {
            var obj = new Process();
            var day = DateTime.Now.Day.ToString();
            var month = DateTime.Now.Month.ToString();
            var year = DateTime.Now.Year.ToString();
            obj.ReadExistingExcel(GetTestDataList(), "USD", day, month, year);


        }

        List<ExcellData> GetTestDataList()
        {
            var dataList = new List<ExcellData>
            {
                new ExcellData
            {
                FullName = "WhiteStone",
                AccountNumber = "90390737",
                RegisterNumber = "14523698452",
                Iban = "TR555555555555555555555555",
                Amount="20"
            },
                new ExcellData
            {
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                Amount="20",
                RegisterNumber = "14523698452",
                Iban = "TR555555555555555555555555"

            },
            new ExcellData
            {
                Amount="20",
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                RegisterNumber = "14523698452",
                Iban = "TR555555555555555555555555"
            },
               new ExcellData
            {
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                RegisterNumber = "14523698452",
                Iban = "TR555555555555555555555555",
                Amount="20"
            },
                new ExcellData
            {
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                RegisterNumber = "14523698452",
                Iban = "TR555555555555555555555555"
            },
            new ExcellData
            {
                Amount="20",
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                RegisterNumber = "14523698452",
                Iban = "TR555555555555555555555555"
            },
        };

            return dataList;


        }

    }
}
