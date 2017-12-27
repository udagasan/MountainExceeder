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
        public void CreateNewExcellAndFillFromDataTable()
        {
            var obj = new Process();
            obj.CreateNewExcellAndFillFromDataTable();


        }

        [TestMethod]
        public void ReadExistingExcel()
        {
            var obj = new Process();
            obj.ReadExistingExcel(GetTestDataList());


        }

        List<ExcellData> GetTestDataList()
        {
            var dataList = new List<ExcellData>
            {
                new ExcellData
            {
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                RegisterNumber = "8902",
                Iban = "işjsişodfjşso9ıerı*ewqrı",
                Currenc = "USD",
                Amount="20"
            },
                new ExcellData
            {
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                RegisterNumber = "8902",
                Amount="20",
                Iban = "işjsişodfjşso9ıerı*ewqrı",
                Currenc = "USD"
            },
            new ExcellData
            {
                Amount="20",
                FullName = "Uğur Dağaşan",
                AccountNumber = "90390737",
                RegisterNumber = "8902",
                Iban = "işjsişodfjşso9ıerı*ewqrı",
                Currenc = "USD"
            }
        };

            return dataList;


        }

    }
}
