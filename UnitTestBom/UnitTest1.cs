using Microsoft.VisualStudio.TestTools.UnitTesting;
using AccessToBom_Cube;

namespace UnitTestBom
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string[] result = null;
            string pathFile = @"A:\My\10-тонный электрический подъемный кран\ramka-1.SLDDRW";
            ISpec sp = new Spec();
            result = sp.GetListBom(pathFile);
            Assert.AreEqual(result.Length, 2);
        }
    }
}
