using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using EdmLib;
using System.Windows.Forms;

namespace AccessToBom_Cube
{
    [Guid("B2C9AB54-D4D1-441F-8B59-FBB3FC1FEFFB")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISpec
    {
        [DispId(1)]
        string[] GetListBom(string root);

    }
    [Guid("4C84858D-190B-4236-B3BC-23AAAE6A84C6")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccessToBom_Cube.Spec")]
    public class Spec : ISpec

    {
        IEdmVault5 vault5 = null;
        IEdmVault7 vault7 = null;
        int hwnd = 0;
        List<string> listBom = null;
       public Spec()
        {
          vault5 = new EdmVault5();
        }
 
        public string[] GetListBom(string root)
        {
            string[] res = null;
            IEdmFile7 File = null;
            IEdmFolder5 ParentFolder = null;
            IEdmBomMgr2 edmBomMgr = null;
             EdmBomLayout2[] ppoRetLayouts = null;
            Array ppoRetLayouts1 = null;
            EdmBomLayout2 ppoRetLayout = default(EdmBomLayout2);
            IEdmBomView3 bomView;
          
            try
            {

                vault5.LoginAuto("My", hwnd);
                vault7 = (IEdmVault7)vault5;
                File = (IEdmFile7)vault7.GetFileFromPath(root, out ParentFolder);
                if (File != null)
                {
                    listBom = new List<string>();
                    listBom.Add(root);
                    vault7 = (IEdmVault7)vault5;
                    edmBomMgr =(IEdmBomMgr2) vault7.CreateUtility(EdmUtility.EdmUtil_BomMgr);
                    edmBomMgr.GetBomLayouts2(out ppoRetLayouts1);
                    ppoRetLayouts = (EdmBomLayout2[])ppoRetLayouts1;


                    int i = 0;
                    int arrSize = ppoRetLayouts.Length;

                    while (i < arrSize)
                    {

                        if (ppoRetLayouts[i].mbsLayoutName == "Спецификации")
                        {
                            ppoRetLayout = ppoRetLayouts[i];
                            break;
                        }
                        i = i + 1;

                    }

                    bomView = (IEdmBomView3)File.GetComputedBOM(ppoRetLayout.mbsLayoutName, 0, "@", (int)EdmBomFlag.EdmBf_AsBuilt);

                 
                    /*
                    EdmBomColumn[] ppoColumns = null;

                    bomView.GetColumns(out ppoColumns);
                    i = 0;
                    arrSize = ppoColumns.Length;
              
                    while (i < arrSize)
                    {
                        if (ppoColumns[i].mbsCaption == "Имя файла")
                        {

                        }
                        i = i + 1;
                    }
                    */

                    object[] ppoRows = null;
                    Array ppoRows1 = null;
                    IEdmBomCell ppoRow = default(IEdmBomCell);

                    string tempPath = "";
                    bomView.GetRows(out ppoRows1);
                    ppoRows =(object[]) ppoRows1;
                    i = 0;
                    arrSize = ppoRows.Length;

           
                    while (i < arrSize-1)
                    {
                        ppoRow = (IEdmBomCell)ppoRows[i];
                        tempPath = ppoRow.GetPathName();
                        listBom.Add(tempPath);
                        i = i + 1;
                    }

                    int count = listBom.Count;
                    res = listBom.ToArray();

                }

            }
            catch (Exception)
            {

                throw;
            }
            return res;
   
        }

      
    }
}

