using ExcelDna.ComInterop;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace NewComServer
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return 
                @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                  <ribbon>
                    <tabs>
                      <tab idMso='TabAddIns'>
                        <group id='GroupExcelDnaCom' label='Excel-DNA COM'>
                          <button id='ButtonRegister' imageMso='HappyFace' size='large' label='Register' onAction='Register_Click'/>
                          <button id='ButtonUnregister' imageMso='Delete' size='large' label='Unregister' onAction='Unregister_Click'/>
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        public void Register_Click(IRibbonControl control)
        {
            ComServer.DllRegisterServer();
        }

        public void Unregister_Click(IRibbonControl control)
        {
            ComServer.DllUnregisterServer();
        }
    }
}
