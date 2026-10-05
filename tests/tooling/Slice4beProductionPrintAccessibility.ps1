# Excel's native preview ribbon is not always exposed through UI Automation.
function Initialize-ProductionPrintAccessibility {
    if('InvSysPrintPreviewButton' -as [type]){return}
    Add-Type -AssemblyName Accessibility
    Add-Type -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location) -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Accessibility;
public sealed class InvSysPrintPreviewButton {
    IAccessible owner; object child;
    InvSysPrintPreviewButton(IAccessible owner, object child){this.owner=owner;this.child=child;}
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
    delegate bool WindowCallback(IntPtr window,IntPtr state);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr window,WindowCallback callback,IntPtr state);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr window,uint id,ref Guid iid,[MarshalAs(UnmanagedType.Interface)] out IAccessible accessible);
    [DllImport("oleacc.dll")] static extern int AccessibleChildren(IAccessible parent,int start,int count,[Out,MarshalAs(UnmanagedType.LPArray,SizeParamIndex=2)] object[] children,out int obtained);
    static InvSysPrintPreviewButton Search(IAccessible accessible,object child,int depth,ref int remaining){
        if(accessible==null || depth>35 || --remaining<0)return null;
        try{
            string name=null;try{name=accessible.get_accName(child);}catch(COMException){}
            if(name!=null && name.Replace("\r"," ").Replace("\n"," ").Trim()=="Close Print Preview"){
                int state=Convert.ToInt32(accessible.get_accState(child));
                if((state & 0x18001)==0)return new InvSysPrintPreviewButton(accessible,child);
            }
            if(!(child is int) || (int)child!=0)return null;
            int count=Math.Min(accessible.accChildCount,2000),obtained;
            if(count==0)return null;
            object[] children=new object[count];
            if(AccessibleChildren(accessible,0,count,children,out obtained)<0)return null;
            for(int i=0;i<obtained;i++){
                IAccessible nested=children[i] as IAccessible;
                var result=nested!=null ? Search(nested,0,depth+1,ref remaining) : Search(accessible,children[i],depth+1,ref remaining);
                if(result!=null)return result;
            }
        }catch(COMException){}
        return null;
    }
    public static InvSysPrintPreviewButton Find(IntPtr window,uint expectedProcess){
        uint process;GetWindowThreadProcessId(window,out process);
        if(process!=expectedProcess)throw new InvalidOperationException("Preview window ownership changed.");
        var found=FindAt(window);
        if(found!=null)return found;
        EnumChildWindows(window,(nested,state)=>{found=FindAt(nested);return found==null;},IntPtr.Zero);
        return found;
    }
    static InvSysPrintPreviewButton FindAt(IntPtr window){
        Guid iid=new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");IAccessible root;
        if(AccessibleObjectFromWindow(window,0xFFFFFFFC,ref iid,out root)<0)return null;
        int remaining=5000;return Search(root,0,0,ref remaining);
    }
    public void Close(){owner.accDoDefaultAction(child);}
}
'@
}
