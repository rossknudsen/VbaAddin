using System.Reflection;


public static class WindowsExtension
{
    public static NetOffice.VBIDEApi.Window CreateToolWindowEx(this NetOffice.VBIDEApi._Windows windows, NetOffice.VBIDEApi.AddIn addInInst, string progId, string caption, string guidPosition, ref object docObj)
    {
        ParameterModifier[] modifiers = NetOffice.Invoker.CreateParamModifiers(false, false, false, false, true);
        object[] paramsArray = NetOffice.Invoker.ValidateParamsArray(addInInst, progId, caption, guidPosition, docObj);
        object returnItem = NetOffice.Invoker.MethodReturn(windows, "CreateToolWindow", paramsArray, modifiers);
        docObj = paramsArray[4];
        NetOffice.VBIDEApi.Window newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(windows, returnItem, NetOffice.VBIDEApi.Window.LateBindingApiWrapperType) as NetOffice.VBIDEApi.Window;
        return newObject;
    }
}