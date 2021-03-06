﻿Imports System.Reflection
Imports NetOffice.Invoker


Public Module WindowsExtension
    <System.Runtime.CompilerServices.Extension> _
    Public Function CreateToolWindowEx(windows As NetOffice.VBIDEApi._Windows, addInInst As NetOffice.VBIDEApi.AddIn, progId As String, caption As String, guidPosition As String, ByRef docObj As Object) As NetOffice.VBIDEApi.Window
        Dim modifiers As ParameterModifier() = NetOffice.Invoker.CreateParamModifiers(False, False, False, False, True)
        Dim paramsArray As Object() = NetOffice.Invoker.ValidateParamsArray(addInInst, progId, caption, guidPosition, docObj)
        Dim returnItem As Object = NetOffice.Invoker.MethodReturn(windows, "CreateToolWindow", paramsArray, modifiers)
        docObj = paramsArray(4)
        Dim newObject As NetOffice.VBIDEApi.Window = TryCast(NetOffice.Factory.CreateKnownObjectFromComProxy(windows, returnItem, NetOffice.VBIDEApi.Window.LateBindingApiWrapperType), NetOffice.VBIDEApi.Window)
        Return newObject
    End Function
End Module