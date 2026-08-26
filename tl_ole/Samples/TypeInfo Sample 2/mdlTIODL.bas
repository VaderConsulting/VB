Attribute VB_Name = "mdlTIODL"
Option Explicit

Dim IID_NULL As UUID

Dim m_oImportedLibs As Collection
Dim m_oTypes As Collection

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Const KEY_QUERY_VALUE = &H1

Private Const REG_SZ = 1

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'
' DecompileConsts
'
' Decompile constants in an enum or module
'
' Parameters:
'
' oTI - ITypeInfo of the enum or module
' tTA - TYPEATTR of the enum or module
'
Private Function DecompileConsts(ByVal oTI As ITypeInfo, tTA As TYPEATTR) As String
Dim bIsModule As Boolean
Dim lIdx As Long
Dim tVD As VARDESC
Dim lPtr As Long
Dim sName As String
Dim sHelp As String
Dim sHlpFile As String
Dim lHlpCtx As Long
Dim sAttrs As String

   ' Check if the ITypeInfo is a module
   bIsModule = tTA.TYPEKIND = TKIND_MODULE
   
   ' Enumerate all constants
   For lIdx = 0 To tTA.cVars - 1
         
      ' Get a pointer to the VARDESC
      ' struct for this constant
      lPtr = oTI.GetVarDesc(lIdx)
         
      ' Copy from the pointer
      MoveMemory tVD, ByVal lPtr, LenB(tVD)
         
      ' Get the name
      oTI.GetDocumentation tVD.memid, sName, sHelp, lHlpCtx, sHlpFile
      sName = RTrimNull(sName)

      ' Clear the attributes
      sAttrs = vbNullString
      
      ' Add the helpfile and context (if any)
      If LenB(sHelp) <> 0 Then sAttrs = sAttrs & "helpstring(""" & RTrimNull(sHelp) & """), "
      If lHlpCtx <> 0 Then sAttrs = sAttrs & "helpcontext(0x" & Hex$(lHlpCtx) & "), "
      
      ' Add the attributes
      If LenB(sAttrs) <> 0 Then DecompileConsts = DecompileConsts & "    [" & Left$(sAttrs, Len(sAttrs) - 2) & "]" & vbCrLf
      
      ' Resolve the type and value of the const
      If bIsModule Then
         DecompileConsts = DecompileConsts & "    const " & ResolveVarType(oTI, tVD.elemdescVar.tdesc, sName) & ResolveConstVal(oTI, lIdx, "") & ";" & vbCrLf
      Else
         DecompileConsts = DecompileConsts & "    " & ResolveConstVal(oTI, lIdx, sName) & "," & vbCrLf
      End If
      
      ' Release the pointer
      oTI.ReleaseFuncDesc lPtr
         
   Next

End Function
'
' DecompileFunctions
'
' Decompiles functions in a module or interface
'
' Parameters:
'
' oTI - The module/interface ITypeInfo
'
Private Function DecompileFunctions(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim tFD As FUNCDESC
Dim lIdx As Long
Dim lPtr As Long
Dim lHlpCtx As Long
Dim iOrd As Integer
Dim bIsModule As Boolean
Dim sName As String
Dim sDllEntry As String
Dim sHelp As String
Dim sAttrs As String

   ' Get the TYPEATTR
   tTA = GetTypeAttr(oTI)
   
   ' Check if this a module
   bIsModule = tTA.TYPEKIND = TKIND_MODULE
   
   ' Enumerate all functions
   For lIdx = 0 To tTA.cFuncs - 1
      
      ' Get the function FUNCDESC
      lPtr = oTI.GetFuncDesc(lIdx)
      MoveMemory tFD, ByVal lPtr, LenB(tFD)
      
      ' Get the function name, helpstring and helpcontext
      oTI.GetDocumentation tFD.memid, sName, sHelp, lHlpCtx, vbNullString
      
      ' Get the attributes
      sAttrs = FuncFlags(tFD.wFuncFlags)
      If LenB(sHelp) <> 0 Then sAttrs = sAttrs & "helpstring(""" & RTrimNull(sHelp) & """), "
      If lHlpCtx <> 0 Then sAttrs = sAttrs & "helpcontext(0x" & Hex$(lHlpCtx) & "), "
      
      ' Add the invoke kind
      Select Case tFD.invkind
         Case INVOKE_PROPERTYGET
            sAttrs = sAttrs & "propget, "
         Case INVOKE_PROPERTYPUT
            sAttrs = sAttrs & "propput, "
         Case INVOKE_PROPERTYPUTREF
            sAttrs = sAttrs & "propputref, "
      End Select
      
      ' Add the DISPID
      If tTA.TYPEKIND = TKIND_DISPATCH Or tTA.TYPEKIND = TKIND_INTERFACE Then sAttrs = sAttrs & "id(0x" & Hex$(tFD.memid) & "), "
      
      ' If this is a module get the DLL entry
      If bIsModule Then
      
         ' Get DLL entry
         oTI.GetDllEntry tFD.memid, INVOKE_FUNC Or INVOKE_PROPERTYGET, vbNullString, sDllEntry, iOrd
         
         ' Add the function name or ordinal
         sAttrs = sAttrs & "entry("
         If LenB(sDllEntry) <> 0 Then
            sAttrs = sAttrs & """" & sDllEntry & """"
         Else
            sAttrs = sAttrs & "0x" & iOrd
         End If
         sAttrs = sAttrs & "), "
         
      End If
      
      If tFD.cParamsOpt = -1 Then sAttrs = sAttrs & "vararg, "
      
      ' Add the attributes
      If LenB(sAttrs) <> 0 Then DecompileFunctions = DecompileFunctions & "    [" & Left$(sAttrs, Len(sAttrs) - 2) & "]" & vbCrLf
      
      ' Resolve the return value
      DecompileFunctions = DecompileFunctions & "    " & ResolveVarType(oTI, tFD.elemdescFunc.tdesc, sName) & "("
      
      ' Decompile the parameters
      DecompileFunctions = DecompileFunctions & DecompileParams(oTI, tFD)
      
      ' Close the function
      DecompileFunctions = DecompileFunctions & ");" & vbCrLf & vbCrLf
      
      ' Release the pointer to FUNCDESC
      oTI.ReleaseFuncDesc lPtr
   
   Next

End Function
'
' DecompileParams
'
' Decompiles functions parameters
'
' Parameters:
'
' oTI - ITypeInfo of the module/interface where the function is
' tFD - FUNCDESC of the function
'
Private Function DecompileParams(ByVal oTI As ITypeInfo, tFD As FUNCDESC) As String
Dim tED As ELEMDESC
Dim lIdx As Long
Dim sAttrs As String
ReDim sNames(-1 To tFD.cParams) As String

   ' Decompile only if there're parameters
   If tFD.cParams Then
      
      DecompileParams = vbCrLf
      
      ' Get the parameter names
      oTI.GetNames tFD.memid, sNames(-1), tFD.cParams + 1
      
      For lIdx = 0 To tFD.cParams - 1
         
         ' Get the parameter description
         MoveMemory tED, ByVal tFD.lprgELEMDESCParam + LenB(tED) * lIdx, LenB(tED)
         
         DecompileParams = DecompileParams & "        "
         
         ' Add the attributes
         sAttrs = vbNullString
         If tED.PARAMDESC.wParamFlags And PARAMFLAG_FIN Then sAttrs = sAttrs & "in, "
         ' BUG IN OLELIB.TLB:  PARAMFLAG_FLONG should be PARAMFLAG_FLCID
         If tED.PARAMDESC.wParamFlags And PARAMFLAG_FLONG Then sAttrs = sAttrs & "lcid, "
         If tED.PARAMDESC.wParamFlags And PARAMFLAG_FOUT Then sAttrs = sAttrs & "out, "
         If tED.PARAMDESC.wParamFlags And PARAMFLAG_FRETVAL Then sAttrs = sAttrs & "retval, "
         If tED.PARAMDESC.wParamFlags And PARAMFLAG_FOPT Then
         
            ' ODL doesn't supports the optional attribute so
            ' if there's no default value use 0.
            If (tED.PARAMDESC.wParamFlags And PARAMFLAG_FHASDEFAULT) = 0 Then
               sAttrs = sAttrs & "defaultvalue(0), "
            End If
            
         End If
         
         ' Get the defaultvalue if the parameter is optional
         If tED.PARAMDESC.wParamFlags And PARAMFLAG_FHASDEFAULT Then
            Dim vVar As Variant
            
            ' Get the value
            MoveMemory vVar, ByVal tED.PARAMDESC.pPARAMDESCEX + 8, 16
            
            ' Add the value
            If VarType(vVar) = vbString Then
               sAttrs = sAttrs & "defaultvalue(""" & vVar & """), "
            Else
               sAttrs = sAttrs & "defaultvalue(" & vVar & "), "
            End If
            
            ' Clear the variant without it releasing it content
            MoveMemory vVar, ByVal 0&, 2&
            
         End If
         If LenB(sAttrs) <> 0 Then DecompileParams = DecompileParams & "[" & Left$(sAttrs, Len(sAttrs) - 2) & "] "
         
         ' Resolve the parameter type
         If sNames(lIdx) = "" Then sNames(lIdx) = "Param" & lIdx
         DecompileParams = DecompileParams & ResolveVarType(oTI, tED.tdesc, sNames(lIdx))
         
         If lIdx < tFD.cParams - 1 Then DecompileParams = DecompileParams & "," & vbCrLf
         
      Next
      
   Else
   
      ' The function does not have parameters
      DecompileParams = DecompileParams & "void"
      
   End If

End Function
'
' DecompileProps
'
' Decompiles a dispinterface properties
'
' Parameter
'
' oTI - The dispinterface ITypeInfo
'
Private Function DecompileProps(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim tVD As VARDESC
Dim lIdx As Long
Dim lPtr As Long
Dim lHlpCtx As Long
Dim sName As String
Dim sHelp As String
Dim sAttrs As String

   ' Get the TYPEATTR
   tTA = GetTypeAttr(oTI)
   
   DecompileProps = DecompileProps & "properties:" & vbCrLf
   
   ' Enumerate all properties
   For lIdx = 0 To tTA.cVars - 1
      
      ' Get FUNCDESC
      lPtr = oTI.GetVarDesc(lIdx)
      MoveMemory tVD, ByVal lPtr, LenB(tVD)
      
      ' Get function name and help
      oTI.GetDocumentation tVD.memid, sName, sHelp, lHlpCtx, vbNullString

      ' Add attributes
      sAttrs = VarFlags(tVD.wVarFlags)
      If LenB(sHelp) <> 0 Then sAttrs = sAttrs & "helpstring(""" & RTrimNull(sHelp) & """), "
      If lHlpCtx <> 0 Then sAttrs = sAttrs & "helpcontext(0x" & Hex$(lHlpCtx) & "), "
      If tTA.TYPEKIND = TKIND_DISPATCH Then sAttrs = sAttrs & "id(0x" & Hex$(tVD.memid) & "), "
      If LenB(sAttrs) <> 0 Then DecompileProps = DecompileProps & "    [" & Left$(sAttrs, Len(sAttrs) - 2) & "]" & vbCrLf
      
      ' Resolve property type
      DecompileProps = DecompileProps & "    " & ResolveVarType(oTI, tVD.elemdescVar.tdesc, sName) & ";" & vbCrLf & vbCrLf
      
      oTI.ReleaseVarDesc lPtr
   
   Next

End Function
'
' GetDllName
'
' Returns the DLL name of a module
'
' Parameters:
'
' oTI - The module ITypeInfo
'
Private Function GetDllName(ByVal oTI As ITypeInfo) As String
Dim tFD As FUNCDESC
Dim lPtr As Long

   On Error Resume Next
   
   ' Get the first function FUNCDESC
   lPtr = oTI.GetFuncDesc(0)
   
   If Err.Number <> 0 Then
   
      ' There's no function in the module
      GetDllName = "nodll"
      
   Else
      
      MoveMemory tFD, ByVal lPtr, LenB(tFD)
      
      ' Get the module name
      oTI.GetDllEntry tFD.memid, INVOKE_FUNC Or INVOKE_PROPERTYGET, GetDllName, vbNullString, 0
      
      oTI.ReleaseFuncDesc lPtr
      
   End If
   
End Function

Function DecompileTypeLib(ByVal TL As Variant) As String
Dim oTL As ITypeLib
Dim oTI As ITypeInfo
Dim lTKind As Long
Dim lIdx As Long
Dim lPtr As Long
Dim tLA As TLIBATTR
Dim sName As String
Dim sHelp As String
Dim lHlpCtx As Long
Dim sHlpFile As String

   If VarType(TL) = vbString Then
   
      ' Load the typelib
      Set oTL = LoadTypeLibEx(TL, REGKIND_NONE)
      
   Else
   
      Set oTL = TL
      
   End If
   
   DecompileTypeLib = DecompileTypeLib & "// ODL Extracted by Edanmo's ITypeInfo Sample" & vbCrLf
      
   lPtr = oTL.GetLibAttr
   MoveMemory tLA, ByVal lPtr, LenB(tLA)
   
   DecompileTypeLib = DecompileTypeLib & "[" & vbCrLf
   
   sName = Space$(38)
   StringFromGUID2 tLA.iid, sName, 39
   DecompileTypeLib = DecompileTypeLib & "    uuid(" & Mid$(sName, 2, 36) & "), " & vbCrLf
   
   oTL.GetDocumentation -1, sName, sHelp, lHlpCtx, sHlpFile
   
   If tLA.wMajorVerNum <> 0 Or tLA.wMinorVerNum <> 0 Then DecompileTypeLib = DecompileTypeLib & "    version(" & tLA.wMajorVerNum & "." & tLA.wMinorVerNum & "), " & vbCrLf
   If LenB(sHelp) <> 0 Then DecompileTypeLib = DecompileTypeLib & "    helpstring(""" & RTrimNull(sHelp) & """), " & vbCrLf
   If LenB(sHlpFile) <> 0 Then DecompileTypeLib = DecompileTypeLib & "    helpfile(""" & RTrimNull(sHlpFile) & """), " & vbCrLf
   If lHlpCtx <> 0 Then DecompileTypeLib = DecompileTypeLib & "    helpcontext(0x" & Hex$(lHlpCtx) & "), " & vbCrLf
   
   If tLA.wLibFlags And LIBFLAG_FRESTRICTED Then DecompileTypeLib = DecompileTypeLib & "    restricted, " & vbCrLf
   If tLA.wLibFlags And LIBFLAG_FCONTROL Then DecompileTypeLib = DecompileTypeLib & "    control, " & vbCrLf
   If tLA.wLibFlags And LIBFLAG_FHIDDEN Then DecompileTypeLib = DecompileTypeLib & "    hidden, " & vbCrLf
   
   If tLA.lcid <> 0 Then DecompileTypeLib = DecompileTypeLib & "    lcid(0x" & Hex$(tLA.lcid) & "), " & vbCrLf
   
   DecompileTypeLib = DecompileTypeLib & "]" & vbCrLf
   DecompileTypeLib = DecompileTypeLib & "library " & sName & " {" & vbCrLf
   
   oTL.ReleaseTLibAttr lPtr
   
   Set m_oImportedLibs = New Collection
   Set m_oTypes = New Collection
   
   ' Enumerate all typelib members
   For lIdx = 0 To oTL.GetTypeInfoCount - 1
   
      ' Get member type
      lTKind = oTL.GetTypeInfoType(lIdx)
      
      ' Get member TypeInfo
      Set oTI = oTL.GetTypeInfo(lIdx)
      
      Select Case lTKind
      
         Case TKIND_COCLASS
            
            DecompileTypeLib = DecompileTypeLib & DecompileCoclass(oTI)
                  
         Case TKIND_INTERFACE, TKIND_DISPATCH
         
            DecompileTypeLib = DecompileTypeLib & DecompileInterface(oTI)
                  
         Case TKIND_MODULE
            
            DecompileTypeLib = DecompileTypeLib & DecompileModule(oTI)
                        
         Case TKIND_ENUM
            
            DecompileTypeLib = DecompileTypeLib & DecompileEnum(oTI)

         Case TKIND_ALIAS
         
            DecompileTypeLib = DecompileTypeLib & DecompileTypedef(oTI)
         
         Case TKIND_RECORD
            
            DecompileTypeLib = DecompileTypeLib & DecompileStruct(oTI)
            
         Case TKIND_UNION
         
            DecompileTypeLib = DecompileTypeLib & DecompileUnion(oTI)
            
      End Select

      DecompileTypeLib = DecompileTypeLib & vbCrLf
      
   Next

   Dim slibs As String, vVar As Variant
   
   For Each vVar In m_oImportedLibs
      slibs = slibs & "importlib(""" & vVar & """);" & vbCrLf
   Next
   
   slibs = slibs & vbCrLf & vbCrLf
   slibs = slibs & "// Forward declare all types" & vbCrLf
   For Each vVar In m_oTypes
      slibs = slibs & vVar & ";" & vbCrLf
   Next
   
   lIdx = InStr(DecompileTypeLib, "library " & sName)
   DecompileTypeLib = Left$(DecompileTypeLib, lIdx + 11 + Len(sName)) & vbCrLf & slibs & vbCrLf & Mid$(DecompileTypeLib, lIdx + 12 + Len(sName))
   
   DecompileTypeLib = DecompileTypeLib & "};"
   
End Function
Private Function ResolveTypeLib(TL As ITypeLib) As String
Dim hKey As Long
Dim LA As TLIBATTR
Dim lPtr As Long
Dim sLID As String
Dim lIdx As Long

   lPtr = TL.GetLibAttr
   MoveMemory LA, ByVal lPtr, LenB(LA)
   
   sLID = Space$(38)
   StringFromGUID2 LA.iid, sLID, 39
   
   sLID = "TypeLib\" & sLID & "\" & LA.wMajorVerNum & "." & LA.wMinorVerNum & "\" & LA.lcid & "\win32"
   
   If RegOpenKeyEx(HKEY_CLASSES_ROOT, sLID, 0, KEY_QUERY_VALUE, hKey) = 0 Then
      ResolveTypeLib = Space$(260)
      RegQueryValueExStr hKey, vbNullString, 0, REG_SZ, ResolveTypeLib, 260
      
      ResolveTypeLib = Left$(ResolveTypeLib, InStr(ResolveTypeLib, vbNullChar) - 1)
      
      For lIdx = Len(ResolveTypeLib) To 1 Step -1
         If Mid$(ResolveTypeLib, lIdx, 1) = "\" Then
            ResolveTypeLib = LCase$(Mid$(ResolveTypeLib, lIdx + 1))
            Exit For
         End If
      Next
      
      RegCloseKey hKey
   End If
   
   TL.ReleaseTLibAttr lPtr
   
End Function


Private Function FuncFlags(ByVal Flags As Long) As String

   If (Flags And FUNCFLAG_FBINDABLE) Then FuncFlags = FuncFlags & "bindable, "
   If (Flags And FUNCFLAG_FDEFAULTBIND) Then FuncFlags = FuncFlags & "defaultbind, "
   If (Flags And FUNCFLAG_FDEFAULTCOLLELEM) Then FuncFlags = FuncFlags & "defaultcollelem, "
   If (Flags And FUNCFLAG_FDISPLAYBIND) Then FuncFlags = FuncFlags & "displaybind, "
   If (Flags And FUNCFLAG_FHIDDEN) Then FuncFlags = FuncFlags & "hidden, "
   If (Flags And FUNCFLAG_FIMMEDIATEBIND) Then FuncFlags = FuncFlags & "immediatebind, "
   If (Flags And FUNCFLAG_FNONBROWSABLE) Then FuncFlags = FuncFlags & "nonbrowsable, "
   ' If (Flags And FUNCFLAG_FREPLACEABLE) Then FuncFlags = FuncFlags & "replaceable, "
   If (Flags And FUNCFLAG_FREQUESTEDIT) Then FuncFlags = FuncFlags & "requestedit, "
   If (Flags And FUNCFLAG_FRESTRICTED) Then FuncFlags = FuncFlags & "restricted, "
   If (Flags And FUNCFLAG_FSOURCE) Then FuncFlags = FuncFlags & "source, "
   If (Flags And FUNCFLAG_FUIDEFAULT) Then FuncFlags = FuncFlags & "uidefault, "
   
End Function

Public Function TypeInfoFromObject(ByVal Obj As olelib.IDispatch) As ITypeInfo

   On Error Resume Next
   
   Set TypeInfoFromObject = Obj.GetTypeInfo(0, 0)

   If TypeInfoFromObject Is Nothing Then
      Dim PCI As IProvideClassInfo
      
      Set PCI = Obj
      
      Set TypeInfoFromObject = PCI.GetClassInfo
      
   End If
   
End Function


Private Function VarFlags(ByVal Flags As Long) As String

   If (Flags And VARFLAG_FBINDABLE) Then VarFlags = VarFlags & "bindable, "
   If (Flags And VARFLAG_FDEFAULTBIND) Then VarFlags = VarFlags & "defaultbind, "
   If (Flags And VARFLAG_FDEFAULTCOLLELEM) Then VarFlags = VarFlags & "defaultcollelem, "
   If (Flags And VARFLAG_FDISPLAYBIND) Then VarFlags = VarFlags & "displaybind, "
   If (Flags And VARFLAG_FHIDDEN) Then VarFlags = VarFlags & "hidden, "
   If (Flags And VARFLAG_FIMMEDIATEBIND) Then VarFlags = VarFlags & "immediatebind, "
   If (Flags And VARFLAG_FNONBROWSABLE) Then VarFlags = VarFlags & "nonbrowsable, "
   If (Flags And VARFLAG_FREQUESTEDIT) Then VarFlags = VarFlags & "requestedit, "
   If (Flags And VARFLAG_FRESTRICTED) Then VarFlags = VarFlags & "restricted, "
   If (Flags And VARFLAG_FSOURCE) Then VarFlags = VarFlags & "source, "
   If (Flags And VARFLAG_FUIDEFAULT) Then VarFlags = VarFlags & "uidefault, "
   
End Function


Private Function ResolveConstVal(ByVal TI As ITypeInfo, ByVal lIdx As Long, ByVal sName As String) As String
Dim lPtr As Long
Dim tVD As VARDESC
Dim vVar As Variant

   lPtr = TI.GetVarDesc(lIdx)
   
   MoveMemory tVD, ByVal lPtr, LenB(tVD)
   
   MoveMemory vVar, ByVal tVD.oInst_varValue, 16&
   
   If VarType(vVar) = vbString Then
      ResolveConstVal = sName & " = """ & vVar & """"
   Else
      ResolveConstVal = sName & " = 0x" & Hex$(vVar)
   End If
   
   MoveMemory vVar, ByVal 0&, 2&
   
   TI.ReleaseVarDesc lPtr
   
End Function

'
' DecompileEnum
'
' Decompiles an enum
'
' Parameters
'
' oTI - The enum ITypeInfo
'
Function DecompileEnum(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim lHlpCtx As Long
Dim sEnumName As String
Dim sHelp As String
Dim sHlpFile As String
Dim sIID As String
   
   ' Get TYPEATTR struct from ITypeInfo
   tTA = GetTypeAttr(oTI)
        
   ' Get the name, helpstring, helpfile and helpcontext
   oTI.GetDocumentation DISPID_UNKNOWN, sEnumName, sHelp, lHlpCtx, sHlpFile
   sEnumName = RTrimNull(sEnumName)

   DecompileEnum = "[" & vbCrLf
   
   ' Add the uuid
   If IsEqualGUID(IID_NULL, tTA.iid) = S_OK Then
      sIID = Space$(38)
      StringFromGUID2 tTA.iid, sIID, 39
      DecompileEnum = DecompileEnum & "    uuid(" & Mid$(sIID, 2, 36) & "), " & vbCrLf
   End If
   
   ' Add the helpstring, helpfile and helpcontext
   If LenB(sHelp) <> 0 Then DecompileEnum = DecompileEnum & "    helpstring(""" & RTrimNull(sHelp) & """), " & vbCrLf
   If LenB(sHlpFile) <> 0 Then DecompileEnum = DecompileEnum & "    helpfile(""" & RTrimNull(sHlpFile) & """), " & vbCrLf
   If lHlpCtx <> 0 Then DecompileEnum = DecompileEnum & "    helpcontext(0x" & Hex$(lHlpCtx) & "), " & vbCrLf
   
   ' Add the version
   If tTA.wMinorVerNum <> 0 Or tTA.wMajorVerNum <> 0 Then
      DecompileEnum = DecompileEnum & "    version(" & tTA.wMajorVerNum & "." & tTA.wMinorVerNum & "), " & vbCrLf
   End If
      
   If DecompileEnum = "[" & vbCrLf Then
      DecompileEnum = "typedef "
   Else
      DecompileEnum = "typedef " & vbCrLf & Left$(DecompileEnum, Len(DecompileEnum) - 4) & vbCrLf & "]" & vbCrLf
   End If
   
   DecompileEnum = DecompileEnum & "enum tag" & sEnumName & " {" & vbCrLf
   
   ' Check if the module/enum
   ' contains constants
   If tTA.cVars > 0 Then
      
      ' Decompile the constants
      DecompileEnum = DecompileEnum & DecompileConsts(oTI, tTA)
      
      DecompileEnum = DecompileEnum & "} " & sEnumName & ";" & vbCrLf
   
   End If

End Function
'
' DecompileCoclass
'
' Decompiles a coclass
'
' Parameters:
'
' oTI - The coclass ITypeInfo
'
Function DecompileCoclass(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim oInterfaceTI As ITypeInfo
Dim tInterfaceTA As TYPEATTR
Dim hRef As Long
Dim sName As String
Dim sCoclassName As String
Dim sHelp As String
Dim lHlpCtx As Long
Dim sHlpFile As String
Dim lIdx As Long
Dim sAttrs As String
Dim sIID As String
Dim lFlags As Long

   ' Get TYPEATTR from ITypeInfo
   tTA = GetTypeAttr(oTI)
        
   ' Get the coclass name, helpstring, helpfile and helpcontext
   oTI.GetDocumentation DISPID_UNKNOWN, sCoclassName, sHelp, lHlpCtx, sHlpFile
   sCoclassName = RTrimNull(sCoclassName)
   
   DecompileCoclass = "[" & vbCrLf
   
   ' Add the uuid
   If IsEqualGUID(IID_NULL, tTA.iid) = S_OK Then
      sIID = Space$(38)
      StringFromGUID2 tTA.iid, sIID, 39
      DecompileCoclass = DecompileCoclass & "    uuid(" & Mid$(sIID, 2, 36) & "), " & vbCrLf
   End If
   
   ' Add the helpfile, helpcontext and helpstring
   If LenB(sHelp) <> 0 Then
      DecompileCoclass = DecompileCoclass & "    helpstring(""" & RTrimNull(sHelp) & """), " & vbCrLf
   End If
   If LenB(sHlpFile) <> 0 Then
      DecompileCoclass = DecompileCoclass & "    helpfile(""" & RTrimNull(sHlpFile) & """), " & vbCrLf
   End If
   If lHlpCtx <> 0 Then
      DecompileCoclass = DecompileCoclass & "    helpcontext(0x" & Hex$(lHlpCtx) & "), " & vbCrLf
   End If
   
   ' Add the version
   If tTA.wMinorVerNum <> 0 Or tTA.wMajorVerNum <> 0 Then
      DecompileCoclass = DecompileCoclass & "    version(" & tTA.wMajorVerNum & "." & tTA.wMinorVerNum & "), " & vbCrLf
   End If
      
   ' Add the attributes to the coclass
   If DecompileCoclass = "[" & vbCrLf Then
      DecompileCoclass = vbNullString
   Else
      DecompileCoclass = Left$(DecompileCoclass, Len(DecompileCoclass) - 4) & vbCrLf & "]" & vbCrLf
   End If
   
   ' Add the name
   DecompileCoclass = DecompileCoclass & "coclass " & sCoclassName & " {" & vbCrLf
   
   ' Enumerate the implemented interfaces
   For lIdx = 0 To tTA.cImplTypes - 1
         
      ' Get the interface hRef
      hRef = oTI.GetRefTypeOfImplType(lIdx)
      
      On Error Resume Next
      
      Err.Clear
      
      ' Get the interface ITypeInfo
      Set oInterfaceTI = oTI.GetRefTypeInfo(hRef)
      
      If Err.Number = 0 Then
         ' Get the interface name
         oInterfaceTI.GetDocumentation DISPID_UNKNOWN, sName, vbNullString, 0, vbNullString
         sName = RTrimNull(sName)
      Else
         sName = "UNKNOWNTYPE"
      End If
      
      DecompileCoclass = DecompileCoclass & "    "
      
      ' Get the attributes
      lFlags = oTI.GetImplTypeFlags(lIdx)
      
      ' Add the attributes
      sAttrs = vbNullString
      If lFlags And IMPLTYPEFLAG_FDEFAULT Then sAttrs = sAttrs & "default, "
      If lFlags And IMPLTYPEFLAG_FSOURCE Then sAttrs = sAttrs & "source, "
      If lFlags And IMPLTYPEFLAG_FRESTRICTED Then sAttrs = sAttrs & "restricted, "
      If lFlags And IMPLTYPEFLAG_FDEFAULTVTABLE Then sAttrs = sAttrs & "defaultvtable, "
      If LenB(sAttrs) <> 0 Then DecompileCoclass = DecompileCoclass & "[" & Left$(sAttrs, Len(sAttrs) - 2) & "] "
      
      ' Get the TYPEATTR of the interface
      tInterfaceTA = GetTypeAttr(oInterfaceTI)
      
      ' Add the type kind
      Select Case tInterfaceTA.TYPEKIND
      
         Case TKIND_INTERFACE
            DecompileCoclass = DecompileCoclass & "interface "
            
         Case TKIND_DISPATCH
         
            ' Check if this dispinterface is a dual interface
            If tInterfaceTA.wTypeFlags And TYPEFLAG_FDUAL Then
               DecompileCoclass = DecompileCoclass & "interface "
            Else
               DecompileCoclass = DecompileCoclass & "dispinterface "
            End If
            
      End Select
      
      ' Add the interface name
      DecompileCoclass = DecompileCoclass & sName & ";" & vbCrLf
      
   Next
   
   ' Close the coclass
   DecompileCoclass = DecompileCoclass & "}; " & vbCrLf

End Function
'
' DecompileInterface
'
' Decompiles a interface or dispinterface
'
' Parameters:
'
' oTI - The interface/dispinterface ITypeInfo
'
Function DecompileInterface(ByVal oTI As ITypeInfo) As String
Dim oTI2 As ITypeInfo
Dim tTA As TYPEATTR
Dim hRef As Long
Dim sName As String
Dim sIName As String
Dim sHelp As String
Dim lHlpCtx As Long
Dim sHlpFile As String
Dim sIID As String

   ' Get TYPEATTR struct from ITypeInfo
   tTA = GetTypeAttr(oTI)
   
   ' If the type is TKIND_DISPATCH and has
   ' the dual attribute, get the interface
   If tTA.TYPEKIND = TKIND_DISPATCH And (tTA.wTypeFlags And TYPEFLAG_FDUAL) <> 0 Then
      Set oTI = oTI.GetRefTypeInfo(oTI.GetRefTypeOfImplType(-1))
      tTA = GetTypeAttr(oTI)
   End If

   ' Get the type documentation
   oTI.GetDocumentation DISPID_UNKNOWN, sIName, sHelp, lHlpCtx, sHlpFile
   sIName = RTrimNull(sIName)
   
   DecompileInterface = "[" & vbCrLf
   
   If tTA.TYPEKIND <> TKIND_DISPATCH Then DecompileInterface = DecompileInterface & "    odl, " & vbCrLf
   
   ' Add helpstring, helpfile and helpcontext attributes
   If LenB(sHelp) <> 0 Then DecompileInterface = DecompileInterface & "    helpstring(""" & RTrimNull(sHelp) & """), " & vbCrLf
   If LenB(sHlpFile) <> 0 Then DecompileInterface = DecompileInterface & "    helpfile(""" & RTrimNull(sHlpFile) & """), " & vbCrLf
   If lHlpCtx <> 0 Then DecompileInterface = DecompileInterface & "    helpcontext(0x" & Hex$(lHlpCtx) & "), " & vbCrLf
   
   ' Add uuid attribute
   If IsEqualGUID(IID_NULL, tTA.iid) = S_OK Then
      sIID = Space$(38)
      StringFromGUID2 tTA.iid, sIID, 39
      DecompileInterface = DecompileInterface & "    uuid(" & Mid$(sIID, 2, 36) & "), " & vbCrLf
   End If

   ' Add version attribute
   If tTA.wMinorVerNum <> 0 Or tTA.wMajorVerNum <> 0 Then
      DecompileInterface = DecompileInterface & "    version(" & tTA.wMajorVerNum & "." & tTA.wMinorVerNum & ")," & vbCrLf
   End If
   
   ' Add other attributes
   If tTA.wTypeFlags And TYPEFLAG_FDUAL Then DecompileInterface = DecompileInterface & "    dual, " & vbCrLf
   If tTA.wTypeFlags And TYPEFLAG_FOLEAUTOMATION Then DecompileInterface = DecompileInterface & "    oleautomation, " & vbCrLf
   If tTA.wTypeFlags And TYPEFLAG_FAGGREGATABLE Then DecompileInterface = DecompileInterface & "    aggregatable, " & vbCrLf
   If tTA.wTypeFlags And TYPEFLAG_FHIDDEN Then DecompileInterface = DecompileInterface & "    hidden, " & vbCrLf
   If tTA.wTypeFlags And TYPEFLAG_FNONEXTENSIBLE Then DecompileInterface = DecompileInterface & "    nonextensible, " & vbCrLf
   If tTA.wTypeFlags And TYPEFLAG_FRESTRICTED Then DecompileInterface = DecompileInterface & "    restricted, " & vbCrLf
   
   If DecompileInterface = "[" & vbCrLf Then
      ' No attributes, clear the return value
      DecompileInterface = vbNullString
   Else
      ' Add the attributes to the return value
      DecompileInterface = Left$(DecompileInterface, Len(DecompileInterface) - 4) & vbCrLf & "]" & vbCrLf
   End If
   
   ' Add the interface type
   If tTA.TYPEKIND = TKIND_DISPATCH Then
      DecompileInterface = DecompileInterface & "dispinterface " & sIName
      m_oTypes.Add "dispinterface " & sIName
   Else
      DecompileInterface = DecompileInterface & "interface " & sIName
      m_oTypes.Add "interface " & sIName
   End If
   
   ' Check if the interface inherits from other type
   If tTA.cImplTypes And tTA.TYPEKIND = TKIND_INTERFACE Then
   
      ' Get the parent ITypeInfo
      hRef = oTI.GetRefTypeOfImplType(0)
      Set oTI2 = oTI.GetRefTypeInfo(hRef)
      
      On Error Resume Next
      
      Dim oTL As ITypeLib, oTL2 As ITypeLib, sTL As String, sLib As String
      
      ' Get the typelibs
      oTI.GetContainingTypeLib oTL
      oTI2.GetContainingTypeLib oTL2
      
      If Not oTL Is oTL2 Then
         
         sTL = ResolveTypeLib(oTL)
         m_oImportedLibs.Add sTL, sTL
      
         ' Get the lib name
         oTL2.GetDocumentation DISPID_UNKNOWN, sLib, vbNullString, 0, vbNullString
         sLib = RTrimNull(sLib) & "."
      
      End If
      
      ' Get the type name
      oTI2.GetDocumentation DISPID_UNKNOWN, sName, vbNullString, 0, vbNullString
      sName = RTrimNull(sName)
      
      ' Add the parent name
      DecompileInterface = DecompileInterface & " : " & sLib & sName & " {" & vbCrLf
      
   Else
      DecompileInterface = DecompileInterface & " {" & vbCrLf
   End If
   
   ' If the type is TKIND_DISPATCH enumerate the properties
   If tTA.TYPEKIND = TKIND_DISPATCH Then
      DecompileInterface = DecompileInterface & DecompileProps(oTI)
      DecompileInterface = DecompileInterface & "methods:" & vbCrLf
   End If
   
   ' Enumerate methods
   DecompileInterface = DecompileInterface & DecompileFunctions(oTI)
   
   DecompileInterface = DecompileInterface & "}; " & vbCrLf
   
End Function
'
' DecompileModule
'
' Decompiles a module
'
' Parameters:
'
' oTI - ITypeInfo of the module
'
Function DecompileModule(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim hRef As Long
Dim lHlpCtx As Long
Dim sName As String
Dim sModuleName As String
Dim sHelp As String
Dim sHlpFile As String
Dim sIID As String

   ' Get TYPEATTR struct from ITypeInfo
   tTA = GetTypeAttr(oTI)
   
   ' If the type is TKIND_DISPATCH and has
   ' the dual attribute, get the interface
   If tTA.TYPEKIND = TKIND_DISPATCH And (tTA.wTypeFlags And TYPEFLAG_FDUAL) <> 0 Then
      Set oTI = oTI.GetRefTypeInfo(oTI.GetRefTypeOfImplType(-1))
      tTA = GetTypeAttr(oTI)
   End If

   ' Get the type documentation
   oTI.GetDocumentation DISPID_UNKNOWN, sModuleName, sHelp, lHlpCtx, sHlpFile
   sModuleName = RTrimNull(sModuleName)
   
   DecompileModule = "[" & vbCrLf
   
   ' Add uuid attribute
   If IsEqualGUID(IID_NULL, tTA.iid) = S_OK Then
      sIID = Space$(38)
      StringFromGUID2 tTA.iid, sIID, 39
      DecompileModule = DecompileModule & "    uuid(" & Mid$(sIID, 2, 36) & "), " & vbCrLf
   End If
   
   ' Add helpstring, helpfile and helpcontext attributes
   If LenB(sHelp) <> 0 Then DecompileModule = DecompileModule & "    helpstring(""" & RTrimNull(sHelp) & """), " & vbCrLf
   If LenB(sHlpFile) <> 0 Then DecompileModule = DecompileModule & "    helpfile(""" & RTrimNull(sHlpFile) & """), " & vbCrLf
   If lHlpCtx <> 0 Then DecompileModule = DecompileModule & "    helpcontext(0x" & Hex$(lHlpCtx) & "), " & vbCrLf
   
   ' Add version attribute
   If tTA.wMinorVerNum <> 0 Or tTA.wMajorVerNum <> 0 Then
      DecompileModule = DecompileModule & "    version(" & tTA.wMajorVerNum & "." & tTA.wMinorVerNum & ")," & vbCrLf
   End If
   
   ' Add other attributes
   If tTA.wTypeFlags And TYPEFLAG_FHIDDEN Then DecompileModule = DecompileModule & "    hidden, " & vbCrLf
   If tTA.wTypeFlags And TYPEFLAG_FRESTRICTED Then DecompileModule = DecompileModule & "    restricted, " & vbCrLf
   
   DecompileModule = DecompileModule & "    dllname(""" & GetDllName(oTI) & """), " & vbCrLf
   
   If DecompileModule = "[" & vbCrLf Then
      ' No attributes, clear the return value
      DecompileModule = vbNullString
   Else
      ' Add the attributes to the return value
      DecompileModule = Left$(DecompileModule, Len(DecompileModule) - 4) & vbCrLf & "]" & vbCrLf
   End If
   
   DecompileModule = DecompileModule & "module " & sModuleName & " {" & vbCrLf
   
   ' If the type is TKIND_DISPATCH enumerate the properties
   If tTA.TYPEKIND = TKIND_DISPATCH Then
      DecompileModule = DecompileModule & DecompileProps(oTI)
      DecompileModule = DecompileModule & "methods:" & vbCrLf
   End If
   
   ' Enumerate methods
   DecompileModule = DecompileModule & DecompileConsts(oTI, tTA)
   DecompileModule = DecompileModule & DecompileFunctions(oTI)
   
   DecompileModule = DecompileModule & "}; " & vbCrLf
   
End Function
'
' DecompileStruct
'
' Decompiles a struct
'
' Parameters:
'
' oTI - The struct ITypeInfo
'
Function DecompileStruct(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim tVD As VARDESC
Dim lPtr As Long
Dim lIdx As Long
Dim sName As String
Dim sRecName As String
Dim sHelp As String
Dim lHlpCtx As Long
Dim sHlpFile As String
Dim sAttrs As String
Dim sIID As String
   
   ' Get TYPEATTR struct from ITypeInfo
   tTA = GetTypeAttr(oTI)
        
   ' Get the name, helpstring, helpfile and helpcontext
   oTI.GetDocumentation DISPID_UNKNOWN, sRecName, sHelp, lHlpCtx, sHlpFile
   sRecName = RTrimNull(sRecName)

   DecompileStruct = "[" & vbCrLf
   
   ' Add the uuid
   If IsEqualGUID(IID_NULL, tTA.iid) = S_OK Then
      sIID = Space$(38)
      StringFromGUID2 tTA.iid, sIID, 39
      DecompileStruct = DecompileStruct & "    uuid(" & Mid$(sIID, 2, 36) & "), " & vbCrLf
   End If
   
   ' Add the helpstring, helpfile and helpcontext
   If LenB(sHelp) <> 0 Then DecompileStruct = DecompileStruct & "    helpstring(""" & RTrimNull(sHelp) & """), " & vbCrLf
   If LenB(sHlpFile) <> 0 Then DecompileStruct = DecompileStruct & "    helpfile(""" & RTrimNull(sHlpFile) & """), " & vbCrLf
   If lHlpCtx <> 0 Then DecompileStruct = DecompileStruct & "    helpcontext(0x" & Hex$(lHlpCtx) & "), " & vbCrLf
     
   ' Add the version
   If tTA.wMinorVerNum <> 0 Or tTA.wMajorVerNum <> 0 Then
      DecompileStruct = DecompileStruct & "    version(" & tTA.wMajorVerNum & "." & tTA.wMinorVerNum & "), " & vbCrLf
   End If
   
   If tTA.wTypeFlags And FUNCFLAG_FHIDDEN Then DecompileStruct = DecompileStruct & "    hidden, " & vbCrLf
   
   If DecompileStruct = "[" & vbCrLf Then
      DecompileStruct = "typedef "
   Else
      DecompileStruct = "typedef " & vbCrLf & Left$(DecompileStruct, Len(DecompileStruct) - 4) & vbCrLf & "]" & vbCrLf
   End If
   
   DecompileStruct = "// record size: " & tTA.cbSizeInstance & " bytes" & vbCrLf & DecompileStruct

   DecompileStruct = DecompileStruct & "struct tag" & sRecName & " {" & vbCrLf
   
   For lIdx = 0 To tTA.cVars - 1
         
      ' Get a pointer to the VARDESC
      ' struct for this constant
      lPtr = oTI.GetVarDesc(lIdx)
         
      ' Copy from the pointer
      MoveMemory tVD, ByVal lPtr, LenB(tVD)
         
      ' Get the name
      oTI.GetDocumentation tVD.memid, sName, sHelp, lHlpCtx, sHlpFile
      sName = RTrimNull(sName)
      
      sAttrs = vbNullString
      
      If LenB(sHelp) <> 0 Then sAttrs = sAttrs & "helpstring(""" & RTrimNull(sHelp) & """), "
      If lHlpCtx <> 0 Then sAttrs = sAttrs & "helpcontext(0x" & Hex$(lHlpCtx) & "), "
      
      If LenB(sAttrs) <> 0 Then DecompileStruct = DecompileStruct & "    [" & Left$(sAttrs, Len(sAttrs) - 2) & "]" & vbCrLf
      DecompileStruct = DecompileStruct & "    " & ResolveVarType(oTI, tVD.elemdescVar.tdesc, sName) & ";" & vbCrLf
      
      ' Release the pointer
      oTI.ReleaseFuncDesc lPtr
         
   Next
   
   DecompileStruct = DecompileStruct & "} " & sRecName & ";" & vbCrLf

End Function
'
' DecompileUnion
'
' Decompiles an union
'
Function DecompileUnion(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim lPtr As Long
Dim tVD As VARDESC
Dim sName As String
Dim sRecName As String
Dim sHelp As String
Dim lHlpCtx As Long
Dim sHlpFile As String
Dim lIdx As Long
Dim sAttrs As String
Dim sIID As String
   
   ' Get TYPEATTR struct from ITypeInfo
   tTA = GetTypeAttr(oTI)
        
   On Error Resume Next

   oTI.GetDocumentation DISPID_UNKNOWN, sRecName, sHelp, lHlpCtx, sHlpFile
   
   DecompileUnion = "[" & vbCrLf
   
   If LenB(sHelp) <> 0 Then DecompileUnion = DecompileUnion & "    helpstring(""" & RTrimNull(sHelp) & """), " & vbCrLf
   If LenB(sHlpFile) <> 0 Then DecompileUnion = DecompileUnion & "    helpfile(""" & RTrimNull(sHlpFile) & """), " & vbCrLf
   If lHlpCtx <> 0 Then DecompileUnion = DecompileUnion & "    helpcontext(0x" & Hex$(lHlpCtx) & "), " & vbCrLf
   
   If IsEqualGUID(IID_NULL, tTA.iid) = S_OK Then
      sIID = Space$(38)
      StringFromGUID2 tTA.iid, sIID, 39
      DecompileUnion = DecompileUnion & "    uuid(" & Mid$(sIID, 2, 36) & "), " & vbCrLf
   End If
  
   If tTA.wMinorVerNum <> 0 Or tTA.wMajorVerNum <> 0 Then
      DecompileUnion = DecompileUnion & "    version(" & tTA.wMajorVerNum & "." & tTA.wMinorVerNum & "), " & vbCrLf
   End If
   
   If tTA.wTypeFlags And FUNCFLAG_FHIDDEN Then DecompileUnion = DecompileUnion & "    hidden, " & vbCrLf
   
   If DecompileUnion = "[" & vbCrLf Then
      DecompileUnion = "typedef "
   Else
      DecompileUnion = "typedef " & vbCrLf & Left$(DecompileUnion, Len(DecompileUnion) - 4) & vbCrLf & "]" & vbCrLf
   End If
   
   DecompileUnion = "// union size: " & tTA.cbSizeInstance & " bytes" & vbCrLf & DecompileUnion
   DecompileUnion = DecompileUnion & "union tag" & sRecName & " {" & vbCrLf
      
   ' Check if the module/enum
   ' contains constants
   If tTA.cVars > 0 Then
      
      For lIdx = 0 To tTA.cVars - 1
            
         ' Get a pointer to the VARDESC
         ' struct for this constant
         lPtr = oTI.GetVarDesc(lIdx)
            
         ' Copy from the pointer
         MoveMemory tVD, ByVal lPtr, LenB(tVD)
            
         ' Get the name
         oTI.GetDocumentation tVD.memid, sName, sHelp, lHlpCtx, sHlpFile
         
         sAttrs = vbNullString
         If LenB(sHelp) <> 0 Then sAttrs = sAttrs & "helpstring(""" & RTrimNull(sHelp) & """), "
         If LenB(sHlpFile) <> 0 Then sAttrs = sAttrs & "helpfile(""" & RTrimNull(sHlpFile) & """), "
         If lHlpCtx <> 0 Then sAttrs = sAttrs & "helpcontext(0x" & Hex$(lHlpCtx) & "), "
         
         If LenB(sAttrs) <> 0 Then DecompileUnion = DecompileUnion & "    [" & Left$(sAttrs, Len(sAttrs) - 2) & "]" & vbCrLf
         DecompileUnion = DecompileUnion & "    " & ResolveVarType(oTI, tVD.elemdescVar.tdesc, sName) & "," & vbCrLf
         
         ' Release the pointer
         oTI.ReleaseFuncDesc lPtr
            
      Next
      
      DecompileUnion = DecompileUnion & "} " & sRecName & ";" & vbCrLf
   
   End If

End Function




'
' DecompileTypedef
'
' Decompiles a typedef
'
' Parameters:
'
' oTI - The typedef ITypeInfo
'
Function DecompileTypedef(ByVal oTI As ITypeInfo) As String
Dim tTA As TYPEATTR
Dim sName As String
Dim sHelp As String
Dim lHlpCtx As Long
Dim sHlpFile As String
Dim sIID As String
   
   ' Get TYPEATTR struct from ITypeInfo
   tTA = GetTypeAttr(oTI)
        
   ' Get the name, helpstring, helpfile and helpcontext
   oTI.GetDocumentation DISPID_UNKNOWN, sName, sHelp, lHlpCtx, sHlpFile
   
   DecompileTypedef = "typedef [public, "
   
   ' Add the uuid
   If IsEqualGUID(IID_NULL, tTA.iid) = S_OK Then
      sIID = Space$(38)
      StringFromGUID2 tTA.iid, sIID, 39
      DecompileTypedef = DecompileTypedef & "uuid(" & Mid$(sIID, 2, 36) & "), "
   End If
   
   ' Add the attributes
   If LenB(sHelp) <> 0 Then DecompileTypedef = DecompileTypedef & "    helpstring(""" & RTrimNull(sHelp) & """), "
   If LenB(sHlpFile) <> 0 Then DecompileTypedef = DecompileTypedef & "    helpfile(""" & RTrimNull(sHlpFile) & """), "
   If lHlpCtx <> 0 Then DecompileTypedef = DecompileTypedef & "    helpcontext(0x" & Hex$(lHlpCtx) & "), "
   
   DecompileTypedef = Left$(DecompileTypedef, Len(DecompileTypedef) - 2) & "] "
   
   ' Resolve the type
   DecompileTypedef = DecompileTypedef & ResolveVarType(oTI, tTA.tdescAlias, sName)
   
   DecompileTypedef = DecompileTypedef & ";" & vbCrLf

End Function
'
' GetTypeAttr
'
' Returns the TYPEATTR of a ITypeInfo
'
Function GetTypeAttr(ByVal oTInfo As ITypeInfo) As TYPEATTR
Dim TAPtr As Long

   ' Get TYPEATTR pointer
   TAPtr = oTInfo.GetTypeAttr
   
   ' Copy the struct from the pointer
   MoveMemory GetTypeAttr, ByVal TAPtr, Len(GetTypeAttr)
   
   ' Release the pointer
   oTInfo.ReleaseTypeAttr TAPtr
   
End Function


Private Function ResolveVarType(TI As ITypeInfo, TD As TYPEDESC, ByVal VarName As String) As String
Dim tTD As TYPEDESC
Dim oTI As ITypeInfo
Dim tTA As TYPEATTR
Dim lPtr As Long
Dim llIdx As Long
Dim sName As String

   If TD.vt And VT_BYREF Then
      ResolveVarType = "* " & VarName
      TD.vt = TD.vt And Not VT_BYREF
   Else
      ResolveVarType = VarName
   End If

   Select Case TD.vt
   
      Case VT_ARRAY
      
      Case VT_I2
         ResolveVarType = "short " & ResolveVarType
      
      Case VT_I4
         ResolveVarType = "long " & ResolveVarType
      
      Case VT_INT
         ResolveVarType = "int " & ResolveVarType
      
      Case VT_R4
         ResolveVarType = "float " & ResolveVarType
      
      Case VT_R8
         ResolveVarType = "double " & ResolveVarType
      
      Case VT_CY
         ResolveVarType = "CURRENCY " & ResolveVarType
      
      Case VT_DATE
         ResolveVarType = "double " & ResolveVarType
      
      Case VT_BSTR
         ResolveVarType = "BSTR " & ResolveVarType
      
      Case VT_DISPATCH
         ResolveVarType = "IDispatch *" & ResolveVarType
         
         On Error Resume Next
         m_oImportedLibs.Add "stdole2.tlb", "stdole2.tlb"
      
      Case VT_BOOL
         ResolveVarType = "BOOLEAN " & ResolveVarType
      
      Case VT_VARIANT
         ResolveVarType = "VARIANT " & ResolveVarType
      
      Case VT_UNKNOWN
         ResolveVarType = "IUnknown *" & ResolveVarType
         
         On Error Resume Next
         m_oImportedLibs.Add "stdole2.tlb", "stdole2.tlb"
      
      Case VT_I1
         ResolveVarType = "char " & ResolveVarType
      
      Case VT_UI1
         ResolveVarType = "unsigned char " & ResolveVarType
      
      Case VT_UI2
         ResolveVarType = "unsigned short " & ResolveVarType
      
      Case VT_UI4
         ResolveVarType = "unsigned long " & ResolveVarType
      
      Case VT_UINT
         ResolveVarType = "unsigned int " & ResolveVarType
      
      Case VT_I8
         ResolveVarType = "LONGLONG " & ResolveVarType
      
      Case VT_UI8
         ResolveVarType = "ULONGLONG " & ResolveVarType
      
      Case VT_VOID
         ResolveVarType = "void " & ResolveVarType
      
      Case VT_HRESULT
         ResolveVarType = "HRESULT " & ResolveVarType
      
      Case VT_LPSTR
         ResolveVarType = "LPSTR " & ResolveVarType
      
      Case VT_LPWSTR
         ResolveVarType = "LPWSTR " & ResolveVarType
      
      Case VT_FILETIME
         ResolveVarType = "FILETIME " & ResolveVarType
      
      Case VT_BLOB
         ResolveVarType = "BLOB " & ResolveVarType
      
      Case VT_STREAM, VT_STREAMED_OBJECT
         ResolveVarType = "IStream *" & ResolveVarType
      
      Case VT_STORAGE, VT_STORED_OBJECT
         ResolveVarType = "IStorage *" & ResolveVarType
      
      Case VT_CLSID
         ResolveVarType = "GUID " & ResolveVarType
      
      Case VT_USERDEFINED
      
         Dim TL1 As ITypeLib, TL2 As ITypeLib
         
         On Error Resume Next
         
         Set oTI = TI.GetRefTypeInfo(TD.pTypeDesc)
         
         If Err.Number = 0 Then
            lPtr = oTI.GetTypeAttr
            
            oTI.GetContainingTypeLib TL1
            TI.GetContainingTypeLib TL2
            If Not TL2 Is TL1 Then
               On Error Resume Next
               sName = ResolveTypeLib(TL1)
               m_oImportedLibs.Add sName, sName
            End If
            
            MoveMemory tTA, ByVal lPtr, LenB(tTA)
                     
            oTI.GetDocumentation DISPID_UNKNOWN, sName, "", 0, ""
            
            ResolveVarType = sName & " " & VarName
            
            oTI.ReleaseTypeAttr lPtr
         Else
            ResolveVarType = "UNKNOWNTYPE " & VarName
         End If
         
      Case VT_PTR
         
         MoveMemory tTD, ByVal TD.pTypeDesc, LenB(tTD)
         ResolveVarType = ResolveVarType(TI, tTD, "") & "*" & VarName
      
      Case VT_SAFEARRAY
      
         MoveMemory tTD, ByVal TD.pTypeDesc, LenB(tTD)
         ResolveVarType = "SAFEARRAY(" & ResolveVarType(TI, tTD, "") & ") " & VarName
      
      Case VT_CARRAY
      
         Dim tAD As ARRAYDESC, tSAB As SAFEARRAYBOUND
         
         MoveMemory tAD, ByVal TD.pTypeDesc, LenB(tAD)
         
         ResolveVarType = ResolveVarType(TI, tAD.tdescElem, "") & VarName & " ["
         
         For llIdx = 0 To tAD.cDims - 1
            MoveMemory tSAB, ByVal TD.pTypeDesc + 12 + LenB(tSAB) * llIdx, Len(tSAB)
            ResolveVarType = ResolveVarType & tSAB.cElements
            If llIdx > tAD.cDims - 1 Then ResolveVarType = ResolveVarType & ", "
         Next
         
         ResolveVarType = ResolveVarType & "]"
      
      Case VT_ERROR
         ResolveVarType = "SCODE " & ResolveVarType
            
   End Select
   
End Function



