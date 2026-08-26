Attribute VB_Name = "mdlDispInvoke"
'*********************************************************************************************
'
' IDispatch Invoke Sample
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 02/28/2000
'
'*********************************************************************************************
Option Explicit

Enum vbInvoke
   vbGet = INVOKE_PROPERTYGET
   vbLet = INVOKE_PROPERTYPUT
   vbSet = INVOKE_PROPERTYPUTREF
End Enum
'*********************************************************************************************
'
' Invoke
' ======
'
' Invokes a method or property of an object given its name or lDISPID.
' This function does exactly the same as the VB6 CallByName function.
'
' Parameters:
'
' Obj                : IDispatch interface of target object.
' MemberNameOrDISPID : Name of the property/method or DISPID
' CallType           : Type of the call. (Property Let, Get, Set or method/function)
' Args               : List of arguments passed to the property/method
'
'*********************************************************************************************
Public Function Invoke( _
      oDispatch As olelib.IDispatch, _
      ByVal MemberNameOrDISPID As Variant, _
      ByVal CallType As olelib.INVOKEKIND, _
      ParamArray Args() As Variant) As Variant

Dim IID_NULL As olelib.UUID           ' NULL interface ID
Dim lDISPID As Long                   ' Dispatch ID
Dim tDISPPARAMS As olelib.DISPPARAMS ' Parameters UDT
Dim tEXCEPINFO As olelib.EXCEPINFO   ' Exception Error info
Dim avParams() As Variant             ' Parameters array
Dim avNamedParams() As Long           ' Named parameters array
Dim lArgErr As Long                   ' Argument that produced the error
Dim vResult As Variant                ' vResult value
Dim lResult As Long                   ' Invoke return value
Dim lIdx As Long, lMax As Long        ' Index and number of parameters
    
   ' If MemberNameOrDISPID is numeric
   ' it represents the lDISPID of the
   ' member otherwise get the lDISPID
   ' of the given member name
   If IsNumeric(MemberNameOrDISPID) Then
      lDISPID = MemberNameOrDISPID
   Else
      oDispatch.GetIDsOfNames IID_NULL, CStr(MemberNameOrDISPID), 1, 0, lDISPID
   End If
        
   If Not IsMissing(Args) Then
   
      ' Get parameters count
      lMax = UBound(Args)
         
      ' Redim the parameters array
      ReDim avParams(0 To lMax)
               
      ' Fill parameters arrays. The array has
      ' to be filled in reseversed order
      ' so the first element is the last
      ' parameter.
      For lIdx = 0 To lMax
      
         If IsObject(Args(lIdx)) Then
            Set avParams(lMax - lIdx) = Args(lIdx)
         Else
            avParams(lMax - lIdx) = Args(lIdx)
         End If
         
      Next
               
      ' Fill the tDISPPARAMS structure
      With tDISPPARAMS
         .cArgs = UBound(avParams) + 1
         .rgPointerToVariantArray = VarPtr(avParams(0))
      End With
         
      ' If the procedure is a
      ' Property Put or Property Set
      ' the last parameter represents
      ' the new property value and has
      ' to be set as DISPID_PROPERTYPUT
      ' named argument.
      If CallType = INVOKE_PROPERTYPUT Or CallType = INVOKE_PROPERTYPUTREF Then
            
         ReDim avNamedParams(0 To 0)
         
         avNamedParams(0) = DISPID_PROPERTYPUT
         
         With tDISPPARAMS
            .cNamedArgs = 1
            .rgPointerToLONGNamedArgs = VarPtr(avNamedParams(0))
         End With
         
      End If
   
   End If
   
   ' Invoke method/property
   lResult = oDispatch.Invoke(lDISPID, IID_NULL, 0, CallType, tDISPPARAMS, VarPtr(vResult), tEXCEPINFO, lArgErr)
   
   If lResult <> 0 Then
   
      ' There was an error
        
      If lResult = DISP_E_EXCEPTION Then
                
         With tEXCEPINFO
            Err.Raise .wCode, .Source, .Description, .HelpFile, .dwHelpContext
         End With
            
      Else
                
         Err.Raise lResult
            
      End If
                
   End If
   
   ' Property Set and Property Let
   ' doesn't return values
   If CallType = INVOKE_FUNC Or CallType = INVOKE_PROPERTYGET Then
      
      If IsObject(vResult) Then
       
         Set Invoke = vResult
         
      Else
      
         Invoke = vResult
         
      End If
      
   End If
    
End Function
