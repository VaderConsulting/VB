/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Dec 12 18:18:19 2002
 */
/* Compiler settings for D:\cpp\Com\SchedTask\SchedTask.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __SchedTask_h__
#define __SchedTask_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IetScheduledTask_FWD_DEFINED__
#define __IetScheduledTask_FWD_DEFINED__
typedef interface IetScheduledTask IetScheduledTask;
#endif 	/* __IetScheduledTask_FWD_DEFINED__ */


#ifndef __etScheduledTask_FWD_DEFINED__
#define __etScheduledTask_FWD_DEFINED__

#ifdef __cplusplus
typedef class etScheduledTask etScheduledTask;
#else
typedef struct etScheduledTask etScheduledTask;
#endif /* __cplusplus */

#endif 	/* __etScheduledTask_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __IetScheduledTask_INTERFACE_DEFINED__
#define __IetScheduledTask_INTERFACE_DEFINED__

/* interface IetScheduledTask */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IetScheduledTask;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("656ED66F-20F1-11D4-A450-0004AC9B6CD1")
    IetScheduledTask : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Program( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Program( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Parameters( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Parameters( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartingDir( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartingDir( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Comment( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Comment( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartYear( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartYear( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartMonth( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartMonth( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartDay( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartDay( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartHour( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartHour( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartMinute( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartMinute( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TriggerType( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TriggerType( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CreateTask( 
            BSTR bstrTaskName,
            BOOL bFailifExists) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddTrigger( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RemoveTask( 
            BSTR bstrTaskName) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Reset( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EndYear( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EndYear( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EndMonth( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EndMonth( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EndDay( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EndDay( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_MinutesInterval( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_MinutesInterval( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TargetComputer( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TargetComputer( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_MinutesDuration( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_MinutesDuration( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DayOfTheWeek( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DayOfTheWeek( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DayOfTheMonth( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DayOfTheMonth( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Enum( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TaskCount( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TaskItem( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TaskName( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DontStartIfOnBatteries( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DontStartIfOnBatteries( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_KillIfOnBatteries( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_KillIfOnBatteries( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Username( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Username( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Password( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Password( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Creator( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LocalAccount( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_LocalAccount( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE EditTask( 
            BSTR bstrTaskName) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_WhichWeek( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_WhichWeek( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddMonth( 
            short nMonth) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ActivateTask( 
            BSTR bstrTaskName) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ExitCode( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DeleteWhenDone( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DeleteWhenDone( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RunOnlyIfLogged( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RunOnlyIfLogged( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartOnlyIfIdle( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartOnlyIfIdle( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RunIfConnectedInternet( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RunIfConnectedInternet( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetTrigger( 
            short nIndex) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RestartOnIdleResume( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RestartOnIdleResume( 
            /* [in] */ BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Run( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Terminate( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetStatus( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DaysInterval( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DaysInterval( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_WeeksInterval( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_WeeksInterval( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LastTimeRun( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NextTimeRun( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RemoveTrigger( 
            short nIndex) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TriggerString( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Disabled( 
            /* [retval][out] */ BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Disabled( 
            /* [in] */ BOOL newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IetScheduledTaskVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IetScheduledTask __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IetScheduledTask __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IetScheduledTask __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Program )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Program )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Parameters )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Parameters )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartingDir )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StartingDir )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Comment )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Comment )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartYear )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StartYear )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartMonth )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StartMonth )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartDay )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StartDay )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartHour )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StartHour )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartMinute )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StartMinute )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TriggerType )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TriggerType )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CreateTask )( 
            IetScheduledTask __RPC_FAR * This,
            BSTR bstrTaskName,
            BOOL bFailifExists);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddTrigger )( 
            IetScheduledTask __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RemoveTask )( 
            IetScheduledTask __RPC_FAR * This,
            BSTR bstrTaskName);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Reset )( 
            IetScheduledTask __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EndYear )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EndYear )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EndMonth )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EndMonth )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EndDay )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EndDay )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_MinutesInterval )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_MinutesInterval )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TargetComputer )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TargetComputer )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_MinutesDuration )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_MinutesDuration )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DayOfTheWeek )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DayOfTheWeek )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DayOfTheMonth )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DayOfTheMonth )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Enum )( 
            IetScheduledTask __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TaskCount )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TaskItem )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TaskName )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DontStartIfOnBatteries )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DontStartIfOnBatteries )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_KillIfOnBatteries )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_KillIfOnBatteries )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Username )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Username )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Password )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Password )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Creator )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Creator )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LocalAccount )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_LocalAccount )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EditTask )( 
            IetScheduledTask __RPC_FAR * This,
            BSTR bstrTaskName);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_WhichWeek )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_WhichWeek )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddMonth )( 
            IetScheduledTask __RPC_FAR * This,
            short nMonth);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ActivateTask )( 
            IetScheduledTask __RPC_FAR * This,
            BSTR bstrTaskName);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ExitCode )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DeleteWhenDone )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DeleteWhenDone )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RunOnlyIfLogged )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RunOnlyIfLogged )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartOnlyIfIdle )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StartOnlyIfIdle )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RunIfConnectedInternet )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RunIfConnectedInternet )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTrigger )( 
            IetScheduledTask __RPC_FAR * This,
            short nIndex);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RestartOnIdleResume )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RestartOnIdleResume )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Run )( 
            IetScheduledTask __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Terminate )( 
            IetScheduledTask __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetStatus )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DaysInterval )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DaysInterval )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_WeeksInterval )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_WeeksInterval )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LastTimeRun )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NextTimeRun )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RemoveTrigger )( 
            IetScheduledTask __RPC_FAR * This,
            short nIndex);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TriggerString )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Disabled )( 
            IetScheduledTask __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Disabled )( 
            IetScheduledTask __RPC_FAR * This,
            /* [in] */ BOOL newVal);
        
        END_INTERFACE
    } IetScheduledTaskVtbl;

    interface IetScheduledTask
    {
        CONST_VTBL struct IetScheduledTaskVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IetScheduledTask_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IetScheduledTask_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IetScheduledTask_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IetScheduledTask_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IetScheduledTask_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IetScheduledTask_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IetScheduledTask_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IetScheduledTask_get_Program(This,pVal)	\
    (This)->lpVtbl -> get_Program(This,pVal)

#define IetScheduledTask_put_Program(This,newVal)	\
    (This)->lpVtbl -> put_Program(This,newVal)

#define IetScheduledTask_get_Parameters(This,pVal)	\
    (This)->lpVtbl -> get_Parameters(This,pVal)

#define IetScheduledTask_put_Parameters(This,newVal)	\
    (This)->lpVtbl -> put_Parameters(This,newVal)

#define IetScheduledTask_get_StartingDir(This,pVal)	\
    (This)->lpVtbl -> get_StartingDir(This,pVal)

#define IetScheduledTask_put_StartingDir(This,newVal)	\
    (This)->lpVtbl -> put_StartingDir(This,newVal)

#define IetScheduledTask_get_Comment(This,pVal)	\
    (This)->lpVtbl -> get_Comment(This,pVal)

#define IetScheduledTask_put_Comment(This,newVal)	\
    (This)->lpVtbl -> put_Comment(This,newVal)

#define IetScheduledTask_get_StartYear(This,pVal)	\
    (This)->lpVtbl -> get_StartYear(This,pVal)

#define IetScheduledTask_put_StartYear(This,newVal)	\
    (This)->lpVtbl -> put_StartYear(This,newVal)

#define IetScheduledTask_get_StartMonth(This,pVal)	\
    (This)->lpVtbl -> get_StartMonth(This,pVal)

#define IetScheduledTask_put_StartMonth(This,newVal)	\
    (This)->lpVtbl -> put_StartMonth(This,newVal)

#define IetScheduledTask_get_StartDay(This,pVal)	\
    (This)->lpVtbl -> get_StartDay(This,pVal)

#define IetScheduledTask_put_StartDay(This,newVal)	\
    (This)->lpVtbl -> put_StartDay(This,newVal)

#define IetScheduledTask_get_StartHour(This,pVal)	\
    (This)->lpVtbl -> get_StartHour(This,pVal)

#define IetScheduledTask_put_StartHour(This,newVal)	\
    (This)->lpVtbl -> put_StartHour(This,newVal)

#define IetScheduledTask_get_StartMinute(This,pVal)	\
    (This)->lpVtbl -> get_StartMinute(This,pVal)

#define IetScheduledTask_put_StartMinute(This,newVal)	\
    (This)->lpVtbl -> put_StartMinute(This,newVal)

#define IetScheduledTask_get_TriggerType(This,pVal)	\
    (This)->lpVtbl -> get_TriggerType(This,pVal)

#define IetScheduledTask_put_TriggerType(This,newVal)	\
    (This)->lpVtbl -> put_TriggerType(This,newVal)

#define IetScheduledTask_CreateTask(This,bstrTaskName,bFailifExists)	\
    (This)->lpVtbl -> CreateTask(This,bstrTaskName,bFailifExists)

#define IetScheduledTask_AddTrigger(This)	\
    (This)->lpVtbl -> AddTrigger(This)

#define IetScheduledTask_RemoveTask(This,bstrTaskName)	\
    (This)->lpVtbl -> RemoveTask(This,bstrTaskName)

#define IetScheduledTask_Reset(This)	\
    (This)->lpVtbl -> Reset(This)

#define IetScheduledTask_get_EndYear(This,pVal)	\
    (This)->lpVtbl -> get_EndYear(This,pVal)

#define IetScheduledTask_put_EndYear(This,newVal)	\
    (This)->lpVtbl -> put_EndYear(This,newVal)

#define IetScheduledTask_get_EndMonth(This,pVal)	\
    (This)->lpVtbl -> get_EndMonth(This,pVal)

#define IetScheduledTask_put_EndMonth(This,newVal)	\
    (This)->lpVtbl -> put_EndMonth(This,newVal)

#define IetScheduledTask_get_EndDay(This,pVal)	\
    (This)->lpVtbl -> get_EndDay(This,pVal)

#define IetScheduledTask_put_EndDay(This,newVal)	\
    (This)->lpVtbl -> put_EndDay(This,newVal)

#define IetScheduledTask_get_MinutesInterval(This,pVal)	\
    (This)->lpVtbl -> get_MinutesInterval(This,pVal)

#define IetScheduledTask_put_MinutesInterval(This,newVal)	\
    (This)->lpVtbl -> put_MinutesInterval(This,newVal)

#define IetScheduledTask_get_TargetComputer(This,pVal)	\
    (This)->lpVtbl -> get_TargetComputer(This,pVal)

#define IetScheduledTask_put_TargetComputer(This,newVal)	\
    (This)->lpVtbl -> put_TargetComputer(This,newVal)

#define IetScheduledTask_get_MinutesDuration(This,pVal)	\
    (This)->lpVtbl -> get_MinutesDuration(This,pVal)

#define IetScheduledTask_put_MinutesDuration(This,newVal)	\
    (This)->lpVtbl -> put_MinutesDuration(This,newVal)

#define IetScheduledTask_get_DayOfTheWeek(This,pVal)	\
    (This)->lpVtbl -> get_DayOfTheWeek(This,pVal)

#define IetScheduledTask_put_DayOfTheWeek(This,newVal)	\
    (This)->lpVtbl -> put_DayOfTheWeek(This,newVal)

#define IetScheduledTask_get_DayOfTheMonth(This,pVal)	\
    (This)->lpVtbl -> get_DayOfTheMonth(This,pVal)

#define IetScheduledTask_put_DayOfTheMonth(This,newVal)	\
    (This)->lpVtbl -> put_DayOfTheMonth(This,newVal)

#define IetScheduledTask_Enum(This)	\
    (This)->lpVtbl -> Enum(This)

#define IetScheduledTask_get_TaskCount(This,pVal)	\
    (This)->lpVtbl -> get_TaskCount(This,pVal)

#define IetScheduledTask_put_TaskItem(This,newVal)	\
    (This)->lpVtbl -> put_TaskItem(This,newVal)

#define IetScheduledTask_get_TaskName(This,pVal)	\
    (This)->lpVtbl -> get_TaskName(This,pVal)

#define IetScheduledTask_get_DontStartIfOnBatteries(This,pVal)	\
    (This)->lpVtbl -> get_DontStartIfOnBatteries(This,pVal)

#define IetScheduledTask_put_DontStartIfOnBatteries(This,newVal)	\
    (This)->lpVtbl -> put_DontStartIfOnBatteries(This,newVal)

#define IetScheduledTask_get_KillIfOnBatteries(This,pVal)	\
    (This)->lpVtbl -> get_KillIfOnBatteries(This,pVal)

#define IetScheduledTask_put_KillIfOnBatteries(This,newVal)	\
    (This)->lpVtbl -> put_KillIfOnBatteries(This,newVal)

#define IetScheduledTask_get_Username(This,pVal)	\
    (This)->lpVtbl -> get_Username(This,pVal)

#define IetScheduledTask_put_Username(This,newVal)	\
    (This)->lpVtbl -> put_Username(This,newVal)

#define IetScheduledTask_get_Password(This,pVal)	\
    (This)->lpVtbl -> get_Password(This,pVal)

#define IetScheduledTask_put_Password(This,newVal)	\
    (This)->lpVtbl -> put_Password(This,newVal)

#define IetScheduledTask_get_Creator(This,pVal)	\
    (This)->lpVtbl -> get_Creator(This,pVal)

#define IetScheduledTask_put_Creator(This,newVal)	\
    (This)->lpVtbl -> put_Creator(This,newVal)

#define IetScheduledTask_get_LocalAccount(This,pVal)	\
    (This)->lpVtbl -> get_LocalAccount(This,pVal)

#define IetScheduledTask_put_LocalAccount(This,newVal)	\
    (This)->lpVtbl -> put_LocalAccount(This,newVal)

#define IetScheduledTask_EditTask(This,bstrTaskName)	\
    (This)->lpVtbl -> EditTask(This,bstrTaskName)

#define IetScheduledTask_get_WhichWeek(This,pVal)	\
    (This)->lpVtbl -> get_WhichWeek(This,pVal)

#define IetScheduledTask_put_WhichWeek(This,newVal)	\
    (This)->lpVtbl -> put_WhichWeek(This,newVal)

#define IetScheduledTask_AddMonth(This,nMonth)	\
    (This)->lpVtbl -> AddMonth(This,nMonth)

#define IetScheduledTask_ActivateTask(This,bstrTaskName)	\
    (This)->lpVtbl -> ActivateTask(This,bstrTaskName)

#define IetScheduledTask_get_ExitCode(This,pVal)	\
    (This)->lpVtbl -> get_ExitCode(This,pVal)

#define IetScheduledTask_get_DeleteWhenDone(This,pVal)	\
    (This)->lpVtbl -> get_DeleteWhenDone(This,pVal)

#define IetScheduledTask_put_DeleteWhenDone(This,newVal)	\
    (This)->lpVtbl -> put_DeleteWhenDone(This,newVal)

#define IetScheduledTask_get_RunOnlyIfLogged(This,pVal)	\
    (This)->lpVtbl -> get_RunOnlyIfLogged(This,pVal)

#define IetScheduledTask_put_RunOnlyIfLogged(This,newVal)	\
    (This)->lpVtbl -> put_RunOnlyIfLogged(This,newVal)

#define IetScheduledTask_get_StartOnlyIfIdle(This,pVal)	\
    (This)->lpVtbl -> get_StartOnlyIfIdle(This,pVal)

#define IetScheduledTask_put_StartOnlyIfIdle(This,newVal)	\
    (This)->lpVtbl -> put_StartOnlyIfIdle(This,newVal)

#define IetScheduledTask_get_RunIfConnectedInternet(This,pVal)	\
    (This)->lpVtbl -> get_RunIfConnectedInternet(This,pVal)

#define IetScheduledTask_put_RunIfConnectedInternet(This,newVal)	\
    (This)->lpVtbl -> put_RunIfConnectedInternet(This,newVal)

#define IetScheduledTask_GetTrigger(This,nIndex)	\
    (This)->lpVtbl -> GetTrigger(This,nIndex)

#define IetScheduledTask_get_RestartOnIdleResume(This,pVal)	\
    (This)->lpVtbl -> get_RestartOnIdleResume(This,pVal)

#define IetScheduledTask_put_RestartOnIdleResume(This,newVal)	\
    (This)->lpVtbl -> put_RestartOnIdleResume(This,newVal)

#define IetScheduledTask_Run(This)	\
    (This)->lpVtbl -> Run(This)

#define IetScheduledTask_Terminate(This)	\
    (This)->lpVtbl -> Terminate(This)

#define IetScheduledTask_GetStatus(This,pVal)	\
    (This)->lpVtbl -> GetStatus(This,pVal)

#define IetScheduledTask_get_DaysInterval(This,pVal)	\
    (This)->lpVtbl -> get_DaysInterval(This,pVal)

#define IetScheduledTask_put_DaysInterval(This,newVal)	\
    (This)->lpVtbl -> put_DaysInterval(This,newVal)

#define IetScheduledTask_get_WeeksInterval(This,pVal)	\
    (This)->lpVtbl -> get_WeeksInterval(This,pVal)

#define IetScheduledTask_put_WeeksInterval(This,newVal)	\
    (This)->lpVtbl -> put_WeeksInterval(This,newVal)

#define IetScheduledTask_get_LastTimeRun(This,pVal)	\
    (This)->lpVtbl -> get_LastTimeRun(This,pVal)

#define IetScheduledTask_get_NextTimeRun(This,pVal)	\
    (This)->lpVtbl -> get_NextTimeRun(This,pVal)

#define IetScheduledTask_RemoveTrigger(This,nIndex)	\
    (This)->lpVtbl -> RemoveTrigger(This,nIndex)

#define IetScheduledTask_get_TriggerString(This,pVal)	\
    (This)->lpVtbl -> get_TriggerString(This,pVal)

#define IetScheduledTask_get_Disabled(This,pVal)	\
    (This)->lpVtbl -> get_Disabled(This,pVal)

#define IetScheduledTask_put_Disabled(This,newVal)	\
    (This)->lpVtbl -> put_Disabled(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_Program_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_Program_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_Program_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_Program_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_Parameters_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_Parameters_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_Parameters_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_Parameters_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_StartingDir_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_StartingDir_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_StartingDir_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_StartingDir_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_Comment_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_Comment_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_StartYear_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_StartYear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_StartYear_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_StartYear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_StartMonth_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_StartMonth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_StartMonth_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_StartMonth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_StartDay_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_StartDay_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_StartDay_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_StartDay_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_StartHour_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_StartHour_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_StartHour_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_StartHour_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_StartMinute_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_StartMinute_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_StartMinute_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_StartMinute_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_TriggerType_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_TriggerType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_TriggerType_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_TriggerType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_CreateTask_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    BSTR bstrTaskName,
    BOOL bFailifExists);


void __RPC_STUB IetScheduledTask_CreateTask_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_AddTrigger_Proxy( 
    IetScheduledTask __RPC_FAR * This);


void __RPC_STUB IetScheduledTask_AddTrigger_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_RemoveTask_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    BSTR bstrTaskName);


void __RPC_STUB IetScheduledTask_RemoveTask_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_Reset_Proxy( 
    IetScheduledTask __RPC_FAR * This);


void __RPC_STUB IetScheduledTask_Reset_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_EndYear_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_EndYear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_EndYear_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_EndYear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_EndMonth_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_EndMonth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_EndMonth_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_EndMonth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_EndDay_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_EndDay_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_EndDay_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_EndDay_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_MinutesInterval_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_MinutesInterval_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_MinutesInterval_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_MinutesInterval_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_TargetComputer_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_TargetComputer_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_TargetComputer_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_TargetComputer_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_MinutesDuration_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_MinutesDuration_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_MinutesDuration_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_MinutesDuration_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_DayOfTheWeek_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_DayOfTheWeek_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_DayOfTheWeek_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_DayOfTheWeek_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_DayOfTheMonth_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_DayOfTheMonth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_DayOfTheMonth_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_DayOfTheMonth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_Enum_Proxy( 
    IetScheduledTask __RPC_FAR * This);


void __RPC_STUB IetScheduledTask_Enum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_TaskCount_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_TaskCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_TaskItem_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_TaskItem_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_TaskName_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_TaskName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_DontStartIfOnBatteries_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_DontStartIfOnBatteries_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_DontStartIfOnBatteries_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_DontStartIfOnBatteries_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_KillIfOnBatteries_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_KillIfOnBatteries_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_KillIfOnBatteries_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_KillIfOnBatteries_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_Username_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_Username_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_Username_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_Username_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_Password_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_Password_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_Password_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_Password_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_Creator_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_Creator_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_Creator_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IetScheduledTask_put_Creator_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_LocalAccount_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_LocalAccount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_LocalAccount_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_LocalAccount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_EditTask_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    BSTR bstrTaskName);


void __RPC_STUB IetScheduledTask_EditTask_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_WhichWeek_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_WhichWeek_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_WhichWeek_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_WhichWeek_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_AddMonth_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    short nMonth);


void __RPC_STUB IetScheduledTask_AddMonth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_ActivateTask_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    BSTR bstrTaskName);


void __RPC_STUB IetScheduledTask_ActivateTask_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_ExitCode_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_ExitCode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_DeleteWhenDone_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_DeleteWhenDone_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_DeleteWhenDone_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_DeleteWhenDone_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_RunOnlyIfLogged_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_RunOnlyIfLogged_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_RunOnlyIfLogged_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_RunOnlyIfLogged_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_StartOnlyIfIdle_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_StartOnlyIfIdle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_StartOnlyIfIdle_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_StartOnlyIfIdle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_RunIfConnectedInternet_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_RunIfConnectedInternet_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_RunIfConnectedInternet_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_RunIfConnectedInternet_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_GetTrigger_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    short nIndex);


void __RPC_STUB IetScheduledTask_GetTrigger_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_RestartOnIdleResume_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_RestartOnIdleResume_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_RestartOnIdleResume_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_RestartOnIdleResume_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_Run_Proxy( 
    IetScheduledTask __RPC_FAR * This);


void __RPC_STUB IetScheduledTask_Run_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_Terminate_Proxy( 
    IetScheduledTask __RPC_FAR * This);


void __RPC_STUB IetScheduledTask_Terminate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_GetStatus_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_GetStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_DaysInterval_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_DaysInterval_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_DaysInterval_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_DaysInterval_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_WeeksInterval_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_WeeksInterval_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_WeeksInterval_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IetScheduledTask_put_WeeksInterval_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_LastTimeRun_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_LastTimeRun_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_NextTimeRun_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_NextTimeRun_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_RemoveTrigger_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    short nIndex);


void __RPC_STUB IetScheduledTask_RemoveTrigger_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_TriggerString_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_TriggerString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_get_Disabled_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *pVal);


void __RPC_STUB IetScheduledTask_get_Disabled_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IetScheduledTask_put_Disabled_Proxy( 
    IetScheduledTask __RPC_FAR * This,
    /* [in] */ BOOL newVal);


void __RPC_STUB IetScheduledTask_put_Disabled_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IetScheduledTask_INTERFACE_DEFINED__ */



#ifndef __SCHEDTASKLib_LIBRARY_DEFINED__
#define __SCHEDTASKLib_LIBRARY_DEFINED__

/* library SCHEDTASKLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_SCHEDTASKLib;

EXTERN_C const CLSID CLSID_etScheduledTask;

#ifdef __cplusplus

class DECLSPEC_UUID("656ED670-20F1-11D4-A450-0004AC9B6CD1")
etScheduledTask;
#endif
#endif /* __SCHEDTASKLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
