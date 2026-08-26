/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Dec 12 18:18:19 2002
 */
/* Compiler settings for D:\cpp\Com\SchedTask\SchedTask.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_IetScheduledTask = {0x656ED66F,0x20F1,0x11D4,{0xA4,0x50,0x00,0x04,0xAC,0x9B,0x6C,0xD1}};


const IID LIBID_SCHEDTASKLib = {0x656ED663,0x20F1,0x11D4,{0xA4,0x50,0x00,0x04,0xAC,0x9B,0x6C,0xD1}};


const CLSID CLSID_etScheduledTask = {0x656ED670,0x20F1,0x11D4,{0xA4,0x50,0x00,0x04,0xAC,0x9B,0x6C,0xD1}};


#ifdef __cplusplus
}
#endif

