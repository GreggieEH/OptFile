

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Dec 13 16:41:31 2016
 */
/* Compiler settings for MyIDL.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __MyIDL_h_h__
#define __MyIDL_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IOPTFile_FWD_DEFINED__
#define __IOPTFile_FWD_DEFINED__
typedef interface IOPTFile IOPTFile;

#endif 	/* __IOPTFile_FWD_DEFINED__ */


#ifndef ___OPTFile_FWD_DEFINED__
#define ___OPTFile_FWD_DEFINED__
typedef interface _OPTFile _OPTFile;

#endif 	/* ___OPTFile_FWD_DEFINED__ */


#ifndef __OPTFile_FWD_DEFINED__
#define __OPTFile_FWD_DEFINED__

#ifdef __cplusplus
typedef class OPTFile OPTFile;
#else
typedef struct OPTFile OPTFile;
#endif /* __cplusplus */

#endif 	/* __OPTFile_FWD_DEFINED__ */


#ifndef __IGratingScanInfo_FWD_DEFINED__
#define __IGratingScanInfo_FWD_DEFINED__
typedef interface IGratingScanInfo IGratingScanInfo;

#endif 	/* __IGratingScanInfo_FWD_DEFINED__ */


#ifndef __ICalibrationStandard_FWD_DEFINED__
#define __ICalibrationStandard_FWD_DEFINED__
typedef interface ICalibrationStandard ICalibrationStandard;

#endif 	/* __ICalibrationStandard_FWD_DEFINED__ */


#ifndef __IMyContainer_FWD_DEFINED__
#define __IMyContainer_FWD_DEFINED__
typedef interface IMyContainer IMyContainer;

#endif 	/* __IMyContainer_FWD_DEFINED__ */


#ifndef __IMyExtended_FWD_DEFINED__
#define __IMyExtended_FWD_DEFINED__
typedef interface IMyExtended IMyExtended;

#endif 	/* __IMyExtended_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __OPTFile_LIBRARY_DEFINED__
#define __OPTFile_LIBRARY_DEFINED__

/* library OPTFile */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_OPTFile;

#ifndef __IOPTFile_DISPINTERFACE_DEFINED__
#define __IOPTFile_DISPINTERFACE_DEFINED__

/* dispinterface IOPTFile */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IOPTFile;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("618CCDF2-5921-44d5-8083-A33E71C28BB5")
    IOPTFile : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IOPTFileVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IOPTFile * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IOPTFile * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IOPTFile * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IOPTFile * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IOPTFile * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IOPTFile * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IOPTFile * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IOPTFileVtbl;

    interface IOPTFile
    {
        CONST_VTBL struct IOPTFileVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPTFile_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IOPTFile_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IOPTFile_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IOPTFile_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IOPTFile_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IOPTFile_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IOPTFile_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IOPTFile_DISPINTERFACE_DEFINED__ */


#ifndef ___OPTFile_DISPINTERFACE_DEFINED__
#define ___OPTFile_DISPINTERFACE_DEFINED__

/* dispinterface _OPTFile */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__OPTFile;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("51AC4DB2-3762-401f-8858-2E3321557636")
    _OPTFile : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _OPTFileVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _OPTFile * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _OPTFile * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _OPTFile * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _OPTFile * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _OPTFile * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _OPTFile * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _OPTFile * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _OPTFileVtbl;

    interface _OPTFile
    {
        CONST_VTBL struct _OPTFileVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _OPTFile_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _OPTFile_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _OPTFile_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _OPTFile_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _OPTFile_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _OPTFile_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _OPTFile_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___OPTFile_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_OPTFile;

#ifdef __cplusplus

class DECLSPEC_UUID("A7753A85-2C43-450b-8096-3979F1B625BD")
OPTFile;
#endif

#ifndef __IGratingScanInfo_DISPINTERFACE_DEFINED__
#define __IGratingScanInfo_DISPINTERFACE_DEFINED__

/* dispinterface IGratingScanInfo */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IGratingScanInfo;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("AA55B107-72B2-4586-85BB-15E364184AA2")
    IGratingScanInfo : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IGratingScanInfoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IGratingScanInfo * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IGratingScanInfo * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IGratingScanInfo * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IGratingScanInfo * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IGratingScanInfo * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IGratingScanInfo * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IGratingScanInfo * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IGratingScanInfoVtbl;

    interface IGratingScanInfo
    {
        CONST_VTBL struct IGratingScanInfoVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IGratingScanInfo_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IGratingScanInfo_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IGratingScanInfo_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IGratingScanInfo_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IGratingScanInfo_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IGratingScanInfo_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IGratingScanInfo_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IGratingScanInfo_DISPINTERFACE_DEFINED__ */


#ifndef __ICalibrationStandard_DISPINTERFACE_DEFINED__
#define __ICalibrationStandard_DISPINTERFACE_DEFINED__

/* dispinterface ICalibrationStandard */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_ICalibrationStandard;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("F3290F8D-E4B9-457B-AADD-7DBA7A6FD3D2")
    ICalibrationStandard : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct ICalibrationStandardVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICalibrationStandard * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICalibrationStandard * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICalibrationStandard * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICalibrationStandard * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICalibrationStandard * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICalibrationStandard * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICalibrationStandard * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } ICalibrationStandardVtbl;

    interface ICalibrationStandard
    {
        CONST_VTBL struct ICalibrationStandardVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICalibrationStandard_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICalibrationStandard_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICalibrationStandard_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICalibrationStandard_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICalibrationStandard_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICalibrationStandard_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICalibrationStandard_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __ICalibrationStandard_DISPINTERFACE_DEFINED__ */


#ifndef __IMyContainer_DISPINTERFACE_DEFINED__
#define __IMyContainer_DISPINTERFACE_DEFINED__

/* dispinterface IMyContainer */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IMyContainer;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("73F7C087-4BB0-4470-BFF7-A59F9FDDADD3")
    IMyContainer : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IMyContainerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMyContainer * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMyContainer * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMyContainer * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMyContainer * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMyContainer * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMyContainer * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMyContainer * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IMyContainerVtbl;

    interface IMyContainer
    {
        CONST_VTBL struct IMyContainerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMyContainer_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMyContainer_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMyContainer_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMyContainer_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMyContainer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMyContainer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMyContainer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IMyContainer_DISPINTERFACE_DEFINED__ */


#ifndef __IMyExtended_DISPINTERFACE_DEFINED__
#define __IMyExtended_DISPINTERFACE_DEFINED__

/* dispinterface IMyExtended */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IMyExtended;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("5A7834C6-8391-4F1F-ABF8-053774BC5C87")
    IMyExtended : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IMyExtendedVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMyExtended * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMyExtended * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMyExtended * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMyExtended * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMyExtended * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMyExtended * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMyExtended * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IMyExtendedVtbl;

    interface IMyExtended
    {
        CONST_VTBL struct IMyExtendedVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMyExtended_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMyExtended_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMyExtended_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMyExtended_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMyExtended_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMyExtended_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMyExtended_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IMyExtended_DISPINTERFACE_DEFINED__ */

#endif /* __OPTFile_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


