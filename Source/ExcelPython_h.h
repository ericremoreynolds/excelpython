

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Thu May 22 13:53:19 2014
 */
/* Compiler settings for ExcelPython.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
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


#ifndef __ExcelPython_h_h__
#define __ExcelPython_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IProgressCallback_FWD_DEFINED__
#define __IProgressCallback_FWD_DEFINED__
typedef interface IProgressCallback IProgressCallback;
#endif 	/* __IProgressCallback_FWD_DEFINED__ */


#ifndef __IPyObj_FWD_DEFINED__
#define __IPyObj_FWD_DEFINED__
typedef interface IPyObj IPyObj;
#endif 	/* __IPyObj_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_ExcelPython_0000_0000 */
/* [local] */ 

#define PYTHON_TARGET 27
#define LIBID_ExcelPython LIBID_ExcelPython27


extern RPC_IF_HANDLE __MIDL_itf_ExcelPython_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_ExcelPython_0000_0000_v0_0_s_ifspec;


#ifndef __ExcelPython27_LIBRARY_DEFINED__
#define __ExcelPython27_LIBRARY_DEFINED__

/* library ExcelPython27 */
/* [version][lcid][helpstring][uuid] */ 


EXTERN_C const IID LIBID_ExcelPython27;

#ifndef __IProgressCallback_INTERFACE_DEFINED__
#define __IProgressCallback_INTERFACE_DEFINED__

/* interface IProgressCallback */
/* [object][uuid] */ 


EXTERN_C const IID IID_IProgressCallback;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("f7cf0020-030e-11e2-a21f-0800200c9a66")
    IProgressCallback : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE SetStatus( 
            /* [in] */ BSTR *status) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IProgressCallbackVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IProgressCallback * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IProgressCallback * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IProgressCallback * This);
        
        HRESULT ( STDMETHODCALLTYPE *SetStatus )( 
            IProgressCallback * This,
            /* [in] */ BSTR *status);
        
        END_INTERFACE
    } IProgressCallbackVtbl;

    interface IProgressCallback
    {
        CONST_VTBL struct IProgressCallbackVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IProgressCallback_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IProgressCallback_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IProgressCallback_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IProgressCallback_SetStatus(This,status)	\
    ( (This)->lpVtbl -> SetStatus(This,status) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IProgressCallback_INTERFACE_DEFINED__ */


#ifndef __IPyObj_INTERFACE_DEFINED__
#define __IPyObj_INTERFACE_DEFINED__

/* interface IPyObj */
/* [object][dual][uuid] */ 


EXTERN_C const IID IID_IPyObj;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("78fc63e0-fdb5-11e1-a21f-0800200c9a66")
    IPyObj : public IDispatch
    {
    public:
        virtual HRESULT __stdcall Str( 
            /* [retval][out] */ VARIANT *Result) = 0;
        
        virtual HRESULT __stdcall Var( 
            /* [defaultvalue][in] */ LONG *xlDimensions,
            /* [retval][out] */ VARIANT *Result) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ VARIANT *Key,
            /* [retval][out] */ IPyObj **Result) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Attr( 
            /* [in] */ BSTR Attribute,
            /* [retval][out] */ IPyObj **Result) = 0;
        
        virtual /* [id][propget] */ HRESULT __stdcall get__NewEnum( 
            /* [retval][out] */ IUnknown **ppUnk) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ LONG *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IPyObjVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IPyObj * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IPyObj * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IPyObj * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IPyObj * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IPyObj * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IPyObj * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IPyObj * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        HRESULT ( __stdcall *Str )( 
            IPyObj * This,
            /* [retval][out] */ VARIANT *Result);
        
        HRESULT ( __stdcall *Var )( 
            IPyObj * This,
            /* [defaultvalue][in] */ LONG *xlDimensions,
            /* [retval][out] */ VARIANT *Result);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            IPyObj * This,
            /* [in] */ VARIANT *Key,
            /* [retval][out] */ IPyObj **Result);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE *get_Attr )( 
            IPyObj * This,
            /* [in] */ BSTR Attribute,
            /* [retval][out] */ IPyObj **Result);
        
        /* [id][propget] */ HRESULT ( __stdcall *get__NewEnum )( 
            IPyObj * This,
            /* [retval][out] */ IUnknown **ppUnk);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IPyObj * This,
            /* [retval][out] */ LONG *pVal);
        
        END_INTERFACE
    } IPyObjVtbl;

    interface IPyObj
    {
        CONST_VTBL struct IPyObjVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IPyObj_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IPyObj_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IPyObj_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IPyObj_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IPyObj_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IPyObj_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IPyObj_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IPyObj_Str(This,Result)	\
    ( (This)->lpVtbl -> Str(This,Result) ) 

#define IPyObj_Var(This,xlDimensions,Result)	\
    ( (This)->lpVtbl -> Var(This,xlDimensions,Result) ) 

#define IPyObj_get_Item(This,Key,Result)	\
    ( (This)->lpVtbl -> get_Item(This,Key,Result) ) 

#define IPyObj_get_Attr(This,Attribute,Result)	\
    ( (This)->lpVtbl -> get_Attr(This,Attribute,Result) ) 

#define IPyObj_get__NewEnum(This,ppUnk)	\
    ( (This)->lpVtbl -> get__NewEnum(This,ppUnk) ) 

#define IPyObj_get_Count(This,pVal)	\
    ( (This)->lpVtbl -> get_Count(This,pVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IPyObj_INTERFACE_DEFINED__ */



#ifndef __ExcelPython_MODULE_DEFINED__
#define __ExcelPython_MODULE_DEFINED__


/* module ExcelPython */
/* [dllname][helpstring] */ 

/* [vararg][helpstring][entry] */ HRESULT __stdcall PyDict( 
    /* [in] */ SAFEARRAY * *KeyValuePairs,
    /* [retval][out] */ VARIANT *Result);

/* [vararg][helpstring][entry] */ HRESULT __stdcall PyTuple( 
    /* [in] */ SAFEARRAY * *Elements,
    /* [retval][out] */ VARIANT *Result);

/* [vararg][helpstring][entry] */ HRESULT __stdcall PyList( 
    /* [in] */ SAFEARRAY * *Elements,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyBuiltin( 
    /* [in] */ BSTR Method,
    /* [optional][in] */ VARIANT *Args,
    /* [optional][in] */ VARIANT *KwArgs,
    /* [retval][out] */ VARIANT *Result);

/* [vararg][helpstring][entry] */ HRESULT __stdcall PyStr( 
    /* [in] */ VARIANT *Value,
    /* [in] */ SAFEARRAY * *FormatArgs,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyRepr( 
    /* [in] */ VARIANT *Value,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyLen( 
    /* [in] */ VARIANT *Value,
    /* [retval][out] */ int *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyType( 
    /* [in] */ VARIANT *Value,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyCall( 
    /* [in] */ VARIANT *Instance,
    /* [defaultvalue][in] */ BSTR Method,
    /* [optional][in] */ VARIANT *Args,
    /* [optional][in] */ VARIANT *KwArgs,
    /* [defaultvalue][in] */ BOOL Console,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyIter( 
    /* [in] */ VARIANT *Instance,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyNext( 
    /* [in] */ VARIANT *Iterator,
    /* [out] */ VARIANT *Element,
    /* [retval][out] */ VARIANT_BOOL *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyGet( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ BSTR Attribute,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PySet( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ BSTR Attribute,
    /* [in] */ VARIANT *Value);

/* [helpstring][entry] */ HRESULT __stdcall PyGetAttr( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ BSTR Attribute,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PySetAttr( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ BSTR Attribute,
    /* [in] */ VARIANT *Value);

/* [helpstring][entry] */ HRESULT __stdcall PyDelAttr( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ BSTR Attribute);

/* [helpstring][entry] */ HRESULT __stdcall PyHasAttr( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ BSTR Attribute,
    /* [retval][out] */ VARIANT_BOOL *Result);

/* [entry] */ HRESULT __stdcall PyGetItem( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ VARIANT *Key,
    /* [retval][out] */ VARIANT *Result);

/* [entry] */ HRESULT __stdcall PySetItem( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ VARIANT *Key,
    /* [in] */ VARIANT *Value);

/* [entry] */ HRESULT __stdcall PyDelItem( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ VARIANT *Key);

/* [helpstring][entry] */ HRESULT __stdcall PyContains( 
    /* [in] */ VARIANT *Instance,
    /* [in] */ VARIANT *Value,
    /* [retval][out] */ VARIANT_BOOL *Result);

/* [entry] */ HRESULT __stdcall PyEval( 
    /* [in] */ BSTR Expression,
    /* [optional][in] */ VARIANT *Locals,
    /* [optional][in] */ VARIANT *Globals,
    /* [defaultvalue][in] */ BSTR AddPath,
    /* [defaultvalue][in] */ BSTR Path,
    /* [retval][out] */ IPyObj **Result);

/* [entry] */ HRESULT __stdcall PyExec( 
    /* [in] */ BSTR Statement,
    /* [optional][in] */ VARIANT *Locals,
    /* [optional][in] */ VARIANT *Globals,
    /* [defaultvalue][in] */ BSTR AddPath,
    /* [defaultvalue][in] */ BSTR Path);

/* [entry] */ HRESULT __stdcall PyModule( 
    /* [in] */ BSTR Name,
    /* [defaultvalue][in] */ VARIANT_BOOL *Reload,
    /* [defaultvalue][in] */ BSTR AddPath,
    /* [defaultvalue][in] */ BSTR Path,
    /* [retval][out] */ VARIANT *Result);

/* [entry] */ HRESULT __stdcall PyIsObject( 
    /* [in] */ VARIANT *Value,
    /* [retval][out] */ VARIANT_BOOL *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyToVariant( 
    /* [in] */ VARIANT *Value,
    /* [defaultvalue][in] */ LONG *Dimensions,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyVar( 
    /* [in] */ VARIANT *Value,
    /* [defaultvalue][in] */ LONG *Dimensions,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyToObject( 
    /* [in] */ VARIANT *Value,
    /* [defaultvalue][in] */ LONG *Dimensions,
    /* [retval][out] */ VARIANT *Result);

/* [helpstring][entry] */ HRESULT __stdcall PyObj( 
    /* [in] */ VARIANT *Value,
    /* [defaultvalue][in] */ LONG *Dimensions,
    /* [retval][out] */ VARIANT *Result);

#endif /* __ExcelPython_MODULE_DEFINED__ */
#endif /* __ExcelPython27_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


