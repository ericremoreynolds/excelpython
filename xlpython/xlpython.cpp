#include "xlpython.h"

HINSTANCE hInstanceDLL;

// DLL entry point -- only stores the module handle for later use
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	hInstanceDLL = hinstDLL;

	return TRUE;
}

// change the dimensionality of a VBA array
HRESULT __stdcall XLPyDLLNDims(VARIANT* xlSource, int* xlDimension, bool *xlTranspose, VARIANT* xlDest)
{
	VariantClear(xlDest);

	try
	{
		// determine source dimensions
		int nSrcDims;
		int nSrcRows;
		int nSrcCols;
		SAFEARRAY* pSrcSA;
		if(0 == (xlSource->vt & VT_ARRAY))
		{
			nSrcDims = 0;
			nSrcRows = 1;
			nSrcCols = 1;
			pSrcSA = NULL;
		}
		else
		{
			if(xlSource->vt != (VT_VARIANT | VT_ARRAY) && xlSource->vt != (VT_VARIANT | VT_ARRAY | VT_BYREF))
				throw formatted_exception() << "Source array must be array of variants.";

			pSrcSA = (xlSource->vt & VT_BYREF) ? *xlSource->pparray : xlSource->parray;
			if(pSrcSA->cDims == 1)
			{
				nSrcDims = 1;
				nSrcRows = (int) pSrcSA->rgsabound->cElements;
				nSrcCols = 1;
			}
			else if(pSrcSA->cDims == 2)
			{
				nSrcDims = 2;
				nSrcRows = (int) pSrcSA->rgsabound[0].cElements;
				nSrcCols = (int) pSrcSA->rgsabound[1].cElements;
			}
			else
				throw formatted_exception() << "Source array must be either 1- or 2-dimensional.";
		}

		// determine dest dimension
		int nDestDims = *xlDimension;
		int nDestRows = nSrcRows;
		int nDestCols = nSrcCols;
		if(nDestDims == 1 && nSrcDims == 2)
		{
			if(nSrcRows != 1 && nSrcCols != 1)
				throw formatted_exception() << "When converting from 2- to 1-dimensional array, source must be (1 x n) or (n x 1).";
			if(nSrcCols != 1)
			{
				nDestRows = nSrcCols;
				nDestCols = nSrcRows;
			}
		}
		if(nDestDims == 0 && nSrcRows * nSrcCols != 1)
			throw formatted_exception() << "When converting array to scalar, source must contain only one element.";
		if(nDestDims == 2 && *xlTranspose)
		{
			int tmp = nDestRows;
			nDestRows = nDestCols;
			nDestCols = tmp;
		}

		// create destination safe array -- note that if nDestDims == 0 then AutoSafeArrayCreate does nothing and leaves pointer == NULL
		// for some reason, with 2-dimensional array bounds get swapped around by SafeArrayCreate
		SAFEARRAYBOUND bounds[2];
		bounds[0].lLbound = 1;
		bounds[0].cElements = (ULONG) nDestDims == 2 ? nDestCols : nDestRows;
		bounds[1].lLbound = 1;
		bounds[1].cElements = (ULONG) nDestDims == 2 ? nDestRows : nDestCols;
		AutoSafeArrayCreate asac(VT_VARIANT, nDestDims, bounds);
	
		// copy the data -- note that if NULL is passed to AutoSafeArrayAccessData then it does nothing 
		{
			VARIANT* pSrcData = xlSource;
			AutoSafeArrayAccessData(pSrcSA, (void**) &pSrcData);

			VARIANT* pDestData = xlDest;
			AutoSafeArrayAccessData(asac.pSafeArray, (void**) &pDestData);

			for(int iDestRow=0; iDestRow<nDestRows; iDestRow++)
			{
				for(int iDestCol=0; iDestCol<nDestCols; iDestCol++)
				{
					// safe array data is column-major -- and treating col-major data as row-major is the same as transposing
					int destIdx = iDestRow * nDestCols + iDestCol;
					int srcIdx = *xlTranspose ? iDestRow + iDestCol * nDestRows : destIdx;

					VariantInit(&pDestData[destIdx]);
					VariantCopy(&pDestData[destIdx], &pSrcData[srcIdx]);
				}
			}
		}

		// if we created a safe array, return and release it
		if(asac.pSafeArray != NULL)
		{
			xlDest->vt = VT_VARIANT | VT_ARRAY;
			xlDest->parray = asac.pSafeArray;
			asac.pSafeArray = NULL;
		}

		return S_OK;
	}
	catch(const std::exception& e)
	{
		ToVariant(e.what(), xlDest);
		return E_FAIL;
	}
}

// entry point - returns existing interface if already available, otherwise tries to activate it
HRESULT __stdcall XLPyDLLActivate(VARIANT* xlResult, const char* xlConfigFileName)
{
	try
	{
		VariantClear(xlResult);

		// set default config file
		std::string configFilename = xlConfigFileName;
		if(configFilename.empty())
			configFilename = "xlpython.xlpy";
		Config* pConfig = Config::GetConfig(configFilename);

		// if interface object isn't already available try to create it
		if(pConfig->pInterface == NULL || !pConfig->CheckRPCServer())
			pConfig->ActivateRPCServer();

		// pass it back to VBA
		xlResult->vt = VT_DISPATCH;
		xlResult->pdispVal = pConfig->pInterface;
		xlResult->pdispVal->AddRef();

		return S_OK;
	}
	catch(const std::exception& e)
	{
		ToVariant(e.what(), xlResult);
		return E_FAIL;
	}
}
