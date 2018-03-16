// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// RhinoVariantHelpers.cpp

#include "StdAfx.h"
#include "RhinoVariantHelpers.h"

CRhinoDoc* CRhinoVariantHelpers::Document()
{
  return RhinoApp().ActiveDoc();
}

void CRhinoVariantHelpers::RedrawDocument()
{
  CRhinoDoc* doc = Document();
  if (nullptr != doc)
    doc->Redraw();
}

void CRhinoVariantHelpers::RegenDocument()
{
  CRhinoDoc* doc = Document();
  if (nullptr != doc)
    doc->Regen();
}

bool CRhinoVariantHelpers::StringToUuid(const wchar_t* uuid_str, ON_UUID& uuid)
{
  bool rc = false;
  if (uuid_str && uuid_str[0])
  {
    ON_wString str(uuid_str);
    str.TrimLeftAndRight();
    if (str.Length())
    {
      ON_UUID temp_uuid = ON_UuidFromString(str);
      if (ON_UuidIsNotNil(temp_uuid))
      {
        uuid = temp_uuid;
        rc = true;
      }
    }
  }
  return rc;
}

CString CRhinoVariantHelpers::StringFromUuid(const ON_UUID& uuid)
{
  CString str;
  StringFromUuid(uuid, str);
  return str;
}

bool CRhinoVariantHelpers::StringFromUuid(const ON_UUID& uuid, CString& uuid_str)
{
  bool rc = false;
  ON_wString str;
  if (StringFromUuid(uuid, str))
  {
    uuid_str = str;
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::StringFromUuid(const ON_UUID& uuid, ON_wString& uuid_str)
{
  bool rc = false;
  if (ON_UuidToString(uuid, uuid_str))
    rc = true;
  return rc;
}

int CRhinoVariantHelpers::UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, ON_ClassArray<ON_wString>& strings)
{
  int i;
  for (i = 0; i < uuids.Count(); i++)
  {
    ON_wString uuid_str;
    if (ON_UuidToString(uuids[i], uuid_str))
      strings.Append(uuid_str);
  }
  return strings.Count();
}

int CRhinoVariantHelpers::UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, CStringArray& strings)
{
  int i;
  for (i = 0; i < uuids.Count(); i++)
  {
    CString uuid_str;
    if (StringFromUuid(uuids[i], uuid_str))
      strings.Add(uuid_str);
  }
  return (int)strings.GetCount();
}

bool CRhinoVariantHelpers::RhinoObjRef(const wchar_t* uuid_str, CRhinoObjRef& obj_ref)
{
  CRhinoDoc* doc = Document();
  if (nullptr == doc || nullptr == uuid_str || 0 == uuid_str[0])
    return false;

  ON_wString str(uuid_str);
  str.TrimLeftAndRight();
  if (str.IsEmpty())
    return false;

  ON_UUID uuid = ON_UuidFromString(str);
  if (ON_UuidIsNotNil(uuid))
  {
    const CRhinoObject* obj = CRhinoObject::FromId(doc->RuntimeSerialNumber(), uuid);
    if (nullptr != obj)
    {
      obj_ref = CRhinoObjRef(obj);
      return true;
    }
  }

  return false;
}

bool CRhinoVariantHelpers::IsVariantEmpty(const VARIANT& va)
{
  bool rc = false;

  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    rc = true;
  else
  {
    const VARIANT* pva = &va;
    while (pva->vt == (VT_BYREF | VT_VARIANT))
      pva = pva->pvarVal;

    if (pva->vt == VT_EMPTY)
      rc = true;
  }

  return rc;
}

bool CRhinoVariantHelpers::IsVariantNull(const VARIANT& va)
{
  bool rc = false;

  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return rc;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  if (pva->vt == VT_NULL)
    rc = true;
  else if (pva->vt == VT_I2 && pva->iVal == 1) // value of vbNull constant
    rc = true;

  return rc;
}

bool CRhinoVariantHelpers::IsVariantNullOrEmpty(const VARIANT& va)
{
  return IsVariantNull(va) || IsVariantEmpty(va);
}

bool CRhinoVariantHelpers::IsVariantBoolean(const VARIANT& va)
{
  bool rc = false;

  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return rc;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  if (pva->vt == VT_BOOL)
    rc = true;

  return rc;
}

bool CRhinoVariantHelpers::IsVariantInteger(const VARIANT& va)
{
  bool rc = false;

  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return rc;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  if (pva->vt == VT_I2 || pva->vt == VT_I4)
    rc = true;

  return rc;
}

bool CRhinoVariantHelpers::IsVariantDouble(const VARIANT& va)
{
  bool rc = false;

  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return rc;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  if (pva->vt == VT_R4 || pva->vt == VT_R8)
    rc = true;

  return rc;
}

bool CRhinoVariantHelpers::IsVariantNumber(const VARIANT& va)
{
  bool rc = false;

  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return rc;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  if (pva->vt == VT_I2 || pva->vt == VT_I4 || pva->vt == VT_R4 || pva->vt == VT_R8)
    rc = true;

  return rc;
}

bool CRhinoVariantHelpers::IsVariantString(const VARIANT& va)
{
  bool rc = false;

  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return rc;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  if (pva->vt == VT_BSTR)
    rc = true;

  return rc;
}

bool CRhinoVariantHelpers::IsVariantPoint(const VARIANT& va)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
    return false;

  if (SafeArrayGetDim(psa) != 1)
    return false;

  ON_3dPoint pt;
  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPoint(psa, pt, true);

  return SafeArrayToPoint(psa, pt);
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, bool& b, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  VARIANT_BOOL boolOut = VARIANT_TRUE;
  switch (pva->vt)
  {
  case VT_BOOL:
    boolOut = pva->boolVal;
    break;

  case VT_I2:
    VarBoolFromI2(pva->iVal, &boolOut);
    break;

  case VT_I4:
    VarBoolFromI4(pva->lVal, &boolOut);
    break;

  case VT_R4:
    VarBoolFromR4(pva->fltVal, &boolOut);
    break;

  case VT_R8:
    VarBoolFromR8(pva->dblVal, &boolOut);
    break;

  case VT_BSTR:
    VarBoolFromStr(pva->bstrVal, LOCALE_INVARIANT, 0, &boolOut);
    break;

  case VT_DATE:
    VarBoolFromDate(pva->date, &boolOut);
    break;

  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_boolean_required);
  }
  return false;
  }
  b = (boolOut ? true : false);

  return true;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, int& n, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  long l = 0;
  switch (pva->vt)
  {
  case VT_BOOL:
    VarI4FromBool(pva->boolVal, &l);
    break;

  case VT_I2:
    VarI4FromI2(pva->iVal, &l);
    break;

  case VT_I4:
    l = pva->lVal;
    break;

  case VT_R4:
    VarI4FromR4(pva->fltVal, &l);
    break;

  case VT_R8:
    VarI4FromR8(pva->dblVal, &l);
    break;

  case VT_BSTR:
    VarI4FromStr(pva->bstrVal, LOCALE_INVARIANT, 0, &l);
    break;

  case VT_DATE:
    VarI4FromDate(pva->date, &l);
    break;

  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_float_required);
  }
  return false;
  }
  n = (int)l;

  return true;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, float& f, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    VarR4FromBool(pva->boolVal, &f);
    break;

  case VT_I2:
    VarR4FromI2(pva->iVal, &f);
    break;

  case VT_I4:
    VarR4FromI4(pva->lVal, &f);
    break;

  case VT_R4:
    f = pva->fltVal;
    break;

  case VT_R8:
    VarR4FromR8(pva->dblVal, &f);
    break;

  case VT_BSTR:
    VarR4FromStr(pva->bstrVal, LOCALE_INVARIANT, 0, &f);
    break;

  case VT_DATE:
    VarR4FromDate(pva->date, &f);
    break;

  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_float_required);
  }
  return false;
  }

  return true;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, double& d, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    VarR8FromBool(pva->boolVal, &d);
    break;

  case VT_I2:
    VarR8FromI2(pva->iVal, &d);
    break;

  case VT_I4:
    VarR8FromI4(pva->lVal, &d);
    break;

  case VT_R4:
    VarR8FromR4(pva->fltVal, &d);
    break;

  case VT_R8:
    d = pva->dblVal;
    break;

  case VT_BSTR:
    VarR8FromStr(pva->bstrVal, LOCALE_INVARIANT, 0, &d);
    break;

  case VT_DATE:
    VarR8FromDate(pva->date, &d);
    break;

  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_float_required);
  }
  return false;
  }

  return true;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_wString& str, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  CComBSTR bstr;
  switch (pva->vt)
  {
  case VT_BOOL:
    VarBstrFromBool(pva->boolVal, LOCALE_INVARIANT, 0, &bstr);
    break;

  case VT_I2:
    VarBstrFromI2(pva->iVal, LOCALE_INVARIANT, 0, &bstr);
    break;

  case VT_I4:
    VarBstrFromI4(pva->lVal, LOCALE_INVARIANT, 0, &bstr);
    str = bstr;
    break;

  case VT_UINT:
    VarBstrFromUint(pva->lVal, LOCALE_INVARIANT, 0, &bstr);
    str = bstr;
    break;

  case VT_R4:
    VarBstrFromR4(pva->fltVal, LOCALE_INVARIANT, 0, &bstr);
    str = bstr;
    break;

  case VT_R8:
    VarBstrFromR8(pva->dblVal, LOCALE_INVARIANT, 0, &bstr);
    str = bstr;
    break;

  case VT_BSTR:
    bstr = pva->bstrVal;
    break;

  case VT_DATE:
    VarBstrFromDate(pva->date, LOCALE_INVARIANT, 0, &bstr);
    break;

  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_string_required);
  }
  return false;
  }
  str = bstr;

  return true;
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_SimpleArray<bool>& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToBooleanArray(psa, arr, bQuiet);

  return SafeArrayToBooleanArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_SimpleArray<int>& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToIntegerArray(psa, arr, bQuiet);

  return SafeArrayToIntegerArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_SimpleArray<float>& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToFloatArray(psa, arr, bQuiet);

  return SafeArrayToFloatArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_SimpleArray<double>& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToDoubleArray(psa, arr, bQuiet);

  return SafeArrayToDoubleArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_ClassArray<ON_wString>& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToStringArray(psa, arr, bQuiet);

  return SafeArrayToStringArray(psa, arr);
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_2dPoint& pt, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return false;
  }

  bool rc = false;
  if (psa->fFeatures & FADF_VARIANT)
    rc = VariantArrayToPoint(psa, pt, bQuiet);
  else
    rc = SafeArrayToPoint(psa, pt);

  if (false == rc)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
  }
  else
    rc = pt.IsValid();

  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_3dPoint& pt, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return false;
  }

  bool rc = false;
  if (psa->fFeatures & FADF_VARIANT)
    rc = VariantArrayToPoint(psa, pt, bQuiet);
  else
    rc = SafeArrayToPoint(psa, pt);

  if (false == rc)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
  }
  else
    rc = pt.IsValid();

  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_3fPoint& pt, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return false;
  }

  bool rc = false;
  if (psa->fFeatures & FADF_VARIANT)
    rc = VariantArrayToPoint(psa, pt, bQuiet);
  else
    rc = SafeArrayToPoint(psa, pt);

  if (false == rc)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
  }

  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_4fPoint& pt, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return false;
  }

  bool rc = false;
  if (psa->fFeatures & FADF_VARIANT)
    rc = VariantArrayToPoint(psa, pt, bQuiet);
  else
    rc = SafeArrayToPoint(psa, pt);

  if (false == rc)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
  }

  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_4dPoint& pt, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return false;
  }

  bool rc = false;
  if (psa->fFeatures & FADF_VARIANT)
    rc = VariantArrayToPoint(psa, pt, bQuiet);
  else
    rc = SafeArrayToPoint(psa, pt);

  if (false == rc)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
  }
  else
    rc = pt.IsValid();

  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_3dVector& v, bool bQuiet)
{
  ON_3dPoint pt;
  bool rc = ConvertVariant(va, pt, bQuiet);
  if (rc)
    v = pt;

  return rc;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_2dPointArray& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);

  return SafeArrayToPointArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_3dPointArray& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);

  return SafeArrayToPointArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_3fPointArray& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);

  return SafeArrayToPointArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_4dPointArray& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);

  return SafeArrayToPointArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_4fPointArray& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);

  return SafeArrayToPointArray(psa, arr);
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_SimpleArray<ON_Color>& arr, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return 0;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt & VT_ARRAY)
  {
    if (pva->vt & VT_BYREF)
      psa = *pva->pparray;
    else
      psa = pva->parray;
  }

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToColorArray(psa, arr, bQuiet);

  return SafeArrayToColorArray(psa, arr);
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_Color& color, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    color = (pva->boolVal ? 0 : RGB(255, 255, 255));
    break;
  case VT_I2:
    color = RHINO_CLAMP((unsigned int)pva->iVal, RGB(0, 0, 0), RGB(255, 255, 255));
    break;
  case VT_I4:
    color = RHINO_CLAMP((unsigned int)pva->lVal, RGB(0, 0, 0), RGB(255, 255, 255));
    break;
  case VT_R4:
    color = RHINO_CLAMP((unsigned int)pva->fltVal, RGB(0, 0, 0), RGB(255, 255, 255));
    break;
  case VT_R8:
    color = RHINO_CLAMP((unsigned int)pva->dblVal, RGB(0, 0, 0), RGB(255, 255, 255));
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_integer_required);
  }
  return false;
  }

  return true;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_Xform& xform, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  SAFEARRAY* psa = 0;
  if (pva->vt == (VT_ARRAY | VT_VARIANT))
    psa = pva->parray;
  else if (pva->vt == (VT_ARRAY | VT_VARIANT | VT_BYREF))
    psa = *pva->pparray;

  if (psa == 0)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 2)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_one_dim_required);
    return false;
  }

  long col_count, col_lower, col_upper;
  SafeArrayGetLBound(psa, 1, &col_lower);
  SafeArrayGetUBound(psa, 1, &col_upper);

  long row_count, row_lower, row_upper;
  SafeArrayGetLBound(psa, 2, &row_lower);
  SafeArrayGetUBound(psa, 2, &row_upper);

  long col_size = col_upper - col_lower;
  long row_size = row_upper - row_lower;
  if (col_size != 3 && row_size != 3)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_two_dim_required);
    return false;
  }

  double matrix[4][4];
  long index[2];
  for (col_count = col_lower; col_count <= col_upper; col_count++)
  {
    for (row_count = row_lower; row_count <= row_upper; row_count++)
    {
      matrix[row_count][col_count] = 0.0;

      VARIANT element;
      VariantInit(&element);

      index[0] = row_count;
      index[1] = col_count;

      SafeArrayGetElement(psa, index, &element);

      double d = 0.0;
      if (ConvertVariant(element, d, bQuiet))
        matrix[row_count][col_count] = d;

      VariantClear(&element);
    }
  }

  ON_Xform temp(matrix);
  xform = temp;

  return xform.IsValid() ? true : false;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_Plane& plane, bool bQuiet)
{
  bool rc = false;
  ON_3dPointArray arr;
  if (ConvertVariant(va, arr, bQuiet) && arr.Count() > 2)
  {
    plane.CreateFromFrame(arr[0], arr[1], arr[2]);
    rc = (plane.IsValid() ? true : false);
  }
  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_Line& line, bool bQuiet)
{
  bool rc = false;
  ON_3dPointArray arr;
  if (ConvertVariant(va, arr, bQuiet) && arr.Count() == 2)
  {
    line.Create(arr[0], arr[1]);
    rc = (line.IsValid() ? true : false);
  }
  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, CRhinoObjRef& object_ref, bool bQuiet)
{
  bool rc = false;
  ON_wString str;
  if (ConvertVariant(va, str, bQuiet))
  {
    if (RhinoObjRef(str, object_ref))
      rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_Interval& interval, bool bQuiet)
{
  UNREFERENCED_PARAMETER(bQuiet);

  bool rc = false;
  ON_SimpleArray<double> arr;
  if (ConvertVariant(va, arr) && arr.Count() == 2)
  {
    interval.Set(arr[0], arr[1]);
    rc = true;
  }
  return rc;
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_ClassArray<CRhinoObjRef>& object_refs, bool bQuiet)
{
  if (IsVariantString(va))
  {
    CRhinoObjRef ref;
    if (ConvertVariant(va, ref, bQuiet))
      object_refs.Append(ref);
  }
  else
  {
    ON_ClassArray<ON_wString> strings;
    if (ConvertVariant(va, strings, bQuiet))
    {
      const int string_count = strings.Count();
      int i;
      for (i = 0; i < string_count; i++)
      {
        CRhinoObjRef ref;
        if (RhinoObjRef(strings[i], ref))
          object_refs.Append(ref);
      }
    }
  }
  return object_refs.Count();
}

bool CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_UUID& uuid, bool bQuiet)
{
  bool rc = false;
  ON_wString str;
  if (ConvertVariant(va, str, bQuiet))
    rc = StringToUuid(str, uuid);
  return rc;
}

int CRhinoVariantHelpers::ConvertVariant(const VARIANT& va, ON_SimpleArray<ON_UUID>& uuids, bool bQuiet)
{
  if (IsVariantString(va))
  {
    ON_UUID uuid = ON_nil_uuid;
    if (ConvertVariant(va, uuid, bQuiet))
      uuids.Append(uuid);
  }
  else
  {
    ON_ClassArray<ON_wString> strings;
    if (ConvertVariant(va, strings, bQuiet))
    {
      const int string_count = strings.Count();
      int i;
      for (i = 0; i < string_count; i++)
      {
        ON_UUID uuid = ON_nil_uuid;
        if (StringToUuid(strings[i], uuid))
          uuids.Append(uuid);
      }
    }
  }
  return uuids.Count();
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_2dPoint& pt, COleSafeArray& sa)
{
  COleVariant var[2];
  var[0] = pt.x;
  var[1] = pt.y;
  sa.CreateOneDim(VT_VARIANT, 2, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_2fPoint& pt, COleSafeArray& sa)
{
  COleVariant var[2];
  var[0] = pt.x;
  var[1] = pt.y;
  sa.CreateOneDim(VT_VARIANT, 2, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3dPoint& pt, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3fPoint& pt, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_4dPoint& pt, COleSafeArray& sa)
{
  COleVariant var[4];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  var[3] = pt.w;
  sa.CreateOneDim(VT_VARIANT, 4, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_4fPoint& pt, COleSafeArray& sa)
{
  COleVariant var[4];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  var[3] = pt.w;
  sa.CreateOneDim(VT_VARIANT, 4, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3dVector& v, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = v.x;
  var[1] = v.y;
  var[2] = v.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3fVector& v, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = v.x;
  var[1] = v.y;
  var[2] = v.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_SimpleArray<bool>& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      VARIANT v;
      VariantInit(&v);
      v.vt = VT_BOOL;
      v.boolVal = (arr[i] ? VARIANT_TRUE : VARIANT_FALSE);
      sa.PutElement(index, &v);
      VariantClear(&v);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_SimpleArray<int>& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      VARIANT v;
      VariantInit(&v);
      v.vt = VT_I4;
      v.lVal = arr[i];
      sa.PutElement(index, &v);
      VariantClear(&v);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_SimpleArray<double>& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      VARIANT v;
      VariantInit(&v);
      v.vt = VT_R8;
      v.dblVal = arr[i];
      sa.PutElement(index, &v);
      VariantClear(&v);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_ClassArray<ON_wString>& arr, COleSafeArray& sa, bool bAllowEmptyStrings)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      VARIANT v;
      VariantInit(&v);
      ON_wString w = arr[i];
      CString s(w);
      if (s.IsEmpty() && bAllowEmptyStrings == false)
      {
        v.vt = VT_NULL;
        sa.PutElement(index, &v);
      }
      else
      {
        v.vt = VT_BSTR;
        v.bstrVal = s.AllocSysString();
        sa.PutElement(index, &v);
        SysFreeString(v.bstrVal);
      }
      VariantClear(&v);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const CStringArray& arr, COleSafeArray& sa, bool bAllowEmptyStrings)
{
  bool rc = false;
  int i, count = (int)arr.GetCount(); // (int) cast for x64 builds
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      VARIANT v;
      VariantInit(&v);
      CString s = arr.GetAt(i);
      if (s.IsEmpty() && bAllowEmptyStrings == false)
      {
        v.vt = VT_NULL;
        sa.PutElement(index, &v);
      }
      else
      {
        v.vt = VT_BSTR;
        v.bstrVal = s.AllocSysString();
        sa.PutElement(index, &v);
        SysFreeString(v.bstrVal);
      }
      VariantClear(&v);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_SimpleArray<ON_UUID>& arr, COleSafeArray& sa)
{
  bool rc = false;
  int count = arr.Count();
  if (count > 0)
  {
    CStringArray strings;
    count = UuidArrayToStringArray(arr, strings);
    if (count > 0)
      rc = CreateSafeArray(strings, sa);
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_ClassArray<CRhinoObjRef>& arr, COleSafeArray& sa)
{
  bool rc = false;
  int arr_count = arr.Count();
  if (arr_count > 0)
  {
    ON_ClassArray<ON_wString> strings(arr_count);
    for (int i = 0; i < arr_count; i++)
    {
      ON_wString str;
      if (ON_UuidToString(arr[i].ObjectUuid(), str))
        strings.Append(str);
    }

    if (strings.Count() > 0)
      rc = CreateSafeArray(strings, sa);
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_2dPointArray& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_2fPointArray& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3fPointArray& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3dPointArray& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_4fPointArray& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3dVectorArray& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_3fVectorArray& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_Xform& xform, COleSafeArray& sa)
{
  DWORD numElements[] = { 4, 4 };
  sa.Create(VT_VARIANT, 2, numElements);
  int i, j;
  long index[2];
  for (i = 0; i < 4; i++)
  {
    for (j = 0; j < 4; j++)
    {
      index[0] = (long)i;
      index[1] = (long)j;
      VARIANT v;
      VariantInit(&v);
      v.vt = VT_R8;
      v.dblVal = xform.m_xform[i][j];
      sa.PutElement(index, &v);
      VariantClear(&v);
    }
  }
  return true;
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_Plane& plane, COleSafeArray& sa)
{
  ON_3dPointArray arr;
  arr.Append(plane.origin);
  arr.Append(plane.xaxis);
  arr.Append(plane.yaxis);
  arr.Append(plane.zaxis);
  return CreateSafeArray(arr, sa);
}

bool CRhinoVariantHelpers::CreateSafeArray(const ON_SimpleArray<ON_Plane>& arr, COleSafeArray& sa)
{
  bool rc = false;
  int i, count = arr.Count();
  if (count > 0)
  {
    DWORD numElements[1];
    numElements[0] = (DWORD)count;
    sa.Create(VT_VARIANT, 1, numElements);
    long index[1];
    for (i = 0; i < count; i++)
    {
      index[0] = (long)i;
      COleSafeArray osa;
      CreateSafeArray(arr[i], osa);
      VARIANT va = osa.Detach();
      sa.PutElement(index, (VARIANT*)&va);
      VariantClear(&va);
    }
    rc = true;
  }
  return rc;
}

int CRhinoVariantHelpers::ObjectArrayToStringArray(const ON_SimpleArray<const CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings)
{
  int i, num_added = 0;
  for (i = 0; i < objects.Count(); i++)
  {
    const CRhinoObject* obj = objects[i];
    if (obj)
    {
      ON_wString str;
      if (ON_UuidToString(obj->Attributes().m_uuid, str))
      {
        strings.Append(str);
        num_added++;
      }
    }
  }
  return num_added++;
}

int CRhinoVariantHelpers::ObjectArrayToStringArray(const ON_SimpleArray<CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings)
{
  int i, num_added = 0;
  for (i = 0; i < objects.Count(); i++)
  {
    const CRhinoObject* obj = objects[i];
    if (obj)
    {
      ON_wString str;
      if (ON_UuidToString(obj->Attributes().m_uuid, str))
      {
        strings.Append(str);
        num_added++;
      }
    }
  }
  return num_added++;
}

int CRhinoVariantHelpers::VariantArrayToBooleanArray(SAFEARRAY* psa, ON_SimpleArray<bool>& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        bool b = false;
        if (ConvertVariant(pvData[i], b, bQuiet))
          arr.Append(b);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToIntegerArray(SAFEARRAY* psa, ON_SimpleArray<int>& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        int n = 0;
        if (ConvertVariant(pvData[i], n, bQuiet))
          arr.Append(n);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToFloatArray(SAFEARRAY* psa, ON_SimpleArray<float>& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        float f = 0.0;
        if (ConvertVariant(pvData[i], f, bQuiet))
          arr.Append(f);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToDoubleArray(SAFEARRAY* psa, ON_SimpleArray<double>& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        double d = 0.0;
        if (ConvertVariant(pvData[i], d, bQuiet))
          arr.Append(d);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToStringArray(SAFEARRAY* psa, ON_ClassArray<ON_wString>& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        ON_wString str;
        if (ConvertVariant(pvData[i], str, bQuiet))
          arr.Append(str);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToColorArray(SAFEARRAY* psa, ON_SimpleArray<ON_Color>& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        ON_Color c;
        if (ConvertVariant(pvData[i], c, bQuiet))
          arr.Append(c);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return arr.Count();
}

bool CRhinoVariantHelpers::VariantArrayToPoint(SAFEARRAY* psa, ON_2dPoint& pt, bool bQuiet)
{
  ASSERT(psa);
  bool rc = false;
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu) && (upper - lower == 1))
  {
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      double p[] = { 0.0, 0.0 };
      for (i = lower; i <= upper; i++)
        for (i = lower; i <= upper; i++)
        {
          rc = ConvertVariant(pvData[i], p[i], bQuiet);
          if (!rc)
            break;
        }
      SafeArrayUnaccessData(psa);
      if (rc)
      {
        pt.x = p[0];
        pt.y = p[1];
      }
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return rc;
}

bool CRhinoVariantHelpers::VariantArrayToPoint(SAFEARRAY* psa, ON_3dPoint& pt, bool bQuiet)
{
  ASSERT(psa);
  bool rc = false;
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu) && (upper - lower == 1 || upper - lower == 2))
  {
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      double p[] = { 0.0, 0.0, 0.0 };
      for (i = lower; i <= upper; i++)
      {
        rc = ConvertVariant(pvData[i], p[i], bQuiet);
        if (!rc)
          break;
      }
      SafeArrayUnaccessData(psa);
      if (rc)
      {
        pt.x = p[0];
        pt.y = p[1];
        pt.z = p[2];
      }
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return rc;
}

bool CRhinoVariantHelpers::VariantArrayToPoint(SAFEARRAY* psa, ON_3fPoint& pt, bool bQuiet)
{
  ASSERT(psa);
  bool rc = false;
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu) && (upper - lower == 1 || upper - lower == 2))
  {
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      float p[] = { 0.0, 0.0, 0.0 };
      for (i = lower; i <= upper; i++)
      {
        rc = ConvertVariant(pvData[i], p[i], bQuiet);
        if (!rc)
          break;
      }
      SafeArrayUnaccessData(psa);
      if (rc)
      {
        pt.x = p[0];
        pt.y = p[1];
        pt.z = p[2];
      }
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return rc;
}

bool CRhinoVariantHelpers::VariantArrayToPoint(SAFEARRAY* psa, ON_4dPoint& pt, bool bQuiet)
{
  ASSERT(psa);
  bool rc = false;
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu) && (upper - lower == 2 || upper - lower == 3))
  {
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      double p[] = { 0.0, 0.0, 0.0, 0.0 };
      for (i = lower; i <= upper; i++)
      {
        rc = ConvertVariant(pvData[i], p[i], bQuiet);
        if (!rc)
          break;
      }
      SafeArrayUnaccessData(psa);
      if (rc)
      {
        pt.x = p[0];
        pt.y = p[1];
        pt.z = p[2];
        pt.w = p[3];
      }
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return rc;
}

bool CRhinoVariantHelpers::VariantArrayToPoint(SAFEARRAY* psa, ON_4fPoint& pt, bool bQuiet)
{
  ASSERT(psa);
  bool rc = false;
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu) && (upper - lower == 2 || upper - lower == 3))
  {
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      float p[] = { 0.0, 0.0, 0.0, 0.0 };
      for (i = lower; i <= upper; i++)
      {
        rc = ConvertVariant(pvData[i], p[i], bQuiet);
        if (!rc)
          break;
      }
      SafeArrayUnaccessData(psa);
      if (rc)
      {
        pt.x = p[0];
        pt.y = p[1];
        pt.z = p[2];
        pt.w = p[3];
      }
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
  }

  return rc;
}

int CRhinoVariantHelpers::VariantArrayToPointArray(SAFEARRAY* psa, ON_2dPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 2;
    ON_SimpleArray<double> vals;
    int count = ConvertVariant(va, vals, bQuiet);
    if (count && count % stride == 0)
    {
      int i, point_count = count / stride;
      arr.SetCapacity(point_count);
      arr.SetCount(point_count);
      ON_2dPoint pt(0.0, 0.0);
      const double* ptr = vals.Array();
      for (i = 0; i < point_count; i++)
      {
        pt.x = ptr[0];
        pt.y = ptr[1];
        arr[i] = pt;
        ptr += stride;
      }
      return arr.Count();
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
      return 0;
    }
  }
  else
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        ON_2dPoint pt;
        if (ConvertVariant(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToPointArray(SAFEARRAY* psa, ON_3dPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 3;
    ON_SimpleArray<double> vals;
    int count = ConvertVariant(va, vals, bQuiet);
    if (count && count % stride == 0)
    {
      int i, point_count = count / stride;
      arr.SetCapacity(point_count);
      arr.SetCount(point_count);
      ON_3dPoint pt(0.0, 0.0, 0.0);
      const double* ptr = vals.Array();
      for (i = 0; i < point_count; i++)
      {
        pt.x = ptr[0];
        pt.y = ptr[1];
        pt.z = ptr[2];
        arr[i] = pt;
        ptr += stride;
      }
      return arr.Count();
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
      return 0;
    }
  }
  else
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        ON_3dPoint pt;
        if (ConvertVariant(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToPointArray(SAFEARRAY* psa, ON_3fPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 3;
    ON_SimpleArray<float> vals;
    int count = ConvertVariant(va, vals, bQuiet);
    if (count && count % stride == 0)
    {
      int i, point_count = count / stride;
      arr.SetCapacity(point_count);
      arr.SetCount(point_count);
      ON_3fPoint pt(0.0, 0.0, 0.0);
      const float* ptr = vals.Array();
      for (i = 0; i < point_count; i++)
      {
        pt.x = ptr[0];
        pt.y = ptr[1];
        pt.z = ptr[2];
        arr[i] = pt;
        ptr += stride;
      }
      return arr.Count();
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
      return 0;
    }
  }
  else
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        ON_3fPoint pt;
        if (ConvertVariant(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToPointArray(SAFEARRAY* psa, ON_4dPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 4;
    ON_SimpleArray<double> vals;
    int count = ConvertVariant(va, vals, bQuiet);
    if (count && count % stride == 0)
    {
      int i, point_count = count / stride;
      arr.SetCapacity(point_count);
      arr.SetCount(point_count);
      ON_4dPoint pt(0.0, 0.0, 0.0, 0.0);
      const double* ptr = vals.Array();
      for (i = 0; i < point_count; i++)
      {
        pt.x = ptr[0];
        pt.y = ptr[1];
        pt.z = ptr[2];
        pt.w = ptr[3];
        arr[i] = pt;
        ptr += stride;
      }
      return arr.Count();
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
      return 0;
    }
  }
  else
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        ON_4dPoint pt;
        if (ConvertVariant(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoVariantHelpers::VariantArrayToPointArray(SAFEARRAY* psa, ON_4fPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 4;
    ON_SimpleArray<float> vals;
    int count = ConvertVariant(va, vals, bQuiet);
    if (count && count % stride == 0)
    {
      int i, point_count = count / stride;
      arr.SetCapacity(point_count);
      arr.SetCount(point_count);
      ON_4fPoint pt(0.0, 0.0, 0.0, 0.0);
      const float* ptr = vals.Array();
      for (i = 0; i < point_count; i++)
      {
        pt.x = ptr[0];
        pt.y = ptr[1];
        pt.z = ptr[2];
        pt.w = ptr[3];
        arr[i] = pt;
        ptr += stride;
      }
      return arr.Count();
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(CRhinoVariantHelpers::err_array_required);
      return 0;
    }
  }
  else
  {
    arr.SetCapacity(upper - lower + 1);
    VARIANT* pvData = 0;
    hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        ON_4fPoint pt;
        if (ConvertVariant(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoVariantHelpers::SafeArrayToBooleanArray(SAFEARRAY* psa, ON_SimpleArray<bool>& arr)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARTYPE vt = 0;
    HRESULT hr = SafeArrayGetVartype(psa, &vt);
    if (SUCCEEDED(hr))
    {
      if (vt == VT_BOOL)
      {
        VARIANT_BOOL* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(pvData[i] == VARIANT_TRUE ? true : false);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I4)
      {
        long* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(pvData[i] ? true : false);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I2)
      {
        short* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(pvData[i] ? true : false);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R8)
      {
        double* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(pvData[i] ? true : false);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R4)
      {
        float* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(pvData[i] ? true : false);
          SafeArrayUnaccessData(psa);
        }
      }
    }
  }
  return arr.Count();
}

int CRhinoVariantHelpers::SafeArrayToIntegerArray(SAFEARRAY* psa, ON_SimpleArray<int>& arr)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARTYPE vt = 0;
    HRESULT hr = SafeArrayGetVartype(psa, &vt);
    if (SUCCEEDED(hr))
    {
      if (vt == VT_I4)
      {
        long* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((int)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I2)
      {
        short* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((int)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R8)
      {
        double* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((int)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R4)
      {
        float* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((int)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
    }
  }
  return arr.Count();
}

int CRhinoVariantHelpers::SafeArrayToFloatArray(SAFEARRAY* psa, ON_SimpleArray<float>& arr)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARTYPE vt = 0;
    HRESULT hr = SafeArrayGetVartype(psa, &vt);
    if (SUCCEEDED(hr))
    {
      if (vt == VT_R8)
      {
        double* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((float)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R4)
      {
        float* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((float)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I4)
      {
        long* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((float)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I2)
      {
        short* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((float)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
    }
  }
  return arr.Count();
}

int CRhinoVariantHelpers::SafeArrayToDoubleArray(SAFEARRAY* psa, ON_SimpleArray<double>& arr)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARTYPE vt = 0;
    HRESULT hr = SafeArrayGetVartype(psa, &vt);
    if (SUCCEEDED(hr))
    {
      if (vt == VT_R8)
      {
        double* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((double)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R4)
      {
        float* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((double)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I4)
      {
        long* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((double)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I2)
      {
        short* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append((double)pvData[i]);
          SafeArrayUnaccessData(psa);
        }
      }
    }
  }
  return arr.Count();
}

int CRhinoVariantHelpers::SafeArrayToStringArray(SAFEARRAY* psa, ON_ClassArray<ON_wString>& arr)
{
  ASSERT(psa);
  arr.Empty();
  if (psa->fFeatures & FADF_BSTR)
  {
    long i, lower, upper;
    HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
    HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
    if (SUCCEEDED(hl) && SUCCEEDED(hu))
    {
      arr.SetCapacity(upper - lower + 1);
      BSTR* pvData = 0;
      HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
      if (SUCCEEDED(hr) && pvData)
      {
        for (i = lower; i <= upper; i++)
        {
          ON_wString str(pvData[i]);
          if (!str.IsEmpty())
            arr.Append(str);
        }
        SafeArrayUnaccessData(psa);
      }
    }
  }
  return arr.Count();
}

int CRhinoVariantHelpers::SafeArrayToColorArray(SAFEARRAY* psa, ON_SimpleArray<ON_Color>& arr)
{
  ASSERT(psa);
  arr.Empty();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    arr.SetCapacity(upper - lower + 1);
    VARTYPE vt = 0;
    HRESULT hr = SafeArrayGetVartype(psa, &vt);
    if (SUCCEEDED(hr))
    {
      unsigned int black = 0;
      unsigned int white = RGB(255, 255, 255);
      if (vt == VT_BOOL)
      {
        VARIANT_BOOL* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(ON_Color(pvData[i] ? black : white));
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I4)
      {
        long* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(ON_Color(RHINO_CLAMP((unsigned int)pvData[i], black, white)));
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_I2)
      {
        short* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(ON_Color(RHINO_CLAMP((unsigned int)pvData[i], black, white)));
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R8)
      {
        double* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(ON_Color(RHINO_CLAMP((unsigned int)pvData[i], black, white)));
          SafeArrayUnaccessData(psa);
        }
      }
      else if (vt == VT_R4)
      {
        float* pvData = 0;
        hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
        if (SUCCEEDED(hr) && pvData)
        {
          for (i = lower; i <= upper; i++)
            arr.Append(ON_Color(RHINO_CLAMP((unsigned int)pvData[i], black, white)));
          SafeArrayUnaccessData(psa);
        }
      }
    }
  }
  return arr.Count();
}

bool CRhinoVariantHelpers::SafeArrayToPoint(SAFEARRAY* psa, ON_2dPoint& pt)
{
  ASSERT(psa);
  bool rc = false;
  ON_SimpleArray<double> arr;
  int count = SafeArrayToDoubleArray(psa, arr);
  if (count > 1)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::SafeArrayToPoint(SAFEARRAY* psa, ON_3dPoint& pt)
{
  ASSERT(psa);
  bool rc = false;
  ON_SimpleArray<double> arr;
  int count = SafeArrayToDoubleArray(psa, arr);
  if (count == 2)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = 0.0;
    rc = true;
  }
  else if (count > 2)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = arr[2];
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::SafeArrayToPoint(SAFEARRAY* psa, ON_3fPoint& pt)
{
  ASSERT(psa);
  bool rc = false;
  ON_SimpleArray<float> arr;
  int count = SafeArrayToFloatArray(psa, arr);
  if (count == 2)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = 0.0;
    rc = true;
  }
  else if (count > 2)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = arr[2];
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::SafeArrayToPoint(SAFEARRAY* psa, ON_4dPoint& pt)
{
  ASSERT(psa);
  bool rc = false;
  ON_SimpleArray<double> arr;
  int count = SafeArrayToDoubleArray(psa, arr);
  if (count == 3)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = arr[2];
    pt.w = 0.0;
    rc = true;
  }
  else if (count > 3)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = arr[2];
    pt.w = arr[3];
    rc = true;
  }
  return rc;
}

bool CRhinoVariantHelpers::SafeArrayToPoint(SAFEARRAY* psa, ON_4fPoint& pt)
{
  ASSERT(psa);
  bool rc = false;
  ON_SimpleArray<float> arr;
  int count = SafeArrayToFloatArray(psa, arr);
  if (count == 3)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = arr[2];
    pt.w = 0.0;
    rc = true;
  }
  else if (count > 3)
  {
    pt.x = arr[0];
    pt.y = arr[1];
    pt.z = arr[2];
    pt.w = arr[3];
    rc = true;
  }
  return rc;
}

int CRhinoVariantHelpers::SafeArrayToPointArray(SAFEARRAY* psa, ON_2dPointArray& arr)
{
  ASSERT(psa);
  int rc = 0;
  ON_SimpleArray<double> vals;
  int count = SafeArrayToDoubleArray(psa, vals);
  int stride = 2;
  if (count && count % stride == 0)
  {
    int i, point_count = count / stride;
    arr.SetCapacity(point_count);
    arr.SetCount(point_count);
    ON_2dPoint pt(0.0, 0.0);
    const double* ptr = vals.Array();
    for (i = 0; i < point_count; i++)
    {
      pt.x = ptr[0];
      pt.y = ptr[1];
      arr[i] = pt;
      ptr += stride;
    }
    rc = arr.Count();
  }
  return rc;
}

int CRhinoVariantHelpers::SafeArrayToPointArray(SAFEARRAY* psa, ON_3dPointArray& arr)
{
  ASSERT(psa);
  int rc = 0;
  ON_SimpleArray<double> vals;
  int count = SafeArrayToDoubleArray(psa, vals);
  int stride = 3;
  if (count && count % stride == 0)
  {
    int i, point_count = count / stride;
    arr.SetCapacity(point_count);
    arr.SetCount(point_count);
    ON_3dPoint pt(0.0, 0.0, 0.0);
    const double* ptr = vals.Array();
    for (i = 0; i < point_count; i++)
    {
      pt.x = ptr[0];
      pt.y = ptr[1];
      pt.z = ptr[2];
      arr[i] = pt;
      ptr += stride;
    }
    rc = arr.Count();
  }
  return rc;
}

int CRhinoVariantHelpers::SafeArrayToPointArray(SAFEARRAY* psa, ON_3fPointArray& arr)
{
  ASSERT(psa);
  int rc = 0;
  ON_SimpleArray<float> vals;
  int count = SafeArrayToFloatArray(psa, vals);
  int stride = 3;
  if (count && count % stride == 0)
  {
    int i, point_count = count / stride;
    arr.SetCapacity(point_count);
    arr.SetCount(point_count);
    ON_3fPoint pt(0.0, 0.0, 0.0);
    const float* ptr = vals.Array();
    for (i = 0; i < point_count; i++)
    {
      pt.x = ptr[0];
      pt.y = ptr[1];
      pt.z = ptr[2];
      arr[i] = pt;
      ptr += stride;
    }
    rc = arr.Count();
  }
  return rc;
}

int CRhinoVariantHelpers::SafeArrayToPointArray(SAFEARRAY* psa, ON_4dPointArray& arr)
{
  ASSERT(psa);
  int rc = 0;
  ON_SimpleArray<double> vals;
  int count = SafeArrayToDoubleArray(psa, vals);
  int stride = 4;
  if (count && count % stride == 0)
  {
    int i, point_count = count / stride;
    arr.SetCapacity(point_count);
    arr.SetCount(point_count);
    ON_4dPoint pt(0.0, 0.0, 0.0, 0.0);
    const double* ptr = vals.Array();
    for (i = 0; i < point_count; i++)
    {
      pt.x = ptr[0];
      pt.y = ptr[1];
      pt.z = ptr[2];
      pt.w = ptr[3];
      arr[i] = pt;
      ptr += stride;
    }
    rc = arr.Count();
  }
  return rc;
}

int CRhinoVariantHelpers::SafeArrayToPointArray(SAFEARRAY* psa, ON_4fPointArray& arr)
{
  ASSERT(psa);
  int rc = 0;
  ON_SimpleArray<float> vals;
  int count = SafeArrayToFloatArray(psa, vals);
  int stride = 3;
  if (count && count % stride == 0)
  {
    int i, point_count = count / stride;
    arr.SetCapacity(point_count);
    arr.SetCount(point_count);
    ON_4fPoint pt(0.0, 0.0, 0.0, 0.0);
    const float* ptr = vals.Array();
    for (i = 0; i < point_count; i++)
    {
      pt.x = ptr[0];
      pt.y = ptr[1];
      pt.z = ptr[2];
      pt.w = ptr[3];
      arr[i] = pt;
      ptr += stride;
    }
    rc = arr.Count();
  }
  return rc;
}




void CRhinoVariantHelpers::ThrowOleDispatchException(exception_type type)
{
  ON_wString err;
  switch (type)
  {
  case err_boolean_required:
    err = RHSTR(L"Boolean required.");
    break;
  case err_short_required:
    err = RHSTR(L"Short integer required.");
    break;
  case err_long_requried:
    err = RHSTR(L"Long integer required.");
    break;
  case err_integer_required:
    err = RHSTR(L"Integer required.");
    break;
  case err_float_required:
    err = RHSTR(L"Float required.");
    break;
  case err_double_required:
    err = RHSTR(L"Double required.");
    break;
  case err_string_required:
    err = RHSTR(L"String required.");
    break;
  case err_point_required:
    err = RHSTR(L"Point required.");
    break;
  case err_uuid_required:
    err = RHSTR(L"Identifier required.");
    break;
  case err_array_required:
    err = RHSTR(L"Array required.");
    break;
  case err_array_one_dim_required:
    err = RHSTR(L"One-dimensional array required.");
    break;
  case err_array_two_dim_required:
    err = RHSTR(L"Two-dimensional array required.");
    break;
  case err_boolean_array_required:
    err = RHSTR(L"Array of booleans required.");
    break;
  case err_short_array_required:
    err = RHSTR(L"Array of short integers required.");
    break;
  case err_long_array_requried:
    err = RHSTR(L"Array of long integers required.");
    break;
  case err_integer_array_required:
    err = RHSTR(L"Array of integers required.");
    break;
  case err_float_array_required:
    err = RHSTR(L"Array of floats required.");
    break;
  case err_double_array_required:
    err = RHSTR(L"Array of doubles required.");
    break;
  case err_string_array_required:
    err = RHSTR(L"Array of string required.");
    break;
  case err_point_array_required:
    err = RHSTR(L"Array of points required.");
    break;
  case err_uuid_array_required:
    err = RHSTR(L"Array of identifier required.");
    break;
  case err_unknown:
  default:
    err = RHSTR(L"Unknown error.");
    break;
  }

  ON_wString msg;
  msg.Format(RHSTR(L"Type mismatch in parameter. %s"), static_cast<const wchar_t*>(err));

  ::AfxThrowOleDispatchException((WORD)type, static_cast<const wchar_t*>(msg));
}