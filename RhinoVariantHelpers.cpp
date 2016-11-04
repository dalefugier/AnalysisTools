// Copyright (c) 1993-2016 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// RhinoVariantHelpers.cpp

#include "StdAfx.h"
#include "RhinoVariantHelpers.h"

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

void CRhinoCom::ThrowOleDispatchException(exception_type type)
{
  AFX_MANAGE_STATE(AfxGetStaticModuleState());

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
    break;
  }

  ON_wString msg;
  msg.Format(RHSTR(L"Type mismatch in parameter. %s"), err);

  ::AfxThrowOleDispatchException(type, msg);
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

CRhinoObject* CRhinoCom::RhinoObject(const wchar_t* wstr)
{
  CRhinoDoc* doc = RhinoApp().ActiveDoc();
  if (0 == doc)
    return 0;

  ON_wString str(wstr);
  str.TrimLeftAndRight();
  if (str.IsEmpty())
    return 0;

  ON_UUID uuid = ON_UuidFromString(str);

  CRhinoObject* obj = doc->LookupObject(uuid);
  if (obj)
    return obj;

  CRhinoObjectIterator it(CRhinoObjectIterator::undeleted_objects, CRhinoObjectIterator::active_and_reference_objects);
  it.IncludeLights();
  for (obj = it.First(); obj; obj = it.Next())
  {
    if (ON_UuidCompare(&uuid, &obj->Attributes().m_uuid) == 0)
      return obj;
  }

  return 0;
}

const ON_Curve* CRhinoCom::RhinoCurve(const wchar_t* wstr)
{
  const ON_Curve* crv = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    crv = objref.Curve();
  }
  return crv;
}

const ON_Surface* CRhinoCom::RhinoSurface(const wchar_t* wstr)
{
  const ON_Surface* srf = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    srf = objref.Surface();
  }
  return srf;
}

const ON_BrepFace* CRhinoCom::RhinoFace(const wchar_t* wstr)
{
  const ON_BrepFace* face = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    face = objref.Face();
  }
  return face;
}

const ON_Brep* CRhinoCom::RhinoBrep(const wchar_t* wstr)
{
  const ON_Brep* brep = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    brep = objref.Brep();
  }
  return brep;
}

const ON_Mesh* CRhinoCom::RhinoMesh(const wchar_t* wstr)
{
  const ON_Mesh* mesh = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    mesh = objref.Mesh();
  }
  return mesh;
}

const ON_Point* CRhinoCom::RhinoPoint(const wchar_t* wstr)
{
  const ON_Point* point = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    point = objref.Point();
  }
  return point;
}

const ON_PointCloud* CRhinoCom::RhinoPointCloud(const wchar_t* wstr)
{
  const ON_PointCloud* cloud = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    cloud = objref.PointCloud();
  }
  return cloud;
}

const ON_Annotation2* CRhinoCom::RhinoAnnotation(const wchar_t* wstr)
{
  const ON_Annotation2* annotation = 0;
  CRhinoObject* obj = RhinoObject(wstr);
  if (obj)
  {
    CRhinoObjRef objref(obj);
    annotation = objref.Annotation();
  }
  return annotation;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::IsVariantEmpty(const VARIANT& va)
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

bool CRhinoCom::IsVariantNull(const VARIANT& va)
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

bool CRhinoCom::IsVariantBoolean(const VARIANT& va)
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

bool CRhinoCom::IsVariantInteger(const VARIANT& va)
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

bool CRhinoCom::IsVariantDouble(const VARIANT& va)
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

bool CRhinoCom::IsVariantNumber(const VARIANT& va)
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

bool CRhinoCom::IsVariantString(const VARIANT& va)
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

bool CRhinoCom::IsVariantPoint(const VARIANT& va)
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
  else
    return SafeArrayToPoint(psa, pt);

  return false;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::VariantToBoolean(const VARIANT& va, bool& b, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    b = (pva->boolVal ? true : false);
    break;
  case VT_I2:
    b = (pva->iVal ? true : false);
    break;
  case VT_I4:
    b = (pva->lVal ? true : false);
    break;
  case VT_R4:
    b = (pva->fltVal ? true : false);
    break;
  case VT_R8:
    b = (pva->dblVal ? true : false);
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_boolean_required);
  }
  return false;
  }

  return true;
}

bool CRhinoCom::VariantToBoolean(const VARIANT& va, BOOL& b, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    b = (pva->boolVal ? TRUE : FALSE);
    break;
  case VT_I2:
    b = (pva->iVal ? TRUE : FALSE);
    break;
  case VT_I4:
    b = (pva->lVal ? TRUE : FALSE);
    break;
  case VT_R4:
    b = (pva->fltVal ? TRUE : FALSE);
    break;
  case VT_R8:
    b = (pva->dblVal ? TRUE : FALSE);
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_boolean_required);
  }
  return false;
  }

  return true;
}

bool CRhinoCom::VariantToInteger(const VARIANT& va, int& n, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    n = (pva->boolVal ? (int)TRUE : (int)FALSE);
    break;
  case VT_I2:
    n = (int)pva->iVal;
    break;
  case VT_I4:
    n = (int)pva->lVal;
    break;
  case VT_R4:
    n = (int)pva->fltVal;
    break;
  case VT_R8:
    n = (int)pva->dblVal;
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_integer_required);
  }
  return false;
  }

  return true;
}

bool CRhinoCom::VariantToFloat(const VARIANT& va, float& f, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    f = pva->boolVal ? (float)VARIANT_TRUE : (float)VARIANT_FALSE;
    break;
  case VT_I2:
    f = (float)pva->iVal;
    break;
  case VT_I4:
    f = (float)pva->lVal;
    break;
  case VT_R4:
    f = (float)pva->fltVal;
    break;
  case VT_R8:
    f = (float)pva->dblVal;
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_float_required);
  }
  return false;
  }

  return true;
}

bool CRhinoCom::VariantToDouble(const VARIANT& va, double& d, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    d = pva->boolVal ? (double)VARIANT_TRUE : (double)VARIANT_FALSE;
    break;
  case VT_I2:
    d = (double)pva->iVal;
    break;
  case VT_I4:
    d = (double)pva->lVal;
    break;
  case VT_R4:
    d = (double)pva->fltVal;
    break;
  case VT_R8:
    d = (double)pva->dblVal;
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_float_required);
  }
  return false;
  }

  return true;
}

bool CRhinoCom::VariantToString(const VARIANT& va, CString& str, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    str = pva->boolVal == VARIANT_TRUE ? RHSTR(L"True") : RHSTR(L"False");
    break;
  case VT_I2:
    str.Format(L"%d", pva->iVal);
    break;
  case VT_I4:
    str.Format(L"%d", pva->lVal);
    break;
  case VT_R4:
    str.Format(L"%f", pva->fltVal);
    break;
  case VT_R8:
    str.Format(L"%g", pva->dblVal);
    break;
  case VT_BSTR:
    str = pva->bstrVal;
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_string_required);
  }
  return false;
  }

  return true;
}

bool CRhinoCom::VariantToString(const VARIANT& va, ON_wString& str, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  switch (pva->vt)
  {
  case VT_BOOL:
    str = pva->boolVal == VARIANT_TRUE ? RHSTR(L"True") : RHSTR(L"False");
    break;
  case VT_I2:
    str.Format(L"%d", pva->iVal);
    break;
  case VT_I4:
    str.Format(L"%d", pva->lVal);
    break;
  case VT_R4:
    str.Format(L"%f", pva->fltVal);
    break;
  case VT_R8:
    str.Format(L"%g", pva->dblVal);
    break;
  case VT_BSTR:
    str = pva->bstrVal;
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_string_required);
  }
  return false;
  }

  return true;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

int CRhinoCom::VariantToArray(const VARIANT& va, ON_SimpleArray<bool>& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToBooleanArray(psa, arr, bQuiet);
  else
    return SafeArrayToBooleanArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToArray(const VARIANT& va, ON_SimpleArray<int>& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToIntegerArray(psa, arr, bQuiet);
  else
    return SafeArrayToIntegerArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToArray(const VARIANT& va, ON_SimpleArray<float>& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToFloatArray(psa, arr, bQuiet);
  else
    return SafeArrayToFloatArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToArray(const VARIANT& va, ON_SimpleArray<double>& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToDoubleArray(psa, arr, bQuiet);
  else
    return SafeArrayToDoubleArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToArray(const VARIANT& va, ON_ClassArray<ON_wString>& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToStringArray(psa, arr, bQuiet);
  else
    return SafeArrayToStringArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToArray(const VARIANT& va, CStringArray& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToStringArray(psa, arr, bQuiet);
  else
    return SafeArrayToStringArray(psa, arr);

  return 0;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::VariantToPoint(const VARIANT& va, ON_2dPoint& pt, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
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
      ThrowOleDispatchException(err_array_one_dim_required);
  }
  else
    rc = pt.IsValid();

  return rc;
}

bool CRhinoCom::VariantToPoint(const VARIANT& va, ON_3dPoint& pt, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
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
      ThrowOleDispatchException(err_array_one_dim_required);
  }
  else
    rc = pt.IsValid();

  return rc;
}

bool CRhinoCom::VariantToPoint(const VARIANT& va, ON_3fPoint& pt, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
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
      ThrowOleDispatchException(err_array_one_dim_required);
  }

  return rc;
}

bool CRhinoCom::VariantToPoint(const VARIANT& va, ON_4fPoint& pt, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
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
      ThrowOleDispatchException(err_array_one_dim_required);
  }

  return rc;
}

bool CRhinoCom::VariantToPoint(const VARIANT& va, ON_4dPoint& pt, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
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
      ThrowOleDispatchException(err_array_one_dim_required);
  }
  else
    rc = pt.IsValid();

  return rc;
}

bool CRhinoCom::VariantToVector(const VARIANT& va, ON_3dVector& v, bool bQuiet)
{
  ON_3dPoint pt;
  bool rc = VariantToPoint(va, pt, bQuiet);
  if (rc)
    v = pt;

  return rc;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

int CRhinoCom::VariantToPointArray(const VARIANT& va, ON_2dPointArray& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);
  else
    return SafeArrayToPointArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToPointArray(const VARIANT& va, ON_3dPointArray& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);
  else
    return SafeArrayToPointArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToPointArray(const VARIANT& va, ON_3fPointArray& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);
  else
    return SafeArrayToPointArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToPointArray(const VARIANT& va, ON_4dPointArray& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);
  else
    return SafeArrayToPointArray(psa, arr);

  return 0;
}

int CRhinoCom::VariantToPointArray(const VARIANT& va, ON_4fPointArray& arr, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  if (SafeArrayGetDim(psa) != 1)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
    return 0;
  }

  if (psa->fFeatures & FADF_VARIANT)
    return VariantArrayToPointArray(psa, arr, bQuiet);
  else
    return SafeArrayToPointArray(psa, arr);

  return 0;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::VariantToColor(const VARIANT& va, ON_Color& color, bool bQuiet)
{
  if (va.vt == VT_ERROR && va.scode == DISP_E_PARAMNOTFOUND)
    return false;

  const VARIANT* pva = &va;
  while (pva->vt == (VT_BYREF | VT_VARIANT))
    pva = pva->pvarVal;

  int cf = 0;

  switch (pva->vt)
  {
  case VT_BOOL:
    cf = (pva->boolVal) ? 0 : 16777215; // black (off) or white (on)
    break;
  case VT_I2:
    cf = (int)pva->iVal;
    break;
  case VT_I4:
    cf = (int)pva->lVal;
    break;
  case VT_R4:
    cf = (int)pva->fltVal;
    break;
  case VT_R8:
    cf = (int)pva->dblVal;
    break;
  default:
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_integer_required);
  }
  return false;
  }

  if (cf < 0 || cf > RGB(255, 255, 255))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_integer_required);
    return false;
  }

  color = ON_Color(cf);

  return true;
}


bool CRhinoCom::VariantToXform(const VARIANT& va, ON_Xform& xform, bool bQuiet)
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
      ThrowOleDispatchException(err_array_required);
    return false;
  }

  if (SafeArrayGetDim(psa) != 2)
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_one_dim_required);
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
      ThrowOleDispatchException(err_array_two_dim_required);
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
      if (VariantToDouble(element, d, bQuiet))
        matrix[row_count][col_count] = d;

      VariantClear(&element);
    }
  }

  ON_Xform temp(matrix);
  xform = temp;

  return xform.IsValid() ? true : false;
}

bool CRhinoCom::VariantToPlane(const VARIANT& va, ON_Plane& plane, bool bQuiet)
{
  bool rc = false;
  ON_3dPointArray arr;
  if (VariantToPointArray(va, arr, bQuiet) && arr.Count() > 2)
  {
    plane.CreateFromFrame(arr[0], arr[1], arr[2]);
    rc = (plane.IsValid() ? true : false);
  }
  return rc;
}

bool CRhinoCom::VariantToLine(const VARIANT& va, ON_Line& line, bool bQuiet)
{
  bool rc = false;
  ON_3dPointArray arr;
  if (VariantToPointArray(va, arr, bQuiet) && arr.Count() == 2)
  {
    line.Create(arr[0], arr[1]);
    rc = (line.IsValid() ? true : false);
  }
  return rc;
}

bool CRhinoCom::VariantToObject(const VARIANT& va, const CRhinoObject*& object, bool bQuiet)
{
  bool rc = false;
  ON_wString str;
  if (VariantToString(va, str, bQuiet))
  {
    object = RhinoObject(str);
    if (object)
      rc = true;
  }
  return rc;
}

bool CRhinoCom::VariantToObject(const VARIANT& va, CRhinoObject*& object, bool bQuiet)
{
  bool rc = false;
  ON_wString str;
  if (VariantToString(va, str, bQuiet))
  {
    object = RhinoObject(str);
    if (object)
      rc = true;
  }
  return rc;
}

int CRhinoCom::VariantToObjects(const VARIANT& va, ON_SimpleArray<CRhinoObject*>& objects, bool bQuiet)
{
  if (IsVariantString(va))
  {
    CRhinoObject* obj = 0;
    if (VariantToObject(va, obj, bQuiet))
      objects.Append(obj);
  }
  else
  {
    ON_ClassArray<ON_wString> strings;
    if (VariantToArray(va, strings, bQuiet))
    {
      const int string_count = strings.Count();
      int i;
      for (i = 0; i < string_count; i++)
      {
        CRhinoObject* obj = RhinoObject(strings[i]);
        if (obj)
          objects.Append(obj);
      }
    }
  }
  return objects.Count();
}

int CRhinoCom::VariantToObjects(const VARIANT& va, ON_SimpleArray<const CRhinoObject*>& objects, bool bQuiet)
{
  if (IsVariantString(va))
  {
    const CRhinoObject* obj = 0;
    if (VariantToObject(va, obj, bQuiet))
      objects.Append(obj);
  }
  else
  {
    ON_ClassArray<ON_wString> strings;
    if (VariantToArray(va, strings, bQuiet))
    {
      const int string_count = strings.Count();
      int i;
      for (i = 0; i < string_count; i++)
      {
        const CRhinoObject* obj = RhinoObject(strings[i]);
        if (obj)
          objects.Append(obj);
      }
    }
  }
  return objects.Count();
}

bool CRhinoCom::VariantToUuid(const VARIANT& va, ON_UUID& uuid, bool bQuiet)
{
  bool rc = false;
  ON_wString str;
  if (VariantToString(va, str, bQuiet))
    rc = StringToUUID(str, uuid);
  return rc;
}

int CRhinoCom::VariantToUuids(const VARIANT& va, ON_SimpleArray<ON_UUID>& uuids, bool bQuiet)
{
  if (IsVariantString(va))
  {
    ON_UUID uuid = ON_nil_uuid;
    if (VariantToUuid(va, uuid, bQuiet))
      uuids.Append(uuid);
  }
  else
  {
    ON_ClassArray<ON_wString> strings;
    if (VariantToArray(va, strings, bQuiet))
    {
      const int string_count = strings.Count();
      int i;
      for (i = 0; i < string_count; i++)
      {
        ON_UUID uuid = ON_nil_uuid;
        if (StringToUUID(strings[i], uuid))
          uuids.Append(uuid);
      }
    }
  }
  return uuids.Count();
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

void CRhinoCom::PointToSafeArray(const ON_2dPoint& pt, COleSafeArray& sa)
{
  COleVariant var[2];
  var[0] = pt.x;
  var[1] = pt.y;
  sa.CreateOneDim(VT_VARIANT, 2, var);
}

void CRhinoCom::PointToSafeArray(const ON_2fPoint& pt, COleSafeArray& sa)
{
  COleVariant var[2];
  var[0] = pt.x;
  var[1] = pt.y;
  sa.CreateOneDim(VT_VARIANT, 2, var);
}

void CRhinoCom::PointToSafeArray(const ON_3dPoint& pt, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
}

void CRhinoCom::PointToSafeArray(const ON_3fPoint& pt, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
}

void CRhinoCom::PointToSafeArray(const ON_4dPoint& pt, COleSafeArray& sa)
{
  COleVariant var[4];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  var[3] = pt.w;
  sa.CreateOneDim(VT_VARIANT, 4, var);
}

void CRhinoCom::PointToSafeArray(const ON_4fPoint& pt, COleSafeArray& sa)
{
  COleVariant var[4];
  var[0] = pt.x;
  var[1] = pt.y;
  var[2] = pt.z;
  var[3] = pt.w;
  sa.CreateOneDim(VT_VARIANT, 4, var);
}

void CRhinoCom::VectorToSafeArray(const ON_3dVector& v, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = v.x;
  var[1] = v.y;
  var[2] = v.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
}

void CRhinoCom::VectorToSafeArray(const ON_3fVector& v, COleSafeArray& sa)
{
  COleVariant var[3];
  var[0] = v.x;
  var[1] = v.y;
  var[2] = v.z;
  sa.CreateOneDim(VT_VARIANT, 3, var);
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::BooleanArrayToSafeArray(const ON_SimpleArray<bool>& arr, COleSafeArray& sa)
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

bool CRhinoCom::BooleanArrayToSafeArray(const ON_SimpleArray<BOOL>& arr, COleSafeArray& sa)
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

bool CRhinoCom::IntegerArrayToSafeArray(const ON_SimpleArray<int>& arr, COleSafeArray& sa)
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

bool CRhinoCom::DoubleArrayToSafeArray(const ON_SimpleArray<double>& arr, COleSafeArray& sa)
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

bool CRhinoCom::StringArrayToSafeArray(const ON_ClassArray<ON_wString>& arr, COleSafeArray& sa, bool bAllowEmptyStrings)
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

bool CRhinoCom::StringArrayToSafeArray(const CStringArray& arr, COleSafeArray& sa, bool bAllowEmptyStrings)
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

bool CRhinoCom::UuidArrayToSafeArray(const ON_SimpleArray<ON_UUID>& arr, COleSafeArray& sa)
{
  bool rc = false;
  int count = arr.Count();
  if (count > 0)
  {
    CStringArray strings;
    count = UuidArrayToStringArray(arr, strings);
    if (count > 0)
      rc = StringArrayToSafeArray(strings, sa);
  }
  return rc;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::PointArrayToSafeArray(const ON_2dPointArray& arr, COleSafeArray& sa)
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
      PointToSafeArray(arr[i], osa);
      COleVariant va;
      va.Attach(osa.Detach());
      sa.PutElement(index, va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoCom::PointArrayToSafeArray(const ON_2fPointArray& arr, COleSafeArray& sa)
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
      PointToSafeArray(arr[i], osa);
      COleVariant va;
      va.Attach(osa.Detach());
      sa.PutElement(index, va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoCom::PointArrayToSafeArray(const ON_3dPointArray& arr, COleSafeArray& sa)
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
      PointToSafeArray(arr[i], osa);
      COleVariant va;
      va.Attach(osa.Detach());
      sa.PutElement(index, va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoCom::PointArrayToSafeArray(const ON_4fPointArray& arr, COleSafeArray& sa)
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
      PointToSafeArray(arr[i], osa);
      COleVariant va;
      va.Attach(osa.Detach());
      sa.PutElement(index, va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoCom::VectorArrayToSafeArray(const ON_3dVectorArray& arr, COleSafeArray& sa)
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
      VectorToSafeArray(arr[i], osa);
      COleVariant va;
      va.Attach(osa.Detach());
      sa.PutElement(index, va);
    }
    rc = true;
  }
  return rc;
}

bool CRhinoCom::VectorArrayToSafeArray(const ON_3fVectorArray& arr, COleSafeArray& sa)
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
      VectorToSafeArray(arr[i], osa);
      COleVariant va;
      va.Attach(osa.Detach());
      sa.PutElement(index, va);
    }
    rc = true;
  }
  return rc;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

void CRhinoCom::XformToSafeArray(const ON_Xform& xform, COleSafeArray& sa)
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
}

bool CRhinoCom::PlaneToSafeArray(const ON_Plane& plane, COleSafeArray& sa)
{
  ON_3dPointArray arr;
  arr.Append(plane.origin);
  arr.Append(plane.xaxis);
  arr.Append(plane.yaxis);
  arr.Append(plane.zaxis);
  return PointArrayToSafeArray(arr, sa);
}

void CRhinoCom::MeshFaceToSafeArray(const ON_MeshFace& face, COleSafeArray& sa)
{
  COleVariant var[4];
  var[0] = (long)face.vi[0];
  var[1] = (long)face.vi[1];
  var[2] = (long)face.vi[2];
  var[3] = (long)face.vi[3];
  sa.CreateOneDim(VT_VARIANT, 4, var);
}

bool CRhinoCom::MeshFaceArrayToSafeArray(const ON_SimpleArray<ON_MeshFace>& arr, COleSafeArray& sa)
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
      MeshFaceToSafeArray(arr[i], osa);
      COleVariant va;
      va.Attach(osa.Detach());
      sa.PutElement(index, va);
    }
    rc = true;
  }
  return rc;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::StringToUUID(const wchar_t* uuid_str, ON_UUID& uuid)
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

bool CRhinoCom::UUIDToString(const ON_UUID& uuid, ON_wString& uuid_str)
{
  bool rc = false;
  if (ON_UuidToString(uuid, uuid_str))
  {
    uuid_str.MakeUpper();
    rc = true;
  }
  return rc;
}

bool CRhinoCom::UUIDToString(const ON_UUID& uuid, CString& uuid_str)
{
  bool rc = false;
  ON_wString str;
  if (UUIDToString(uuid, str))
  {
    uuid_str = str;
    uuid_str.MakeUpper();
    rc = true;
  }
  return rc;
}

CString CRhinoCom::UUIDToString(const ON_UUID& uuid)
{
  CString str;
  UUIDToString(uuid, str);
  str.MakeUpper();
  return str;
}

int CRhinoCom::UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, ON_ClassArray<ON_wString>& strings)
{
  int i;
  for (i = 0; i < uuids.Count(); i++)
  {
    ON_wString uuid_str;
    if (UUIDToString(uuids[i], uuid_str))
      strings.Append(uuid_str);
  }
  return strings.Count();
}

int CRhinoCom::UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, CStringArray& strings)
{
  int i;
  for (i = 0; i < uuids.Count(); i++)
  {
    CString uuid_str;
    if (UUIDToString(uuids[i], uuid_str))
      strings.Add(uuid_str);
  }
  return (int)strings.GetCount();
}

bool CRhinoCom::ObjectToString(const CRhinoObject* object, ON_wString& string)
{
  bool rc = false;
  if (object)
  {
    if (ON_UuidToString(object->Attributes().m_uuid, string))
      rc = true;
  }
  return rc;
}

bool CRhinoCom::ObjectToString(const CRhinoObject* object, CString& string)
{
  bool rc = false;
  ON_wString str;
  if (ObjectToString(object, str))
  {
    string = str;
    rc = true;
  }
  return rc;
}

int CRhinoCom::StringArrayToObjectArray(const ON_ClassArray<ON_wString>& strings, ON_SimpleArray<const CRhinoObject*>& objects)
{
  int i, num_added = 0;
  for (i = 0; i < strings.Count(); i++)
  {
    const CRhinoObject* obj = RhinoObject(strings[i]);
    if (obj)
    {
      objects.Append(obj);
      num_added++;
    }
  }
  return num_added;
}

int CRhinoCom::ObjectArrayToStringArray(const ON_SimpleArray<const CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings)
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

int CRhinoCom::ObjectArrayToStringArray(const ON_SimpleArray<CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings)
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

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

int CRhinoCom::VariantArrayToBooleanArray(SAFEARRAY* psa, ON_SimpleArray<bool>& arr, bool bQuiet)
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
        if (VariantToBoolean(pvData[i], b, bQuiet))
          arr.Append(b);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToIntegerArray(SAFEARRAY* psa, ON_SimpleArray<int>& arr, bool bQuiet)
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
        if (VariantToInteger(pvData[i], n, bQuiet))
          arr.Append(n);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToFloatArray(SAFEARRAY* psa, ON_SimpleArray<float>& arr, bool bQuiet)
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
        if (VariantToFloat(pvData[i], f, bQuiet))
          arr.Append(f);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToDoubleArray(SAFEARRAY* psa, ON_SimpleArray<double>& arr, bool bQuiet)
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
        if (VariantToDouble(pvData[i], d, bQuiet))
          arr.Append(d);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToStringArray(SAFEARRAY* psa, ON_ClassArray<ON_wString>& arr, bool bQuiet)
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
        if (VariantToString(pvData[i], str, bQuiet))
          arr.Append(str);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToStringArray(SAFEARRAY* psa, CStringArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.RemoveAll();
  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (SUCCEEDED(hl) && SUCCEEDED(hu))
  {
    VARIANT* pvData = 0;
    HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
    if (SUCCEEDED(hr) && pvData)
    {
      for (i = lower; i <= upper; i++)
      {
        CString str;
        if (VariantToString(pvData[i], str, bQuiet))
          arr.Add(str);
      }
      SafeArrayUnaccessData(psa);
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return (int)arr.GetCount(); // (int) cast for x64 builds
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::VariantArrayToPoint(SAFEARRAY* psa, ON_2dPoint& pt, bool bQuiet)
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
        VariantToDouble(pvData[i], p[i], bQuiet);
      SafeArrayUnaccessData(psa);
      pt.x = p[0];
      pt.y = p[1];
      rc = true;
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return rc;
}

bool CRhinoCom::VariantArrayToPoint(SAFEARRAY* psa, ON_3dPoint& pt, bool bQuiet)
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
        VariantToDouble(pvData[i], p[i], bQuiet);
      SafeArrayUnaccessData(psa);
      pt.x = p[0];
      pt.y = p[1];
      pt.z = p[2];
      rc = true;
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return rc;
}

bool CRhinoCom::VariantArrayToPoint(SAFEARRAY* psa, ON_3fPoint& pt, bool bQuiet)
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
        VariantToFloat(pvData[i], p[i], bQuiet);
      SafeArrayUnaccessData(psa);
      pt.x = p[0];
      pt.y = p[1];
      pt.z = p[2];
      rc = true;
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return rc;
}

bool CRhinoCom::VariantArrayToPoint(SAFEARRAY* psa, ON_4dPoint& pt, bool bQuiet)
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
        VariantToDouble(pvData[i], p[i], bQuiet);
      SafeArrayUnaccessData(psa);
      pt.x = p[0];
      pt.y = p[1];
      pt.z = p[2];
      pt.w = p[3];
      rc = true;
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return rc;
}

bool CRhinoCom::VariantArrayToPoint(SAFEARRAY* psa, ON_4fPoint& pt, bool bQuiet)
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
        VariantToFloat(pvData[i], p[i], bQuiet);
      SafeArrayUnaccessData(psa);
      pt.x = p[0];
      pt.y = p[1];
      pt.z = p[2];
      pt.w = p[3];
      rc = true;
    }
    else
    {
      if (false == bQuiet)
        ThrowOleDispatchException(err_array_required);
    }
  }
  else
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
  }

  return rc;
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

int CRhinoCom::VariantArrayToPointArray(SAFEARRAY* psa, ON_2dPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 2;
    ON_SimpleArray<double> vals;
    int count = VariantToArray(va, vals, bQuiet);
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
        ThrowOleDispatchException(err_array_required);
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
        if (VariantToPoint(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToPointArray(SAFEARRAY* psa, ON_3dPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 3;
    ON_SimpleArray<double> vals;
    int count = VariantToArray(va, vals, bQuiet);
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
        ThrowOleDispatchException(err_array_required);
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
        if (VariantToPoint(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToPointArray(SAFEARRAY* psa, ON_3fPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 3;
    ON_SimpleArray<float> vals;
    int count = VariantToArray(va, vals, bQuiet);
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
        ThrowOleDispatchException(err_array_required);
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
        if (VariantToPoint(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToPointArray(SAFEARRAY* psa, ON_4dPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 4;
    ON_SimpleArray<double> vals;
    int count = VariantToArray(va, vals, bQuiet);
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
        ThrowOleDispatchException(err_array_required);
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
        if (VariantToPoint(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

int CRhinoCom::VariantArrayToPointArray(SAFEARRAY* psa, ON_4fPointArray& arr, bool bQuiet)
{
  ASSERT(psa);
  arr.Empty();

  long i, lower, upper;
  HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
  HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
  if (FAILED(hl) || FAILED(hu))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  VARIANT va;
  VariantInit(&va);
  HRESULT hr = SafeArrayGetElement(psa, &lower, &va);
  if (FAILED(hr))
  {
    if (false == bQuiet)
      ThrowOleDispatchException(err_array_required);
    return 0;
  }

  bool bIsNumber = IsVariantNumber(va);
  VariantClear(&va);
  if (bIsNumber)
  {
    int stride = 4;
    ON_SimpleArray<float> vals;
    int count = VariantToArray(va, vals, bQuiet);
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
        ThrowOleDispatchException(err_array_required);
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
        if (VariantToPoint(pvData[i], pt, bQuiet))
          arr.Append(pt);
      }
      SafeArrayUnaccessData(psa);
    }
  }

  return arr.Count();
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

int CRhinoCom::SafeArrayToBooleanArray(SAFEARRAY* psa, ON_SimpleArray<bool>& arr)
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

int CRhinoCom::SafeArrayToIntegerArray(SAFEARRAY* psa, ON_SimpleArray<int>& arr)
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

int CRhinoCom::SafeArrayToFloatArray(SAFEARRAY* psa, ON_SimpleArray<float>& arr)
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

int CRhinoCom::SafeArrayToDoubleArray(SAFEARRAY* psa, ON_SimpleArray<double>& arr)
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

int CRhinoCom::SafeArrayToStringArray(SAFEARRAY* psa, ON_ClassArray<ON_wString>& arr)
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

int CRhinoCom::SafeArrayToStringArray(SAFEARRAY* psa, CStringArray& arr)
{
  ASSERT(psa);
  arr.RemoveAll();
  if (psa->fFeatures & FADF_BSTR)
  {
    long i, lower, upper;
    HRESULT hl = SafeArrayGetLBound(psa, 1, &lower);
    HRESULT hu = SafeArrayGetUBound(psa, 1, &upper);
    if (SUCCEEDED(hl) && SUCCEEDED(hu))
    {
      BSTR* pvData = 0;
      HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&pvData);
      if (SUCCEEDED(hr) && pvData)
      {
        for (i = lower; i <= upper; i++)
        {
          CString str(pvData[i]);
          if (!str.IsEmpty())
            arr.Add(str);
        }
        SafeArrayUnaccessData(psa);
      }
    }
  }
  return (int)arr.GetCount(); // (int) cast for x64 builds
}

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

bool CRhinoCom::SafeArrayToPoint(SAFEARRAY* psa, ON_2dPoint& pt)
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

bool CRhinoCom::SafeArrayToPoint(SAFEARRAY* psa, ON_3dPoint& pt)
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

bool CRhinoCom::SafeArrayToPoint(SAFEARRAY* psa, ON_3fPoint& pt)
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

bool CRhinoCom::SafeArrayToPoint(SAFEARRAY* psa, ON_4dPoint& pt)
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

bool CRhinoCom::SafeArrayToPoint(SAFEARRAY* psa, ON_4fPoint& pt)
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

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////

int CRhinoCom::SafeArrayToPointArray(SAFEARRAY* psa, ON_2dPointArray& arr)
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

int CRhinoCom::SafeArrayToPointArray(SAFEARRAY* psa, ON_3dPointArray& arr)
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

int CRhinoCom::SafeArrayToPointArray(SAFEARRAY* psa, ON_3fPointArray& arr)
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

int CRhinoCom::SafeArrayToPointArray(SAFEARRAY* psa, ON_4dPointArray& arr)
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

int CRhinoCom::SafeArrayToPointArray(SAFEARRAY* psa, ON_4fPointArray& arr)
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
