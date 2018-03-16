// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// RhinoVariantHelpers.h

#pragma once

class CRhinoVariantHelpers
{
public:
  // Variant validators
  static bool IsVariantEmpty(const VARIANT& va);
  static bool IsVariantNull(const VARIANT& va);
  static bool IsVariantNullOrEmpty(const VARIANT& va);
  static bool IsVariantBoolean(const VARIANT& va);
  static bool IsVariantInteger(const VARIANT& va);
  static bool IsVariantDouble(const VARIANT& va);
  static bool IsVariantNumber(const VARIANT& va);
  static bool IsVariantString(const VARIANT& va);
  static bool IsVariantPoint(const VARIANT& va);

public:
  // Variant to Rhino

  // Variant to object reference conversons
  static bool ConvertVariant(const VARIANT& va, CRhinoObjRef& object_ref, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_ClassArray<CRhinoObjRef>& object_refs, bool bQuiet = false);

  // Variant to uuid conversions
  static bool ConvertVariant(const VARIANT& va, ON_UUID& uuid, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_SimpleArray<ON_UUID>& uuids, bool bQuiet = false);

  // Variant to data type conversions
  static bool ConvertVariant(const VARIANT& va, bool& b, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, int& n, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, float& f, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, double& d, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_wString& s, bool bQuiet = false);

  // Variant to array conversions
  static int ConvertVariant(const VARIANT& va, ON_SimpleArray<bool>& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_SimpleArray<int>& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_SimpleArray<float>& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_SimpleArray<double>& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_ClassArray<ON_wString>& arr, bool bQuiet = false);

  // Variant to point conversions
  static bool ConvertVariant(const VARIANT& va, ON_2dPoint& pt, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_3dPoint& pt, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_3fPoint& pt, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_4dPoint& pt, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_4fPoint& pt, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_3dVector& v, bool bQuiet = false);

  // Variant to point array conversion
  static int ConvertVariant(const VARIANT& va, ON_2dPointArray& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_3dPointArray& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_3fPointArray& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_4dPointArray& arr, bool bQuiet = false);
  static int ConvertVariant(const VARIANT& va, ON_4fPointArray& arr, bool bQuiet = false);

  // Miscellaneous variant conversions
  static bool ConvertVariant(const VARIANT& va, ON_Color& color, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_Xform& xf, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_Plane& plane, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_Line& line, bool bQuiet = false);
  static bool ConvertVariant(const VARIANT& va, ON_Interval& interval, bool bQuiet = false);

  // Miscellaneous variant to array conversions
  int ConvertVariant(const VARIANT& va, ON_SimpleArray<ON_Color>& arr, bool bQuiet = false);

public:

  // Rhino to SafeArray

  // SafeArray conversions - points and vectors
  static bool CreateSafeArray(const ON_2dPoint& pt, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_2fPoint& pt, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3dPoint& pt, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3fPoint& pt, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_4dPoint& pt, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_4fPoint& pt, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3dVector& v, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3fVector& v, COleSafeArray& sa);

  // SafeArray conversions - point and vector arrays
  static bool CreateSafeArray(const ON_2dPointArray& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_2fPointArray& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3fPointArray& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3dPointArray& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_4fPointArray& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3dVectorArray& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_3fVectorArray& arr, COleSafeArray& sa);

  // SafeArray conversions - general arrays
  static bool CreateSafeArray(const ON_SimpleArray<bool>& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_SimpleArray<int>& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_SimpleArray<double>& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_ClassArray<ON_wString>& arr, COleSafeArray& sa, bool bAllowEmptyStrings = true);
  static bool CreateSafeArray(const CStringArray& arr, COleSafeArray& sa, bool bAllowEmptyStrings = true);
  static bool CreateSafeArray(const ON_SimpleArray<ON_UUID>& arr, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_ClassArray<CRhinoObjRef>& arr, COleSafeArray& sa);

  // SafeArray conversions - miscellaneous
  static bool CreateSafeArray(const ON_Xform& xf, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_Plane& plane, COleSafeArray& sa);
  static bool CreateSafeArray(const ON_SimpleArray<ON_Plane>& planes, COleSafeArray& sa);

public:

  // Additional helpers
  static CRhinoDoc* Document();
  static void RedrawDocument();
  static void RegenDocument();

  static bool StringToUuid(const wchar_t* uuid_str, ON_UUID& uuid);

  static CString StringFromUuid(const ON_UUID& uuid);

  static bool StringFromUuid(const ON_UUID& uuid, CString& uuid_str);
  static bool StringFromUuid(const ON_UUID& uuid, ON_wString& uuid_str);

  static int UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, ON_ClassArray<ON_wString>& strings);
  static int UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, CStringArray& strings);

  static int ObjectArrayToStringArray(const ON_SimpleArray<const CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings);
  static int ObjectArrayToStringArray(const ON_SimpleArray<CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings);

private:
  static bool RhinoObjRef(const wchar_t* uuid_str, CRhinoObjRef& ref);

  // Low level members to convert safearrays of variants to arrays
  static int VariantArrayToBooleanArray(SAFEARRAY* psa, ON_SimpleArray<bool>& arr, bool bQuiet = false);
  static int VariantArrayToIntegerArray(SAFEARRAY* psa, ON_SimpleArray<int>& arr, bool bQuiet = false);
  static int VariantArrayToFloatArray(SAFEARRAY* psa, ON_SimpleArray<float>& arr, bool bQuiet = false);
  static int VariantArrayToDoubleArray(SAFEARRAY* psa, ON_SimpleArray<double>& arr, bool bQuiet = false);
  static int VariantArrayToStringArray(SAFEARRAY* psa, ON_ClassArray<ON_wString>& arr, bool bQuiet = false);
  static int VariantArrayToColorArray(SAFEARRAY* psa, ON_SimpleArray<ON_Color>& arr, bool bQuiet = false);

  // Low level members to convert safearrays of variants to points
  static bool VariantArrayToPoint(SAFEARRAY* psa, ON_2dPoint& pt, bool bQuiet = false);
  static bool VariantArrayToPoint(SAFEARRAY* psa, ON_3dPoint& pt, bool bQuiet = false);
  static bool VariantArrayToPoint(SAFEARRAY* psa, ON_3fPoint& pt, bool bQuiet = false);
  static bool VariantArrayToPoint(SAFEARRAY* psa, ON_4dPoint& pt, bool bQuiet = false);
  static bool VariantArrayToPoint(SAFEARRAY* psa, ON_4fPoint& pt, bool bQuiet = false);

  // Low level members to convert safearrays of variants to point arrays
  static int VariantArrayToPointArray(SAFEARRAY* psa, ON_2dPointArray& arr, bool bQuiet = false);
  static int VariantArrayToPointArray(SAFEARRAY* psa, ON_3dPointArray& arr, bool bQuiet = false);
  static int VariantArrayToPointArray(SAFEARRAY* psa, ON_3fPointArray& arr, bool bQuiet = false);
  static int VariantArrayToPointArray(SAFEARRAY* psa, ON_4dPointArray& arr, bool bQuiet = false);
  static int VariantArrayToPointArray(SAFEARRAY* psa, ON_4fPointArray& arr, bool bQuiet = false);

  // Low level members to convert safearrays of data to arrays
  static int SafeArrayToBooleanArray(SAFEARRAY* psa, ON_SimpleArray<bool>& arr);
  static int SafeArrayToIntegerArray(SAFEARRAY* psa, ON_SimpleArray<int>& arr);
  static int SafeArrayToFloatArray(SAFEARRAY* psa, ON_SimpleArray<float>& arr);
  static int SafeArrayToDoubleArray(SAFEARRAY* psa, ON_SimpleArray<double>& arr);
  static int SafeArrayToStringArray(SAFEARRAY* psa, ON_ClassArray<ON_wString>& arr);
  static int SafeArrayToColorArray(SAFEARRAY* psa, ON_SimpleArray<ON_Color>& arr);

  // Low level members to convert safearrays of numbers to points
  static bool SafeArrayToPoint(SAFEARRAY* psa, ON_2dPoint& pt);
  static bool SafeArrayToPoint(SAFEARRAY* psa, ON_3dPoint& pt);
  static bool SafeArrayToPoint(SAFEARRAY* psa, ON_3fPoint& pt);
  static bool SafeArrayToPoint(SAFEARRAY* psa, ON_4dPoint& pt);
  static bool SafeArrayToPoint(SAFEARRAY* psa, ON_4fPoint& pt);

  // Low level members to convert safearrays of numbers to point arrays
  static int SafeArrayToPointArray(SAFEARRAY* psa, ON_2dPointArray& arr);
  static int SafeArrayToPointArray(SAFEARRAY* psa, ON_3dPointArray& arr);
  static int SafeArrayToPointArray(SAFEARRAY* psa, ON_3fPointArray& arr);
  static int SafeArrayToPointArray(SAFEARRAY* psa, ON_4dPointArray& arr);
  static int SafeArrayToPointArray(SAFEARRAY* psa, ON_4fPointArray& arr);

  // Exception handling
  enum exception_type
  {
    err_unknown,
    err_boolean_required,
    err_short_required,
    err_long_requried,
    err_integer_required,
    err_float_required,
    err_double_required,
    err_string_required,
    err_point_required,
    err_uuid_required,
    err_array_required,
    err_array_one_dim_required,
    err_array_two_dim_required,
    err_boolean_array_required,
    err_short_array_required,
    err_long_array_requried,
    err_integer_array_required,
    err_float_array_required,
    err_double_array_required,
    err_string_array_required,
    err_point_array_required,
    err_uuid_array_required,
  };
  static void ThrowOleDispatchException(exception_type type);
};
