/////////////////////////////////////////////////////////////////////////////
// RhinoVariantHelpers.h

#pragma once

class CRhinoCom
{
public:
  // Variant validators
  static bool IsVariantEmpty(const VARIANT& va);
  static bool IsVariantNull(const VARIANT& va);
  static bool IsVariantBoolean(const VARIANT& va);
  static bool IsVariantInteger(const VARIANT& va);
  static bool IsVariantDouble(const VARIANT& va);
  static bool IsVariantNumber(const VARIANT& va);
  static bool IsVariantString(const VARIANT& va);
  static bool IsVariantPoint(const VARIANT& va);

  // Variant to object conversons
  static bool VariantToObject(const VARIANT& va, CRhinoObject*& object, bool bQuiet = false);
  static bool VariantToObject(const VARIANT& va, const CRhinoObject*& object, bool bQuiet = false);
  static int VariantToObjects(const VARIANT& va, ON_SimpleArray<CRhinoObject*>& objects, bool bQuiet = false);
  static int VariantToObjects(const VARIANT& va, ON_SimpleArray<const CRhinoObject*>& object, bool bQuiet = false);

  // Variant to UUID conversions
  static bool VariantToUuid(const VARIANT& va, ON_UUID& uuid, bool bQuiet = false);
  static int VariantToUuids(const VARIANT& va, ON_SimpleArray<ON_UUID>& uuids, bool bQuiet = false);

  // Variant to data type conversions
  static bool VariantToBoolean(const VARIANT& va, bool& b, bool bQuiet = false);
  static bool VariantToBoolean(const VARIANT& va, BOOL& b, bool bQuiet = false);
  static bool VariantToInteger(const VARIANT& va, int& n, bool bQuiet = false);
  static bool VariantToFloat(const VARIANT& va, float& f, bool bQuiet = false);
  static bool VariantToDouble(const VARIANT& va, double& d, bool bQuiet = false);
  static bool VariantToString(const VARIANT& va, ON_wString& s, bool bQuiet = false);
  static bool VariantToString(const VARIANT& va, CString& s, bool bQuiet = false);

  // Variant to array conversions
  static int VariantToArray(const VARIANT& va, ON_SimpleArray<bool>& arr, bool bQuiet = false);
  static int VariantToArray(const VARIANT& va, ON_SimpleArray<int>& arr, bool bQuiet = false);
  static int VariantToArray(const VARIANT& va, ON_SimpleArray<float>& arr, bool bQuiet = false);
  static int VariantToArray(const VARIANT& va, ON_SimpleArray<double>& arr, bool bQuiet = false);
  static int VariantToArray(const VARIANT& va, ON_ClassArray<ON_wString>& arr, bool bQuiet = false);
  static int VariantToArray(const VARIANT& va, CStringArray& arr, bool bQuiet = false);

  // Variant to point conversions
  static bool VariantToPoint(const VARIANT& va, ON_2dPoint& pt, bool bQuiet = false);
  static bool VariantToPoint(const VARIANT& va, ON_3dPoint& pt, bool bQuiet = false);
  static bool VariantToPoint(const VARIANT& va, ON_3fPoint& pt, bool bQuiet = false);
  static bool VariantToPoint(const VARIANT& va, ON_4dPoint& pt, bool bQuiet = false);
  static bool VariantToPoint(const VARIANT& va, ON_4fPoint& pt, bool bQuiet = false);
  static bool VariantToVector(const VARIANT& va, ON_3dVector& v, bool bQuiet = false);

  // Variant to point array conversion
  static int VariantToPointArray(const VARIANT& va, ON_2dPointArray& arr, bool bQuiet = false);
  static int VariantToPointArray(const VARIANT& va, ON_3dPointArray& arr, bool bQuiet = false);
  static int VariantToPointArray(const VARIANT& va, ON_3fPointArray& arr, bool bQuiet = false);
  static int VariantToPointArray(const VARIANT& va, ON_4dPointArray& arr, bool bQuiet = false);
  static int VariantToPointArray(const VARIANT& va, ON_4fPointArray& arr, bool bQuiet = false);

  // Miscellaneous variant conversions
  static bool VariantToColor(const VARIANT& va, ON_Color& color, bool bQuiet = false);
  static bool VariantToXform(const VARIANT& va, ON_Xform& xf, bool bQuiet = false);
  static bool VariantToPlane(const VARIANT& va, ON_Plane& plane, bool bQuiet = false);
  static bool VariantToLine(const VARIANT& va, ON_Line& line, bool bQuiet = false);

  // General array to safearray conversion
  static bool BooleanArrayToSafeArray(const ON_SimpleArray<bool>& arr, COleSafeArray& sa);
  static bool BooleanArrayToSafeArray(const ON_SimpleArray<BOOL>& arr, COleSafeArray& sa);
  static bool IntegerArrayToSafeArray(const ON_SimpleArray<int>& arr, COleSafeArray& sa);
  static bool DoubleArrayToSafeArray(const ON_SimpleArray<double>& arr, COleSafeArray& sa);
  static bool StringArrayToSafeArray(const ON_ClassArray<ON_wString>& arr, COleSafeArray& sa, bool bAllowEmptyStrings = true);
  static bool StringArrayToSafeArray(const CStringArray& arr, COleSafeArray& sa, bool bAllowEmptyStrings = true);
  static bool UuidArrayToSafeArray(const ON_SimpleArray<ON_UUID>& arr, COleSafeArray& sa);

  // Point and vector to safearray conversion
  static void PointToSafeArray(const ON_2dPoint& pt, COleSafeArray& sa);
  static void PointToSafeArray(const ON_2fPoint& pt, COleSafeArray& sa);
  static void PointToSafeArray(const ON_3dPoint& pt, COleSafeArray& sa);
  static void PointToSafeArray(const ON_3fPoint& pt, COleSafeArray& sa);
  static void PointToSafeArray(const ON_4dPoint& pt, COleSafeArray& sa);
  static void PointToSafeArray(const ON_4fPoint& pt, COleSafeArray& sa);
  static void VectorToSafeArray(const ON_3dVector& v, COleSafeArray& sa);
  static void VectorToSafeArray(const ON_3fVector& v, COleSafeArray& sa);

  // Point and vector array to safearray conversion
  static bool PointArrayToSafeArray(const ON_2dPointArray& arr, COleSafeArray& sa);
  static bool PointArrayToSafeArray(const ON_2fPointArray& arr, COleSafeArray& sa);
  static bool PointArrayToSafeArray(const ON_3dPointArray& arr, COleSafeArray& sa);
  static bool PointArrayToSafeArray(const ON_4fPointArray& arr, COleSafeArray& sa);
  static bool VectorArrayToSafeArray(const ON_3dVectorArray& arr, COleSafeArray& sa);
  static bool VectorArrayToSafeArray(const ON_3fVectorArray& arr, COleSafeArray& sa);

  // Miscellaneous safearray conversions
  static void XformToSafeArray(const ON_Xform& xf, COleSafeArray& sa);
  static bool PlaneToSafeArray(const ON_Plane& plane, COleSafeArray& sa);
  static void MeshFaceToSafeArray(const ON_MeshFace& mf, COleSafeArray& sa);
  static bool MeshFaceArrayToSafeArray(const ON_SimpleArray<ON_MeshFace>& arr, COleSafeArray& sa);

  // String to UUID conversions
  static bool StringToUUID(const wchar_t* uuid_str, ON_UUID& uuid);
  static CString UUIDToString(const ON_UUID& uuid);
  static bool UUIDToString(const ON_UUID& uuid, CString& uuid_str);
  static bool UUIDToString(const ON_UUID& uuid, ON_wString& uuid_str);
  static int UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, ON_ClassArray<ON_wString>& strings);
  static int UuidArrayToStringArray(const ON_SimpleArray<ON_UUID>& uuids, CStringArray& strings);

  // String to object conversions
  static bool ObjectToString(const CRhinoObject* object, ON_wString& string);
  static bool ObjectToString(const CRhinoObject* object, CString& string);
  static int StringArrayToObjectArray(const ON_ClassArray<ON_wString>& strings, ON_SimpleArray<const CRhinoObject*>& objects);
  static int ObjectArrayToStringArray(const ON_SimpleArray<const CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings);
  static int ObjectArrayToStringArray(const ON_SimpleArray<CRhinoObject*>& objects, ON_ClassArray<ON_wString>& strings);

  // Object helpers
  static CRhinoObject* RhinoObject(const wchar_t* uuid_str);
  static const ON_Curve* RhinoCurve(const wchar_t* uuid_str);
  static const ON_Surface* RhinoSurface(const wchar_t* uuid_str);
  static const ON_BrepFace* RhinoFace(const wchar_t* uuid_str);
  static const ON_Brep* RhinoBrep(const wchar_t* uuid_str);
  static const ON_Mesh* RhinoMesh(const wchar_t* uuid_str);
  static const ON_Point* RhinoPoint(const wchar_t* uuid_str);
  static const ON_PointCloud* RhinoPointCloud(const wchar_t* uuid_str);
  static const ON_Annotation2* RhinoAnnotation(const wchar_t* uuid_str);

protected:
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

  // Low level members to convert safearrays of variants to arrays
  static int VariantArrayToBooleanArray(SAFEARRAY* psa, ON_SimpleArray<bool>& arr, bool bQuiet = false);
  static int VariantArrayToIntegerArray(SAFEARRAY* psa, ON_SimpleArray<int>& arr, bool bQuiet = false);
  static int VariantArrayToFloatArray(SAFEARRAY* psa, ON_SimpleArray<float>& arr, bool bQuiet = false);
  static int VariantArrayToDoubleArray(SAFEARRAY* psa, ON_SimpleArray<double>& arr, bool bQuiet = false);
  static int VariantArrayToStringArray(SAFEARRAY* psa, ON_ClassArray<ON_wString>& arr, bool bQuiet = false);
  static int VariantArrayToStringArray(SAFEARRAY* psa, CStringArray& arr, bool bQuiet = false);

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
  static int SafeArrayToStringArray(SAFEARRAY* psa, CStringArray& arr);

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
};
