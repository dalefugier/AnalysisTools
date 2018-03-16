// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisObject.cpp

#include "StdAfx.h"
#include "AnalysisObject.h"
#include "AnalysisUserData.h"
#include "RhinoVariantHelpers.h"

// CAnalysisObject

IMPLEMENT_DYNAMIC(CAnalysisObject, CCmdTarget)

CAnalysisObject::CAnalysisObject()
{
  EnableAutomation();
}

CAnalysisObject::~CAnalysisObject()
{
}

void CAnalysisObject::OnFinalRelease()
{
  CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CAnalysisObject, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CAnalysisObject, CCmdTarget)
  DISP_FUNCTION_ID(CAnalysisObject, "IsAnalysisMesh", dispidIsAnalysisMesh, IsAnalysisMesh, VT_VARIANT, VTS_VARIANT)
  DISP_FUNCTION_ID(CAnalysisObject, "AddAnalysisMesh", dispidAddAnalysisMesh, AddAnalysisMesh, VT_VARIANT, VTS_VARIANT VTS_VARIANT VTS_VARIANT)
  DISP_FUNCTION_ID(CAnalysisObject, "AnalysisMeshData", dispidAnalysisMeshData, AnalysisMeshData, VT_VARIANT, VTS_VARIANT VTS_VARIANT)
  DISP_FUNCTION_ID(CAnalysisObject, "AnalysisMeshDisplayRange", dispidAnalysisMeshDisplayRange, AnalysisMeshDisplayRange, VT_VARIANT, VTS_VARIANT VTS_VARIANT)
  DISP_FUNCTION_ID(CAnalysisObject, "AnalysisMeshDataRange", dispidAnalysisMeshDataRange, AnalysisMeshDataRange, VT_VARIANT, VTS_VARIANT)
END_DISPATCH_MAP()

// Note: we add support for IID_IAnalysisObject to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .IDL file.

// {EA4DEC31-FC91-4FB7-8819-0DCC67CB5E3A}
static const IID IID_IAnalysisObject =
{ 0xEA4DEC31, 0xFC91, 0x4FB7, { 0x88, 0x19, 0xD, 0xCC, 0x67, 0xCB, 0x5E, 0x3A } };

BEGIN_INTERFACE_MAP(CAnalysisObject, CCmdTarget)
  INTERFACE_PART(CAnalysisObject, IID_IAnalysisObject, Dispatch)
END_INTERFACE_MAP()

// CAnalysisObject message handlers

VARIANT CAnalysisObject::IsAnalysisMesh(const VARIANT &vaObject)
{
  VARIANT vaResult;
  VariantInit(&vaResult);
  V_VT(&vaResult) = VT_NULL;

  CRhinoObjRef object_ref;
  if (!CRhinoVariantHelpers::ConvertVariant(vaObject, object_ref))
    return vaResult;

  const ON_Mesh* mesh = object_ref.Mesh();
  if (nullptr == mesh)
    return vaResult;

  V_VT(&vaResult) = VT_BOOL;
  const CAnalysisUserData* ud = CAnalysisUserData::Get(mesh);
  vaResult.boolVal = (nullptr != ud) ? VARIANT_TRUE : VARIANT_FALSE;

  return vaResult;
}

VARIANT CAnalysisObject::AddAnalysisMesh(const VARIANT& vaVertices, const VARIANT& vaFaces, const VARIANT& vaData)
{
  VARIANT vaResult;
  VariantInit(&vaResult);
  V_VT(&vaResult) = VT_NULL;

  CRhinoDoc* doc = CRhinoVariantHelpers::Document();
  if (nullptr == doc)
    return vaResult;

  ON_3fPointArray vertices;
  if (!CRhinoVariantHelpers::ConvertVariant(vaVertices, vertices))
    return vaResult;

  ON_4fPointArray faces;
  if (!CRhinoVariantHelpers::ConvertVariant(vaFaces, faces))
    return vaResult;

  ON_SimpleArray<double> data;
  if (!CRhinoVariantHelpers::ConvertVariant(vaData, data) || data.Count() != vertices.Count())
    return vaResult;

  ON_Mesh* mesh = new ON_Mesh(faces.Count(), vertices.Count(), false, false);

  for (int i = 0; i < vertices.Count(); i++)
    mesh->SetVertex(i, vertices[i]);

  for (int i = 0; i < faces.Count(); i++)
  {
    ON_4fPoint face = faces[i];
    if (face.z == face.w)
      mesh->SetTriangle(i, (int)face.x, (int)face.y, (int)face.z);
    else
      mesh->SetQuad(i, (int)face.x, (int)face.y, (int)face.z, (int)face.w);
  }

  mesh->ComputeVertexNormals();
  mesh->Compact();

  if (mesh->IsValid())
  {
    CAnalysisUserData* ud = new CAnalysisUserData();
    if (ud)
    {
      ud->m_a = data;
      double mn = data[0];
      double mx = data[0];
      for (int i = 1; i < data.Count(); i++)
      {
        double x = data[i];
        if (x < mn)
          mn = x;
        else if (x > mx)
          mx = x;
      }
      ud->m_minmax.Set(mn, mx);
      ud->m_redblue.Set(mn, mx);
      mesh->AttachUserData(ud);
      CAnalysisUserData::UpdateColors(mesh);

      CRhinoMeshObject* mesh_object = new CRhinoMeshObject();
      if (mesh_object)
      {
        mesh_object->SetMesh(mesh);
        if (doc->AddObject(mesh_object))
        {
          CString str = CRhinoVariantHelpers::StringFromUuid(mesh_object->ModelObjectId());
          V_VT(&vaResult) = VT_BSTR;
          vaResult.bstrVal = str.AllocSysString();
          doc->Regen();
        }
        else
          delete mesh_object;
      }
      else
        delete mesh;
    }
    else
      delete mesh;
  }

  return vaResult;
}

VARIANT CAnalysisObject::AnalysisMeshData(const VARIANT& vaObject, const VARIANT& vaData)
{
  VARIANT vaResult;
  VariantInit(&vaResult);
  V_VT(&vaResult) = VT_NULL;

  CRhinoObjRef object_ref;
  if (!CRhinoVariantHelpers::ConvertVariant(vaObject, object_ref))
    return vaResult;

  ON_Mesh* mesh = const_cast<ON_Mesh*>(object_ref.Mesh());
  if (nullptr == mesh)
    return vaResult;

  ON_SimpleArray<double> old_data;
  CAnalysisUserData* ud = const_cast<CAnalysisUserData*>(CAnalysisUserData::Get(mesh));
  if (nullptr != ud)
    old_data = ud->m_a;

  ON_SimpleArray<double> new_data;
  if (CRhinoVariantHelpers::ConvertVariant(vaData, new_data) && new_data.Count() == mesh->VertexCount())
  {
    bool bAttach = false;
    if (nullptr == ud)
    {
      ud = new CAnalysisUserData();
      bAttach = true;
    }

    if (ud)
    {
      ud->m_a = new_data;
      double mn = new_data[0];
      double mx = new_data[0];
      for (int i = 1; i < new_data.Count(); i++)
      {
        double x = new_data[i];
        if (x < mn)
          mn = x;
        else if (x > mx)
          mx = x;
      }
      ud->m_minmax.Set(mn, mx);
      ud->m_redblue.Set(mn, mx);

      if (bAttach)
        mesh->AttachUserData(ud);

      CAnalysisUserData::UpdateColors(mesh);

      CRhinoVariantHelpers::RegenDocument();
    }
  }

  if (old_data.Count())
  {
    COleSafeArray sa;
    if (CRhinoVariantHelpers::CreateSafeArray(old_data, sa))
      return sa.Detach();
  }

  return vaResult;
}

VARIANT CAnalysisObject::AnalysisMeshDisplayRange(const VARIANT& vaObject, const VARIANT& vaRange)
{
  VARIANT vaResult;
  VariantInit(&vaResult);
  V_VT(&vaResult) = VT_NULL;

  CRhinoObjRef object_ref;
  if (!CRhinoVariantHelpers::ConvertVariant(vaObject, object_ref))
    return vaResult;

  ON_Mesh* mesh = const_cast<ON_Mesh*>(object_ref.Mesh());
  if (nullptr == mesh)
    return vaResult;

  CAnalysisUserData* ud = const_cast<CAnalysisUserData*>(CAnalysisUserData::Get(mesh));
  if (nullptr == ud)
    return vaResult;

  ON_2dPoint old_range(ud->m_redblue[0], ud->m_redblue[1]);

  ON_2dPoint new_range;
  if (CRhinoVariantHelpers::ConvertVariant(vaRange, new_range) && new_range != old_range)
  {
    ud->m_redblue.Set(new_range.x, new_range.y);
    CAnalysisUserData::UpdateColors(mesh);
    CRhinoVariantHelpers::RegenDocument();
  }

  COleSafeArray sa;
  CRhinoVariantHelpers::CreateSafeArray(old_range, sa);

  return sa.Detach();
}

VARIANT CAnalysisObject::AnalysisMeshDataRange(const VARIANT& vaObject)
{
  VARIANT vaResult;
  VariantInit(&vaResult);
  V_VT(&vaResult) = VT_NULL;

  CRhinoObjRef object_ref;
  if (!CRhinoVariantHelpers::ConvertVariant(vaObject, object_ref))
    return vaResult;

  const ON_Mesh* mesh = object_ref.Mesh();
  if (nullptr == mesh)
    return vaResult;

  const CAnalysisUserData* ud = CAnalysisUserData::Get(mesh);
  if (nullptr == ud)
    return vaResult;

  ON_2dPoint range(ud->m_minmax[0], ud->m_minmax[1]);
  
  COleSafeArray sa;
  CRhinoVariantHelpers::CreateSafeArray(range, sa);

  return sa.Detach();
}
