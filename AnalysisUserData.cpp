// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisUserData.cpp

#include "stdafx.h"
#include "AnalysisUserData.h"
#include "AnalysisToolsPlugIn.h"

ON_OBJECT_IMPLEMENT(CAnalysisUserData, ON_UserData, "E661F7EE-E478-41e4-9EE1-50FA72AE123D");

CAnalysisUserData::~CAnalysisUserData()
{
}

ON_UUID CAnalysisUserData::Id()
{
  return CAnalysisUserData::m_CAnalysisUserData_class_rtti.Uuid();
}

const CAnalysisUserData* CAnalysisUserData::Get(const ON_Mesh* mesh)
{
  const CAnalysisUserData* ud = nullptr;
  if (mesh)
    ud = CAnalysisUserData::Cast(mesh->GetUserData(CAnalysisUserData::Id()));
  return ud;
}

ON_Color CAnalysisUserData::Color(double a) const
{
  ON_Color c;
  double s = 0.0;
  if (m_redblue[0] == m_redblue[1])
  {
    s = m_redblue[0];
    if (a < s)
      c.SetRGB(255, 0, 0);
    else if (a > s)
      c.SetRGB(0, 0, 255);
    else
      c.SetRGB(0, 255, 0);
  }
  else
  {
    s = m_redblue.NormalizedParameterAt(a);
    if (s < 0.0)
      s = 0.0;
    else if (s > 1.0)
      s = 1.0;
    c.SetHSV(s * 4.0 * ON_PI / 3.0, 1.0, 1.0);
  }
  return c;
}

bool CAnalysisUserData::UpdateColors(ON_Mesh* mesh)
{
  bool rc = false;
  const CAnalysisUserData* ud = CAnalysisUserData::Get(mesh);
  if (ud)
  {
    ON_Color c;
    int i;
    const int vcount = ud->m_a.Count();
    if (vcount == mesh->m_V.Count())
    {
      rc = true;
      mesh->m_C.SetCount(0);
      mesh->m_C.Reserve(vcount);
      for (i = 0; i < vcount; i++)
      {
        c = ud->Color(ud->m_a[i]);
        mesh->m_C.Append(c);
      }
    }
  }
  return rc;
}

CAnalysisUserData::CAnalysisUserData()
{
  m_userdata_uuid = CAnalysisUserData::Id();
  m_application_uuid = AnalysisToolsPlugIn().PlugInID();

  m_userdata_copycount = 1;
}

CAnalysisUserData::CAnalysisUserData(const CAnalysisUserData& src)
  : ON_UserData(src)
{
  m_userdata_uuid = CAnalysisUserData::Id();
  m_application_uuid = AnalysisToolsPlugIn().PlugInID();

  m_a = src.m_a;
  m_minmax = src.m_minmax;
  m_redblue = src.m_redblue;
}

CAnalysisUserData& CAnalysisUserData::operator=(const CAnalysisUserData& src)
{
  if (this != &src)
  {
    ON_UUID saved_uuid = m_userdata_uuid;
    ON_UserData::operator=(src);
    m_userdata_uuid = saved_uuid;

    m_a = src.m_a;
    m_minmax = src.m_minmax;
    m_redblue = src.m_redblue;
  }
  return *this;
}

bool CAnalysisUserData::GetDescription(ON_wString& description)
{
  description = RHSTR(L"Analysis Tools data");
  return true;
}

bool CAnalysisUserData::Archive() const
{
  return true;
}

bool CAnalysisUserData::Write(ON_BinaryArchive& archive) const
{
  int major_version = 1;
  int minor_version = 0;

  bool rc = archive.BeginWrite3dmChunk(TCODE_ANONYMOUS_CHUNK, major_version, minor_version);
  if (!rc)
    return false;

  // Write class members
  for (;;)
  {
    // version 1.0 fields

    rc = archive.WriteArray(m_a);
    if (!rc) break;

    rc = archive.WriteInterval(m_minmax);
    if (!rc) break;

    rc = archive.WriteInterval(m_redblue);
    if (!rc) break;

    break;
  }

  // If BeginWrite3dmChunk() returns true, then EndWrite3dmChunk()
  // must be called, even if a write operation failed.
  if (!archive.EndWrite3dmChunk())
    rc = false;

  return rc;
}

bool CAnalysisUserData::Read(ON_BinaryArchive& archive)
{
  m_a.SetCount(0);
  m_minmax.Destroy();
  m_redblue.Destroy();

  int major_version = 0;
  int minor_version = 0;

  bool rc = archive.BeginRead3dmChunk(TCODE_ANONYMOUS_CHUNK, &major_version, &minor_version);
  if (!rc)
    return false;

  // Read class members
  for (;;)
  {
    rc = (1 == major_version);
    if (!rc) break;

    // version 1.0 fields

    rc = archive.ReadArray(m_a);
    if (!rc) break;

    rc = archive.ReadInterval(m_minmax);
    if (!rc) break;

    rc = archive.ReadInterval(m_redblue);
    if (!rc) break;

    break;
  }

  // If BeginRead3dmChunk() returns true, then EndRead3dmChunk()
  // must be called, even if a read operation failed.
  if (!archive.EndRead3dmChunk())
    rc = false;

  return rc;
}