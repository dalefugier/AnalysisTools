// Copyright (c) 1993-2016 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisDialogConduit.cpp

#include "StdAfx.h"
#include "AnalysisDialog.h"
#include "AnalysisDialogConduit.h"

CAnalysisDialogConduit::CAnalysisDialogConduit()
  : CRhinoDisplayConduit(CSupportChannels::SC_DRAWOBJECT | CSupportChannels::SC_PREDRAWOBJECTS, false),
  m_dialog(0)
{
}

void CAnalysisDialogConduit::Attach(const CAnalysisDialog* dialog)
{
  m_dialog = dialog;
}

bool CAnalysisDialogConduit::ExecConduit(CRhinoDisplayPipeline& dp, UINT nActiveChannel, bool& bTerminateChannel)
{
  if (nActiveChannel == CSupportChannels::SC_PREDRAWOBJECTS)
  {
    if (m_dialog)
    {
      int i;
      CRhinoViewport& vp = dp.GetRhinoVP();
      if (vp.DisplayMode() == ON::wireframe_display)
      {
        ON_Color color = RhinoApp().AppSettings().AppearanceSettings().SelectedObjectColor();
        for (i = 0; i < m_dialog->m_meshes.Count(); i++)
        {
          ON_Mesh* mesh = m_dialog->m_meshes[i];
          if (mesh)
            dp.DrawWireframeMesh(*mesh, color);
        }
      }
      else
      {
        for (i = 0; i < m_dialog->m_meshes.Count(); i++)
        {
          ON_Mesh* mesh = m_dialog->m_meshes[i];
          if (mesh)
            dp.DrawShadedMesh(*mesh);
        }
      }
      return true;
    }
  }

  if (nActiveChannel == CSupportChannels::SC_DRAWOBJECT)
  {
    const CRhinoObject* obj = m_pChannelAttrs->m_pObject;
    if (obj && m_dialog && m_dialog->m_objects.Search(obj->Attributes().m_uuid) >= 0)
      m_pChannelAttrs->m_bDrawObject = false;
    return true;
  }

  return true;
}
