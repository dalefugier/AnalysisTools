// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisDialogConduit.cpp

#include "StdAfx.h"
#include "AnalysisDialog.h"
#include "AnalysisDialogConduit.h"

CAnalysisDialogConduit::CAnalysisDialogConduit()
  : CRhinoDisplayConduit(CSupportChannels::SC_DRAWOBJECT | CSupportChannels::SC_PREDRAWOBJECTS, false),
  m_dialog(nullptr)
{
  m_color = RhinoApp().AppSettings().AppearanceSettings().SelectedObjectColor();
}

void CAnalysisDialogConduit::Attach(const CAnalysisDialog* dialog)
{
  m_dialog = dialog;
}

bool CAnalysisDialogConduit::ExecConduit(CRhinoDisplayPipeline& dp, UINT nActiveChannel, bool& bTerminateChannel)
{
  if (nActiveChannel == CSupportChannels::SC_PREDRAWOBJECTS)
  {
    if (nullptr != m_dialog)
    {
      if (!dp.DisplayAttrs()->m_bShadeSurface)
      {
        for (int i = 0; i < m_dialog->m_meshes.Count(); i++)
        {
          ON_Mesh* mesh = m_dialog->m_meshes[i];
          if (nullptr != mesh)
            dp.DrawWireframeMesh(*mesh, m_color);
        }
      }
      else
      {
        for (int i = 0; i < m_dialog->m_meshes.Count(); i++)
        {
          ON_Mesh* mesh = m_dialog->m_meshes[i];
          if (nullptr != mesh)
            dp.DrawShadedMesh(*mesh);
        }
      }
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
