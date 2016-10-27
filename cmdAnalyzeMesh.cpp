/////////////////////////////////////////////////////////////////////////////
// cmdAnalyzeMesh.cpp

#include "StdAfx.h"
#include "AnalysisDialog.h"
#include "AnalysisUserData.h"

/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
//
// BEGIN AnalyzeMesh command
//

#pragma region AnalyzeMesh command

class CAnalysisMeshPicker : public CRhinoGetObject
{
public:
  bool CustomGeometryFilter(
    const CRhinoObject* object,
    const ON_Geometry* geometry,
    ON_COMPONENT_INDEX component_index
  )
    const
  {
    bool rc = false;
    if (object)
    {
      const CRhinoMeshObject* mesh_object = CRhinoMeshObject::Cast(object);
      if (mesh_object && CAnalysisUserData::Get(mesh_object->Mesh()))
        rc = true;
    }
    return rc;
  }
};

/////////////////////////////////////////////////////////////////////////////

class CCommandAnalyzeMesh : public CRhinoCommand
{
public:
  CCommandAnalyzeMesh() {}
  ~CCommandAnalyzeMesh() {}
  UUID CommandUUID()
  {
    // {F76BF44E-B863-401A-AA80-887F8B81FAC6}
    static const GUID AnalyzeMeshCommand_UUID =
    { 0xF76BF44E, 0xB863, 0x401A, { 0xAA, 0x80, 0x88, 0x7F, 0x8B, 0x81, 0xFA, 0xC6 } };
    return AnalyzeMeshCommand_UUID;
  }
  const wchar_t* EnglishCommandName() { return L"AnalyzeMesh"; }
  const wchar_t* LocalCommandName() { return L"AnalyzeMesh"; }
  CRhinoCommand::result RunCommand(const CRhinoCommandContext&);

private:
  CRhinoCommand::result GetDialogParameters(
    const ON_SimpleArray<const CRhinoMeshObject*>& mesh_objects,
    const ON_Interval& minmax,
    const ON_Interval& old_redblue,
    ON_Interval& redblue
  );

  CRhinoCommand::result GetScriptParameters(
    const ON_SimpleArray<const CRhinoMeshObject*>& mesh_objects,
    const ON_Interval& minmax,
    const ON_Interval& old_redblue,
    ON_Interval& redblue
  );
};

// The one and only CCommandAnalyzeMesh object
static class CCommandAnalyzeMesh theAnalyzeMeshCommand;

CRhinoCommand::result CCommandAnalyzeMesh::RunCommand(const CRhinoCommandContext& context)
{
  CAnalysisMeshPicker go;
  go.SetCommandPrompt(RHSTR(L"Select analysis meshes"));
  go.SetGeometryFilter(CRhinoGetObject::mesh_object);
  go.GetObjects(1, 0);
  if (go.CommandResult() != success)
    return go.CommandResult();

  int i, ud_count = 0;
  ON_Interval minmax, redblue;

  ON_SimpleArray<const CRhinoMeshObject*> mesh_objects(go.ObjectCount());
  for (i = 0; i < go.ObjectCount(); i++)
  {
    const CRhinoMeshObject* mesh_object = CRhinoMeshObject::Cast(go.Object(i).Object());
    if (mesh_object)
    {
      const CAnalysisUserData* ud = CAnalysisUserData::Get(mesh_object->Mesh());
      if (ud)
      {
        if (0 == ud_count)
        {
          minmax = ud->m_minmax;
          redblue[0] = ud->m_redblue[0];
          redblue[1] = ud->m_redblue[1];
        }
        else
        {
          minmax.Union(ud->m_minmax);
          if (redblue[0] != ud->m_redblue[0])
            redblue[0] = ON_UNSET_VALUE;
          if (redblue[1] != ud->m_redblue[1])
            redblue[1] = ON_UNSET_VALUE;
        }
        ud_count++;

        mesh_objects.Append(mesh_object);
      }
    }
  }

  if (0 == mesh_objects.Count())
    return failure;

  ON_Interval old_redblue = redblue;

  CRhinoCommand::result rc = cancel;
  if (context.IsInteractive())
    rc = GetDialogParameters(mesh_objects, minmax, old_redblue, redblue);
  else
    rc = GetScriptParameters(mesh_objects, minmax, old_redblue, redblue);

  ON_Interval colors = (rc == success ? redblue : old_redblue);

  for (i = 0; i < mesh_objects.Count(); i++)
  {
    ON_Mesh* mesh = const_cast<ON_Mesh*>(mesh_objects[i]->Mesh());
    if (mesh)
    {
      const CAnalysisUserData* ud = CAnalysisUserData::Get(mesh);
      if (ud)
      {
        if (ON_UNSET_VALUE != colors[0] || ON_UNSET_VALUE != colors[1])
        {
          ON_Interval new_redblue = ud->m_redblue;
          if (ON_UNSET_VALUE != colors[0])
            new_redblue[0] = colors[0];
          if (ON_UNSET_VALUE != colors[1])
            new_redblue[1] = colors[1];

          const_cast<CAnalysisUserData*>(ud)->m_redblue = new_redblue;
          CAnalysisUserData::UpdateColors(mesh);
        }
      }
    }
  }

  context.m_doc.Regen();

  return rc;
}

CRhinoCommand::result CCommandAnalyzeMesh::GetDialogParameters(
  const ON_SimpleArray<const CRhinoMeshObject*>& mesh_objects,
  const ON_Interval& minmax,
  const ON_Interval& old_redblue,
  ON_Interval& redblue
)
{
  RhinoApp().Print(RHSTR(L"Analysis parameter varies from %g to %g.\n"), minmax[0], minmax[1]);

  CAnalysisDialog dlg;
  dlg.m_minmax = minmax;
  dlg.m_range1 = old_redblue[0];
  dlg.m_range2 = old_redblue[1];

  int i;
  for (i = 0; i < mesh_objects.Count(); i++)
  {
    dlg.m_objects.Append(mesh_objects[i]->Attributes().m_uuid);
    dlg.m_meshes.Append(const_cast<ON_Mesh*>(mesh_objects[i]->Mesh()));
  }

  INT_PTR rc = dlg.DoModal();

  redblue[0] = dlg.m_range1;
  redblue[1] = dlg.m_range2;

  if (rc == IDCANCEL)
    return cancel;

  return success;
}

CRhinoCommand::result CCommandAnalyzeMesh::GetScriptParameters(
  const ON_SimpleArray<const CRhinoMeshObject*>& mesh_objects,
  const ON_Interval& minmax,
  const ON_Interval& old_redblue,
  ON_Interval& redblue
)
{
  RhinoApp().Print(RHSTR(L"Analysis parameter varies from %g to %g.\n"), minmax[0], minmax[1]);

  redblue = old_redblue;

  for (;;)
  {
    CRhinoGetOption go;
    go.SetCommandPrompt(RHSTR(L"Color range"));
    go.AcceptNothing();

    int red_opt = (redblue[0] != ON_UNSET_VALUE)
      ? go.AddCommandOptionNumber(RHCMDOPTNAME(L"Red"), &redblue[0], RHSTR(L"Red"))
      : go.AddCommandOption(RHCMDOPTNAME(L"Red"), RHCMDOPTVALUE(L"varies"));

    int blue_opt = (redblue[1] != ON_UNSET_VALUE)
      ? go.AddCommandOptionNumber(RHCMDOPTNAME(L"Blue"), &redblue[1], RHSTR(L"Blue"))
      : go.AddCommandOption(RHCMDOPTNAME(L"Blue"), RHCMDOPTVALUE(L"varies"));

    go.GetOption();
    if (go.CommandResult() != success)
      return go.CommandResult();

    if (CRhinoGet::option == go.Result())
    {
      const CRhinoCommandOption* opt = go.Option();
      if (opt)
      {
        if (red_opt == opt->m_option_index && ON_UNSET_VALUE == redblue[0])
        {
          CRhinoGetNumber gn;
          gn.SetCommandPrompt(RHSTR(L"Red"));
          gn.GetNumber();
          if (gn.CommandResult() != success)
            return gn.CommandResult();
          redblue[0] = gn.Number();
        }
        else if (blue_opt == opt->m_option_index && ON_UNSET_VALUE == redblue[1])
        {
          CRhinoGetNumber gn;
          gn.SetCommandPrompt(RHSTR(L"Blue"));
          gn.GetNumber();
          if (gn.CommandResult() != success)
            return gn.CommandResult();
          redblue[1] = gn.Number();
        }
      }
    }
    else
    {
      break;
    }
  }

  return success;
}

#pragma endregion

//
// END AnalyzeMesh command
//
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
