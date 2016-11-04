// Copyright (c) 1993-2016 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisToolsPlugIn.h

#pragma once

#include "AnalysisObject.h"

class CAnalysisToolsPlugIn : public CRhinoFileImportPlugIn
{
public:
  CAnalysisToolsPlugIn();
  ~CAnalysisToolsPlugIn();

  // Required overrides
  const wchar_t* PlugInName() const;
  const wchar_t* LocalPlugInName() const;
  const wchar_t* PlugInVersion() const;
  GUID PlugInID() const;
  BOOL OnLoadPlugIn();
  void OnUnloadPlugIn();

  // Plug-in COM object access
  LPUNKNOWN GetPlugInObjectInterface(const ON_UUID& iid);

  // File import plug-in overrides
  void AddFileType(ON_ClassArray<CRhinoFileType>&, const CRhinoFileReadOptions&);
  BOOL ReadFile(const wchar_t*, int, CRhinoDoc&, const CRhinoFileReadOptions&);

private:
  ON_Mesh* ReadFalseColorMeshFile(FILE* fp, ON_Mesh* mesh = 0);
  ON_Mesh* ReadStructuredTechPlotFile(FILE* fp, ON_Mesh* mesh = 0);

private:
  ON_wString m_plugin_version;
  int m_false_color_index;
  int m_tecplot_index;
  CAnalysisObject m_object;
};

CAnalysisToolsPlugIn& AnalysisToolsPlugIn();



