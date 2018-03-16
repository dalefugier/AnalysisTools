// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisToolsPlugIn.h

#pragma once

#include "AnalysisObject.h"

// CAnalysisToolsPlugIn
// See AnalysisToolsPlugIn.cpp for the implementation of this class
//

class CAnalysisToolsPlugIn : public CRhinoFileImportPlugIn
{
public:
  CAnalysisToolsPlugIn();
  ~CAnalysisToolsPlugIn() = default;

  // Required overrides
  const wchar_t* PlugInName() const override;
  const wchar_t* PlugInVersion() const override;
  GUID PlugInID() const override;
  
  // Additional overrides
  BOOL OnLoadPlugIn() override;
  void OnUnloadPlugIn() override;
  LPUNKNOWN GetPlugInObjectInterface(const ON_UUID& iid);

  // File import overrides
  void AddFileType(ON_ClassArray<CRhinoFileType>& extensions, const CRhinoFileReadOptions& options) override;
  BOOL ReadFile(const wchar_t* filename, int index, CRhinoDoc& doc, const CRhinoFileReadOptions& options) override;

private:
  ON_Mesh* ReadFalseColorMeshFile(FILE* fp, ON_Mesh* mesh = nullptr);
  ON_Mesh* ReadStructuredTechPlotFile(FILE* fp, ON_Mesh* mesh = nullptr);

private:
  ON_wString m_plugin_version;
  int m_false_color_index;
  int m_tecplot_index;
  CAnalysisObject m_object;
};

// Return a reference to the one and only CAnalysisToolsPlugIn object
CAnalysisToolsPlugIn& AnalysisToolsPlugIn();



