// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisDialogConduit.h

#pragma once

class CAnalysisDialog;

class CAnalysisDialogConduit : public CRhinoDisplayConduit
{
public:
  CAnalysisDialogConduit();
  void Attach(const CAnalysisDialog*);
  bool ExecConduit(CRhinoDisplayPipeline&, UINT, bool&);

private:
  const CAnalysisDialog* m_dialog;
  ON_Color m_color;
};
