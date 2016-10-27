/////////////////////////////////////////////////////////////////////////////
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
};
