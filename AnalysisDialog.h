/////////////////////////////////////////////////////////////////////////////
// AnalysisDialog.h

#pragma once

#include "Resource.h"
#include "AnalysisDialogConduit.h"

class CAnalysisDialog : public CRhinoDialog
{
  DECLARE_DYNAMIC(CAnalysisDialog)

public:
  CAnalysisDialog(CWnd* pParent = NULL);   // standard constructor
  virtual ~CAnalysisDialog();

  // Dialog Data
  enum { IDD = IDD_ANALYSIS_DIALOG };
  CButton m_huebar_button;
  CRhinoUiEditReal m_range1_edit;
  CRhinoUiEditReal m_range2_edit;
  CStatic m_midrange_static;

  CRhinoUiDib m_huebar_dib;
  double m_range1;
  double m_range2;

  ON_Interval m_minmax;
  ON_SimpleArray<ON_UUID> m_objects;
  ON_SimpleArray<ON_Mesh*> m_meshes;

  CAnalysisDialogConduit m_conduit;

public:
  virtual BOOL OnInitDialog();
  afx_msg void OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct);
  afx_msg void OnBnClickedAutoButton();
  afx_msg void OnEnChangeRange1Edit();
  afx_msg void OnEnChangeRange2Edit();
  afx_msg void OnTimer(UINT_PTR nIDEvent);

protected:

  enum EditTimers
  {
    range1_timer = WM_USER,
    range2_timer,
  };

  void OnEnChange(EditTimers timer_id);
  void UpdateRange(EditTimers timer_id);
  void CreateHueBar();

  bool m_range1_timer_on;
  bool m_range2_timer_on;
  bool m_updating;

protected:
  virtual void DoDataExchange(CDataExchange* pDX);
  DECLARE_MESSAGE_MAP()
  virtual void OnOK();
  virtual void OnCancel();
};
