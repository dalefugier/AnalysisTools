/////////////////////////////////////////////////////////////////////////////
// AnalysisDialog.cpp

#include "stdafx.h"
#include "AnalysisDialog.h"
#include "AnalysisUserData.h"

#define CLAMP(V,L,H) ( (V) < (L) ? (L) : ( (V) > (H) ? (H) : (V) ) )
#define LERP(A,L,H) ((L)+((H)-(L))*(A))
#define EPS_DIVIDE	0.000001
static double BAD_COLOR_HUE = 0.0;                 // red
static double GOOD_COLOR_HUE = ON_PI * 4.0 / 3.0; // blue 

/////////////////////////////////////////////////////////////////////////////

IMPLEMENT_DYNAMIC(CAnalysisDialog, CRhinoDialog)

CAnalysisDialog::CAnalysisDialog(CWnd* pParent)
  : CRhinoDialog(CAnalysisDialog::IDD, pParent)
{
  SetEnableDisplayCommands(TRUE);

  m_range1 = m_range2 = 0.0;
  m_range1_timer_on = m_range2_timer_on = false;
  m_updating = true;
}

CAnalysisDialog::~CAnalysisDialog()
{
}

void CAnalysisDialog::DoDataExchange(CDataExchange* pDX)
{
  CRhinoDialog::DoDataExchange(pDX);
  DDX_Control(pDX, IDC_HUEBAR_BUTTON, m_huebar_button);
  DDX_Control(pDX, IDC_RANGE1_EDIT, m_range1_edit);
  DDX_Control(pDX, IDC_RANGE2_EDIT, m_range2_edit);
  DDX_Control(pDX, IDC_MIDRANGE_STATIC, m_midrange_static);
  m_range1_edit.DDX_Text(pDX, IDC_RANGE1_EDIT, m_range1);
  m_range2_edit.DDX_Text(pDX, IDC_RANGE2_EDIT, m_range2);
}

BEGIN_MESSAGE_MAP(CAnalysisDialog, CRhinoDialog)
  ON_WM_DRAWITEM()
  ON_BN_CLICKED(IDC_AUTO_BUTTON, &CAnalysisDialog::OnBnClickedAutoButton)
  ON_EN_CHANGE(IDC_RANGE1_EDIT, &CAnalysisDialog::OnEnChangeRange1Edit)
  ON_EN_CHANGE(IDC_RANGE2_EDIT, &CAnalysisDialog::OnEnChangeRange2Edit)
  ON_WM_TIMER()
END_MESSAGE_MAP()

BOOL CAnalysisDialog::OnInitDialog()
{
  CRhinoDialog::OnInitDialog();

  CreateHueBar();

  m_range1_edit.SetDisplayPrecision(10);
  m_range2_edit.SetDisplayPrecision(10);

  CString s;
  s.Format(L"%.5g", 0.5 * (m_range1 + m_range2));
  m_midrange_static.SetWindowText(s);

  m_conduit.Enable();
  m_conduit.Attach(this);
  CRhinoDoc* doc = RhinoApp().ActiveDoc();
  if (doc)
    doc->Regen();

  m_updating = false;

  return TRUE;
}

static void HSVToRGB(double h, double s, double v, double& r, double& g, double& b)
{
  int i;
  double f, p, q, t;

  if (s < EPS_DIVIDE)
  {
    r = v;
    g = v;
    b = v;
  }
  else
  {
    h = h * 3.0 / ON_PI;  // (6.0 / 2.0 * ON_PI);
    i = (int)floor(h);
    if (i > 5 || i < 0)
    {
      i = 0;
      h = 0.0;
    }
    f = h - i;
    p = v * (1.0 - s);
    q = v * (1.0 - (s * f));
    t = v * (1.0 - (s * (1.0 - f)));

    switch (i)
    {
    case 0:
      r = v;
      g = t;
      b = p;
      break;
    case 1:
      r = q;
      g = v;
      b = p;
      break;
    case 2:
      r = p;
      g = v;
      b = t;
      break;
    case 3:
      r = p;
      g = q;
      b = v;
      break;
    case 4:
      r = t;
      g = p;
      b = v;
      break;
    case 5:
      r = v;
      g = p;
      b = q;
      break;
    }
  }
}

static void HSVToRGBInt(double h, double s, double v, int& r_out, int& g_out, int& b_out)
{
  double r, g, b;

  s = CLAMP(s, 0.0, 1.0);
  v = CLAMP(v, 0.0, 1.0);

  HSVToRGB(h, s, v, r, g, b);

  r = CLAMP(r, 0.0, 1.0);
  g = CLAMP(g, 0.0, 1.0);
  b = CLAMP(b, 0.0, 1.0);

  r *= 255.0;
  g *= 255.0;
  b *= 255.0;

  r_out = ON_Round(r);
  g_out = ON_Round(g);
  b_out = ON_Round(b);
}

void CAnalysisDialog::CreateHueBar()
{
  CRect rect;
  m_huebar_button.GetClientRect(rect);

  int width = rect.Width();
  int height = rect.Height();

  LPBITMAPINFO dib = m_huebar_dib.CreateDib(width, height, 24);
  if (0 == dib)
    return;

  char* baseptr = (char*)m_huebar_dib.FindDIBBits();
  int scan_width = m_huebar_dib.ScanWidth();

  double start_angle = GOOD_COLOR_HUE;
  double end_angle = BAD_COLOR_HUE;

  int i, j;
  for (i = 0; i < height; i++)
  {
    double t = (double)i / (double)(height - 1);
    double angle = LERP(t, start_angle, end_angle);

    int ir, ig, ib;
    HSVToRGBInt(angle, 1.0, 1.0, ir, ig, ib);
    char r = (unsigned char)ir;
    char g = (unsigned char)ig;
    char b = (unsigned char)ib;
    char* ptr = baseptr + (i * scan_width);

    for (j = 0; j < width; j++)
    {
      ptr[0] = b;
      ptr[1] = g;
      ptr[2] = r;
      ptr += 3;
    }
  }
}

void CAnalysisDialog::OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct)
{
  if (IDC_HUEBAR_BUTTON == nIDCtl)
  {
    m_huebar_dib.DCSelectBitmap(true);
    CDC dc;
    dc.Attach(lpDrawItemStruct->hDC);
    dc.BitBlt(0, 0, m_huebar_dib.Width(), m_huebar_dib.Height(), m_huebar_dib, 0, 0, SRCCOPY);
    dc.Detach();
    return;
  }

  CRhinoDialog::OnDrawItem(nIDCtl, lpDrawItemStruct);
}

void CAnalysisDialog::OnBnClickedAutoButton()
{
  if (m_updating)
    return;

  m_range1 = m_minmax[0];
  m_range2 = m_minmax[1];

  m_updating = true;
  UpdateData(FALSE);
  m_updating = false;

  UpdateRange(range1_timer);
}

void CAnalysisDialog::UpdateRange(EditTimers timer_id)
{
  bool& timer_on = (timer_id == range1_timer ? m_range1_timer_on : m_range2_timer_on);
  KillTimer(timer_id);
  timer_on = false;

  UpdateData(TRUE);

  m_updating = true;

  CString s;
  s.Format(L"%.5g", 0.5 * (m_range1 + m_range2));
  m_midrange_static.SetWindowText(s);

  for (int i = 0; i < m_meshes.Count(); i++)
  {
    ON_Mesh* mesh = m_meshes[i];
    if (mesh)
    {
      const CAnalysisUserData* ud = CAnalysisUserData::Get(mesh);
      if (ud)
      {
        ON_Interval redblue = ud->m_redblue;
        redblue[0] = m_range1;
        redblue[1] = m_range2;
        const_cast<CAnalysisUserData*>(ud)->m_redblue = redblue;
        CAnalysisUserData::UpdateColors(mesh);
      }
    }
  }

  CRhinoDoc* doc = RhinoApp().ActiveDoc();
  if (doc)
    doc->Regen();

  m_updating = false;
}

void CAnalysisDialog::OnEnChange(EditTimers timer_id)
{
  bool& on = (timer_id == range1_timer ? m_range1_timer_on : m_range2_timer_on);
  if (on)
    KillTimer(timer_id);

  on = false;

  if (SetTimer(timer_id, 650, 0))
    on = true;
  else
    UpdateRange(timer_id);
}

void CAnalysisDialog::OnEnChangeRange1Edit()
{
  if (m_updating)
    return;

  OnEnChange(range1_timer);
}

void CAnalysisDialog::OnEnChangeRange2Edit()
{
  if (m_updating)
    return;

  OnEnChange(range2_timer);
}

void CAnalysisDialog::OnTimer(UINT_PTR nIDEvent)
{
  switch (nIDEvent)
  {
  case range1_timer:
  case range2_timer:
    UpdateRange((EditTimers)nIDEvent);
    return;
  }
  CRhinoDialog::OnTimer(nIDEvent);
}

void CAnalysisDialog::OnOK()
{
  CRhinoDialog::OnOK();
}

void CAnalysisDialog::OnCancel()
{
  UpdateData(TRUE);
  CRhinoDialog::OnCancel();
}
