/////////////////////////////////////////////////////////////////////////////
// AnalysisToolsPlugIn.cpp 

#include "StdAfx.h"
#include "AnalysisToolsPlugIn.h"
#include "AnalysisUserData.h"

#pragma warning( push )
#pragma warning( disable : 4073 )
#pragma init_seg( lib )
#pragma warning( pop )

// Rhino plug-in declaration
RHINO_PLUG_IN_DECLARE

// Rhino plug-in name
RHINO_PLUG_IN_NAME(L"Analysis Tools");

// Rhino plug-in id
RHINO_PLUG_IN_ID(L"7B9327EE-E27D-42D2-BA44-C8D444FBD62C");

// Rhino plug-in version
RHINO_PLUG_IN_VERSION(__DATE__"  "__TIME__)

// Rhino plug-in developer declarations
RHINO_PLUG_IN_DEVELOPER_ORGANIZATION(L"Robert McNeel & Associates");
RHINO_PLUG_IN_DEVELOPER_ADDRESS(L"3670 Woodland Park Avenue North\015\012Seattle WA 98103");
RHINO_PLUG_IN_DEVELOPER_COUNTRY(L"United States");
RHINO_PLUG_IN_DEVELOPER_PHONE(L"206-545-6877");
RHINO_PLUG_IN_DEVELOPER_FAX(L"206-545-7321");
RHINO_PLUG_IN_DEVELOPER_EMAIL(L"devsupport@mcneel.com");
RHINO_PLUG_IN_DEVELOPER_WEBSITE(L"http://www.rhino3d.com");
RHINO_PLUG_IN_UPDATE_URL(L"https://github.com/dalefugier/analysistools/");

// The one and only CAnalysisToolsPlugIn object
static CAnalysisToolsPlugIn thePlugIn;

const wchar_t* RHSTR(const wchar_t* s)
{
  return s;
}

const wchar_t* RHSTR_LIT(const wchar_t* s)
{
  return s;
}

/////////////////////////////////////////////////////////////////////////////
// CAnalysisToolsPlugIn definition

CAnalysisToolsPlugIn& AnalysisToolsPlugIn()
{
  return thePlugIn;
}

CAnalysisToolsPlugIn::CAnalysisToolsPlugIn()
{
  m_plugin_version = RhinoPlugInVersion();
}

CAnalysisToolsPlugIn::~CAnalysisToolsPlugIn()
{
}

/////////////////////////////////////////////////////////////////////////////
// Required overrides

const wchar_t* CAnalysisToolsPlugIn::PlugInName() const
{
  return RHSTR(L"Analysis Tools");
}

const wchar_t* CAnalysisToolsPlugIn::LocalPlugInName() const
{
  return RHSTR_LIT(L"Analysis Tools");
}

const wchar_t* CAnalysisToolsPlugIn::PlugInVersion() const
{
  return m_plugin_version;
}

GUID CAnalysisToolsPlugIn::PlugInID() const
{
  // {7B9327EE-E27D-42D2-BA44-C8D444FBD62C}
  return ON_UuidFromString(RhinoPlugInId());
}

BOOL CAnalysisToolsPlugIn::OnLoadPlugIn()
{
  return CRhinoFileImportPlugIn::OnLoadPlugIn();
}

void CAnalysisToolsPlugIn::OnUnloadPlugIn()
{
  CRhinoFileImportPlugIn::OnUnloadPlugIn();
}

LPUNKNOWN CAnalysisToolsPlugIn::GetPlugInObjectInterface(const ON_UUID& iid)
{
  LPUNKNOWN lpUnknown = 0;

  if (iid == IID_IUnknown)
    lpUnknown = m_object.GetInterface(&IID_IUnknown);

  else if (iid == IID_IDispatch)
    lpUnknown = m_object.GetInterface(&IID_IDispatch);

  if (lpUnknown)
    m_object.ExternalAddRef();

  return lpUnknown;
}

void CAnalysisToolsPlugIn::AddFileType(ON_ClassArray<CRhinoFileType>& ftypes, const CRhinoFileReadOptions& options)
{
  m_false_color_index = ftypes.Count();
  CRhinoFileType ft1(PlugInID(), RHSTR(L"Rhino Analysis Mesh Files (*.ram)"), L"ram");
  ftypes.Append(ft1);

  m_tecplot_index = ftypes.Count();
  CRhinoFileType ft2(PlugInID(), RHSTR(L"Tecplot Files (*.tp)"), L"tp");
  ftypes.Append(ft2);
}

static bool IsNumeric(wchar_t c)
{
  bool rc = false;
  switch (c)
  {
  case '.':
  case '+':
  case '-':
  case 'e':
  case 'E':
  case 'd':
  case 'D':
    rc = true;
    break;
  default:
    if (c >= '0' && c <= '9')
      rc = true;
    break;
  }

  return rc;
}

static bool IsAlpha(wchar_t c)
{
  return ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'));
}

static bool IsAlphaNumeric(wchar_t c)
{
  return (IsAlpha(c) || IsNumeric(c));
}

static const wchar_t* SkipJunk(const wchar_t* s)
{
  if (s)
  {
    while (*s && !IsAlphaNumeric(*s))
      s++;
  }
  return s;
}

static const wchar_t* SkipAlphaJunk(const wchar_t* s)
{
  if (s)
  {
    while (*s != 0 && *s != '+' && *s != '-' && *s != '.' && (*s < '0' || *s > '9'))
      s++;
  }
  return s;
}

static const wchar_t* ParseInt(const wchar_t* s, int& i)
{
  if (s)
  {
    int value = 0;
    int sgn = 1;
    if ('+' == *s)
      s++;
    else if ('-' == *s)
    {
      sgn = -1;
      s++;
    }

    if (*s < '0' || *s > '9')
      s = 0;
    else
    {
      while (*s >= '0' && *s <= '9')
      {
        value = value * 10 + ((int)(*s - '0'));
        s++;
      }
      i = value * sgn;
    }
  }

  return s;
}

static const wchar_t* ParseDouble(const wchar_t* s, double& x)
{
  if (s)
  {
    wchar_t buffer[512];
    wmemset(buffer, 0, 512);

    int buffer_length = 0;
    bool bHaveDecimal = false;
    bool bHaveMantissa = false;

    if ('+' == *s || '-' == *s)
      buffer[buffer_length++] = *s++;

    // 14-Aug-2013 Dale Fugier
    // Deal with "-.0001" formatted values
    if ('.' == *s)
    {
      bHaveDecimal = true;
      buffer[buffer_length++] = *s++;
    }

    if ((*s < '0' || *s > '9'))
      s = 0;
    else
    {
      while (*s >= '0' && *s <= '9' && buffer_length < 512)
      {
        bHaveMantissa = true;
        buffer[buffer_length++] = *s++;
      }

      if ('.' == *s  && buffer_length < 512)
      {
        if (bHaveDecimal)
          return 0;

        bHaveDecimal = true;
        buffer[buffer_length++] = *s++;

        while (*s >= '0' && *s <= '9' && buffer_length < 512)
        {
          bHaveMantissa = true;
          buffer[buffer_length++] = *s++;
        }
      }

      if (('e' == *s || 'E' == *s) && buffer_length < 512)
      {
        if (0 == buffer_length)
          buffer[buffer_length++] = '1';
        buffer[buffer_length++] = *s++;

        if (('-' == *s || '+' == *s) && buffer_length < 512)
          buffer[buffer_length++] = *s++;

        if (*s < '0' || *s > '9')
          return 0;

        while (*s >= '0' && *s <= '9' && buffer_length < 512)
          buffer[buffer_length++] = *s++;
      }

      if (buffer_length >= 512 || IsNumeric(*s))
        return 0;

      buffer[buffer_length] = 0;
      double v = ON_UNSET_VALUE;

      if (1 == swscanf(buffer, L"%lg", &v))
        x = v;
      else
        s = 0;
    }
  }

  return s;
}

static const wchar_t* ParseCount(const wchar_t* line, const wchar_t* string, int& count)
{
  count = 0;
  if (0 == line || 0 == string)
    return 0;

  const size_t string_length = wcslen(string);
  if (string_length <= 0)
    return 0;

  const wchar_t* s = SkipJunk(line);
  if (0 == s)
    return 0;

  if (_wcsnicmp(string, s, string_length))
    return 0;

  s = SkipJunk(s + string_length);
  int i = 0;
  s = ParseInt(s, count);
  if (0 == s)
    count = 0;

  return s;
}

static const wchar_t* ParseVertex(const wchar_t* line, ON_3dPoint& v, double& c)
{
  double x = ON_UNSET_VALUE, y = ON_UNSET_VALUE, z = ON_UNSET_VALUE, a = ON_UNSET_VALUE;
  line = SkipAlphaJunk(line);
  line = ParseDouble(line, x);
  line = SkipAlphaJunk(line);
  line = ParseDouble(line, y);
  line = SkipAlphaJunk(line);
  line = ParseDouble(line, z);
  line = SkipAlphaJunk(line);
  line = ParseDouble(line, a);
  if (line)
  {
    if (ON_IsValid(x) && ON_IsValid(y) && ON_IsValid(z) && ON_IsValid(a))
    {
      v.Set(x, y, z);
      c = a;
    }
    else
      line = 0;
  }
  return line;
}

static const wchar_t* ParseFace(const wchar_t* line, const int vcount, ON_MeshFace& f)
{
  int a = -1, b = -1, c = -1, d = -1;
  line = SkipAlphaJunk(line);
  line = ParseInt(line, a);
  line = SkipAlphaJunk(line);
  line = ParseInt(line, b);
  line = SkipAlphaJunk(line);
  line = ParseInt(line, c);
  line = SkipAlphaJunk(line);
  if (*line)
  {
    line = ParseInt(line, d);
  }
  else
    d = c;

  if (line)
  {
    if (a >= 0 && b >= 0 && c >= 0 && d >= 0 &&
      a < vcount && b < vcount && c < vcount && d < vcount &&
      a != b && a != c && a != d &&
      b != c && b != d
      )
    {
      f.vi[0] = a;
      f.vi[1] = b;
      f.vi[2] = c;
      f.vi[3] = d;
    }
    else
      line = 0;
  }
  return line;
}

/////////////////////////////////////////////////////////////////////////////

class TP_POINT
{
public:
  int i, j, k;
  int vindex;
  ON_3dPoint v;
  double a;
};

class TP_GRID
{
public:
  TP_GRID(int im, int jm, int km) : IMAX(im), JMAX(jm), KMAX(km), tp(IMAX*JMAX*KMAX) {};
  int IMAX;
  int JMAX;
  int KMAX;
  ON_SimpleArray<TP_POINT> tp;
  int Index(int i, int j, int k) { return (i + (j + k*JMAX)*IMAX); }
  TP_POINT& Point(int i, int j, int k) { return tp[(i + (j + k*JMAX)*IMAX)]; }
};

/////////////////////////////////////////////////////////////////////////////

ON_Mesh* CAnalysisToolsPlugIn::ReadStructuredTechPlotFile(FILE* fp, ON_Mesh* mesh)
{
  if (mesh)
    mesh->Destroy();

  wchar_t line[129];
  wmemset(line, 0, 129);

  int IMAX = 0;
  int JMAX = 0;
  int KMAX = 0;
  const wchar_t* s = 0;

  while (IMAX <= 0)
  {
    while (0 == s || *s == 0)
    {
      s = fgetws(line, 127, fp);
      if (0 == s)
        return 0;
    }
    s = ParseCount(s, L"I=", IMAX);
  }

  while (JMAX <= 0)
  {
    while (0 == s || *s == 0)
    {
      s = fgetws(line, 127, fp);
      if (0 == s)
        return 0;
    }
    s = ParseCount(s, L"J=", JMAX);
  }

  while (KMAX <= 0)
  {
    while (0 == s || *s == 0)
    {
      s = fgetws(line, 127, fp);
      if (0 == s)
        return 0;
    }
    s = ParseCount(s, L"K=", KMAX);
  }

  s = fgetws(line, 127, fp);
  s = fgetws(line, 127, fp);

  TP_POINT p;
  TP_GRID grid(IMAX, JMAX, KMAX);
  int i, j, k;
  for (k = 0; k < KMAX && s; k++)
  {
    for (j = 0; j < JMAX && s; j++)
    {
      for (i = 0; i < IMAX && s; i++)
      {
        s = fgetws(line, 127, fp);
        s = SkipJunk(s);
        s = ParseDouble(s, p.v.x);
        s = SkipJunk(s);
        s = ParseDouble(s, p.v.y);
        s = SkipJunk(s);
        s = ParseDouble(s, p.v.z);
        s = SkipJunk(s);
        s = ParseDouble(s, p.a);
        if (s)
        {
          p.i = i; p.j = j; p.k = k; p.vindex = -1;
          grid.tp.Append(p);
        }
      }
    }
  }

  if (grid.tp.Count() == IMAX * JMAX * KMAX)
  {
    if (0 == mesh)
      mesh = new ON_Mesh();

    CAnalysisUserData* ud = new CAnalysisUserData();
    double mn = 1.0e300;
    double mx = -mn;
    TP_POINT tp00, tp01, tp10, tp11;
    ON_MeshFace f;

    for (k = 0; k < KMAX && s; k++) for (j = 0; j < JMAX && s; j++) for (i = 0; i < IMAX && s; i++)
    {
      grid.Point(i, j, k).vindex = mesh->m_V.Count();
      tp11 = grid.Point(i, j, k);
      ud->m_a.Append(tp11.a);
      mesh->m_V.Append(ON_3fPoint(tp11.v));
      if (tp11.a < mn)
        mn = tp11.a;
      if (tp11.a > mx)
        mx = tp11.a;
      if (i > 0)
      {
        if (j > 0)
        {
          tp00 = grid.Point(i - 1, j - 1, k);
          tp10 = grid.Point(i, j - 1, k);
          tp01 = grid.Point(i - 1, j, k);
          f.vi[0] = tp00.vindex;
          f.vi[1] = tp10.vindex;
          f.vi[2] = tp11.vindex;
          f.vi[3] = tp01.vindex;
          mesh->m_F.Append(f);
          if (k > 0)
          {
            tp11 = grid.Point(i, j, k);
            tp01 = grid.Point(i, j - 1, k);
            tp00 = grid.Point(i, j - 1, k - 1);
            tp10 = grid.Point(i, j, k - 1);
            f.vi[0] = tp00.vindex;
            f.vi[1] = tp10.vindex;
            f.vi[2] = tp11.vindex;
            f.vi[3] = tp01.vindex;
            mesh->m_F.Append(f);

            tp11 = grid.Point(i, j, k);
            tp01 = grid.Point(i - 1, j, k);
            tp00 = grid.Point(i - 1, j, k - 1);
            tp10 = grid.Point(i, j, k - 1);
            f.vi[0] = tp00.vindex;
            f.vi[1] = tp10.vindex;
            f.vi[2] = tp11.vindex;
            f.vi[3] = tp01.vindex;
            mesh->m_F.Append(f);

          }
        }
        else if (k > 0)
        {
          tp00 = grid.Point(i - 1, j, k - 1);
          tp10 = grid.Point(i, j, k - 1);
          tp01 = grid.Point(i - 1, j, k);
          f.vi[0] = tp00.vindex;
          f.vi[1] = tp10.vindex;
          f.vi[2] = tp11.vindex;
          f.vi[3] = tp01.vindex;
          mesh->m_F.Append(f);
        }
      }
      else if (j > 0)
      {
        if (k > 0)
        {
          const TP_POINT tp00 = grid.Point(i, j - 1, k - 1);
          const TP_POINT tp10 = grid.Point(i, j, k - 1);
          const TP_POINT tp01 = grid.Point(i, j - 1, k);
          f.vi[0] = tp00.vindex;
          f.vi[1] = tp10.vindex;
          f.vi[2] = tp11.vindex;
          f.vi[3] = tp01.vindex;
          mesh->m_F.Append(f);
        }
      }
    }

    mesh->ComputeVertexNormals();

    ud->m_minmax.Set(mn, mx);
    ud->m_redblue.Set(mn, mx);
    mesh->AttachUserData(ud);
    CAnalysisUserData::UpdateColors(mesh);
  }
  else
    mesh = 0;

  return mesh;
}

ON_Mesh* CAnalysisToolsPlugIn::ReadFalseColorMeshFile(FILE* fp, ON_Mesh* mesh)
{
  wchar_t line[129];
  wmemset(line, 0, 129);

  int vcount = 0;
  int fcount = 0;

  while (vcount <= 0 && fgetws(line, 127, fp))
    ParseCount(line, L"vertexcount", vcount);
  if (vcount < 3)
    return 0;

  if (fgetws(line, 127, fp))
    ParseCount(line, L"facecount", fcount);
  if (fcount <= 0)
    return 0;

  ON_SimpleArray<ON_3dPoint> v(vcount);
  ON_SimpleArray<double> a(vcount);

  int i;
  for (i = 0; i < vcount; i++)
  {
    if (!fgetws(line, 127, fp))
      return 0;
    if (!ParseVertex(line, v.AppendNew(), a.AppendNew()))
      return 0;
  }

  ON_SimpleArray<ON_MeshFace> f(fcount);
  for (i = 0; i < fcount; i++)
  {
    if (!fgetws(line, 127, fp))
      return 0;
    if (!ParseFace(line, vcount, f.AppendNew()))
      return 0;
  }

  if (mesh)
    mesh->Destroy();
  else
    mesh = new ON_Mesh();

  mesh->m_V.Reserve(vcount);
  mesh->m_F.Reserve(fcount);

  for (i = 0; i < vcount; i++)
    mesh->m_V.AppendNew() = v[i];

  for (i = 0; i < fcount; i++)
    mesh->m_F.AppendNew() = f[i];

  mesh->ComputeVertexNormals();

  CAnalysisUserData* ud = new CAnalysisUserData();
  ud->m_a = a;

  double mn, mx;
  mn = mx = a[0];

  for (i = 1; i < a.Count(); i++)
  {
    double x = a[i];
    if (x < mn)
      mn = x;
    else if (x > mx)
      mx = x;
  }

  ud->m_minmax.Set(mn, mx);
  ud->m_redblue.Set(mn, mx);
  mesh->AttachUserData(ud);
  CAnalysisUserData::UpdateColors(mesh);

  return mesh;
}

BOOL CAnalysisToolsPlugIn::ReadFile(const wchar_t* filename, int index, CRhinoDoc& doc, const CRhinoFileReadOptions& options)
{
  ON_Workspace ws;
  bool rc = false;

  bool bBatchMode = (0 != options.Mode(CRhinoFileReadOptions::BatchMode));

  if (filename && filename[0])
  {
    FILE* fp = ws.OpenFile(filename, L"r");
    if (0 == fp)
    {
      ON_wString msg;
      msg.Format(RHSTR(L"Unable to open file \"%s\""), filename);
      const wchar_t* s = msg.Array();
      if (bBatchMode)
        RhinoApp().Print(L"%s\n", s);
      else
        RhinoMessageBox(s, RHSTR(L"Unable to open file"), MB_OK);
    }
    else
    {
      ON_Mesh* mesh = 0;
      if (index == m_tecplot_index)
        mesh = ReadStructuredTechPlotFile(fp);
      else
        mesh = ReadFalseColorMeshFile(fp);

      if (0 == mesh)
      {
        ON_wString msg;
        msg.Format(RHSTR(L"Unable to read file \"%s\""), filename);
        const wchar_t* s = msg.Array();
        if (bBatchMode)
          RhinoApp().Print(L"%s\n", s);
        else
          RhinoMessageBox(s, RHSTR(L"Unable to open file"), MB_OK);
      }
      else
      {
        CRhinoMeshObject* mesh_object = new CRhinoMeshObject();
        mesh_object->SetMesh(mesh);
        rc = doc.AddObject(mesh_object);
      }
    }
  }

  if (rc &&  options.Mode(CRhinoFileReadOptions::OpenMode))
  {
    ON_SimpleArray<CRhinoView*> view_list;
    doc.GetViewList(view_list, true, false);
    for (int i = 0; i < view_list.Count(); i++)
    {
      CRhinoView* view = view_list[i];
      if (view)
      {
        double border = (view->ActiveViewport().VP().Projection() == ON::parallel_view) ? 1.1 : 0.0;
        RhinoDollyExtents(view->ActiveViewport(), border);
      }
    }

    doc.Redraw();
  }

  return rc;
}