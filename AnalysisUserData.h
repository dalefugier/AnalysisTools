/////////////////////////////////////////////////////////////////////////////
// AnalysisUserData.h

#pragma once

class CAnalysisUserData : public ON_UserData
{
  ON_OBJECT_DECLARE(CAnalysisUserData);

public:

  /*
  Returns:
    The ON_UserData uuid value used to set m_userdata_uuid
    and passed to ON_Object::GetUserData().
  */
  static ON_UUID Id();

  /*
  Description:
    Get a pointer to CAnalysisUserData* from a mesh object.
  Parameters:
    mesh - [in]
  Returns:
    A pointer to CAnalysisUserData user data or NULL
    if none is attached to the mesh.
  */
  static
    const class CAnalysisUserData* Get(const ON_Mesh*);

  /*
  Description:
    Uses the information stored on the CAnalysisUserData
    to set the values in the mesh's m_C[] array.
  Parameters:
    mesh - [in]  If the mesh has CAnalysisUserData user
                 data attached and m_a.Count() = mesh->m_V.Count(), then
                 mesh->m_C[] is set
  Returns:
    True if successful.  False if the mesh doesn't have
    CAnalysisUserData user data or
    CAnalysisUserData.m_a.Count() is not equal
    to the mesh's vertex count.
  */
  static
    bool UpdateColors(ON_Mesh*);

  /*
  Description:
    Calculates the color that corresponds to an analysis parameter.
  Parameters:
    a - [in] analysis parameter
  Returns
    color
  */
  ON_Color Color(double a) const;

  CAnalysisUserData();
  ~CAnalysisUserData();
  CAnalysisUserData(const CAnalysisUserData&);
  CAnalysisUserData& operator=(const CAnalysisUserData&);

  BOOL GetDescription(ON_wString& description);
  BOOL Archive() const;
  BOOL Write(ON_BinaryArchive& archive) const;
  BOOL Read(ON_BinaryArchive& archive);

  // analysis parameters - one for each mesh vertex
  ON_SimpleArray<double> m_a;

  // minimum and maximum values in the m_a[] array.
  ON_Interval m_minmax;

  // m_redblue[0] is the analysis value that corresponds to red.
  // m_redblue[1] is the analysis value that corresponds to blue.
  // See the code for CAnalysisUserData::Color()
  // for more details.
  ON_Interval m_redblue;
};