// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisObject.h

#pragma once

class CAnalysisObject : public CCmdTarget
{
  DECLARE_DYNAMIC(CAnalysisObject)

public:
  CAnalysisObject();
  virtual ~CAnalysisObject();

  virtual void OnFinalRelease();

protected:
  DECLARE_MESSAGE_MAP()
  DECLARE_DISPATCH_MAP()
  DECLARE_INTERFACE_MAP()

  VARIANT IsAnalysisMesh(const VARIANT& vaObject);
  VARIANT AddAnalysisMesh(const VARIANT& vaVertices, const VARIANT& vaFaces, const VARIANT& vaData);
  VARIANT AnalysisMeshData(const VARIANT& vaObject, const VARIANT& vaData);
  VARIANT AnalysisMeshDisplayRange(const VARIANT& vaObject, const VARIANT& vaRange);
  VARIANT AnalysisMeshDataRange(const VARIANT& vaObject);

  enum
  {
    dispidIsAnalysisMesh = 1L,
    dispidAddAnalysisMesh,
    dispidAnalysisMeshData,
    dispidAnalysisMeshDisplayRange,
    dispidAnalysisMeshDataRange,
  };
};


