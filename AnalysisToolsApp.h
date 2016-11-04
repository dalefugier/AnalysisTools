// Copyright (c) 1993-2016 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisToolsApp.h

#pragma once

#ifndef __AFXWIN_H__
#error include 'stdafx.h' before including this file for PCH
#endif

#include "Resource.h"

class CAnalysisToolsApp : public CWinApp
{
public:
  CAnalysisToolsApp();

public:
  virtual BOOL InitInstance();
  virtual int ExitInstance();
  DECLARE_MESSAGE_MAP()
};
