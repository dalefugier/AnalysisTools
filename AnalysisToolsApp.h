// Copyright (c) 1993-2018 Robert McNeel & Associates. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// AnalysisToolsApp.h

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h" // main symbols

// CAnalysisToolsApp
// See AnalysisToolsApp.cpp for the implementation of this class
//

class CAnalysisToolsApp : public CWinApp
{
public:
	CAnalysisToolsApp() = default;
	BOOL InitInstance() override;
	int ExitInstance() override;
	DECLARE_MESSAGE_MAP()
};
