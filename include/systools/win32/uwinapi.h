/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This file incorporates work covered by the following license notice:
 *
 *   Licensed to the Apache Software Foundation (ASF) under one or more
 *   contributor license agreements. See the NOTICE file distributed
 *   with this work for additional information regarding copyright
 *   ownership. The ASF licenses this file to you under the Apache
 *   License, Version 2.0 (the "License"); you may not use this file
 *   except in compliance with the License. You may obtain a copy of
 *   the License at http://www.apache.org/licenses/LICENSE-2.0 .
 */

#ifndef INCLUDED_SYSTOOLS_WIN32_UWINAPI_H
#define INCLUDED_SYSTOOLS_WIN32_UWINAPI_H

#include <sal/macros.h>
#ifdef _UWINAPI_
#   define _KERNEL32_
#   define _USER32_
#   define _SHELL32_
#endif

#ifndef _WINDOWS_
#ifdef _MSC_VER
#   pragma warning(push,1) /* disable warnings within system headers */
#endif
#   include <windows.h>
#ifdef _MSC_VER
#   pragma warning(pop)
#endif
#endif

#ifdef __MINGW32__
#include <basetyps.h>
#ifdef _UWINAPI_
#define WINBASEAPI
#endif
#endif

#ifdef __cplusplus

inline bool IsValidHandle(HANDLE handle)
{
    return  handle != INVALID_HANDLE_VALUE && handle != NULL;
}

#else   /* __cplusplus */

#define IsValidHandle(Handle)   ((DWORD_PTR)(Handle) + 1 > 1)

#endif  /* __cplusplus */

#endif // INCLUDED_SYSTOOLS_WIN32_UWINAPI_H

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */