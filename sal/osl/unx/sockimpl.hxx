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

#ifndef INCLUDED_SAL_OSL_UNX_SOCKIMPL_HXX
#define INCLUDED_SAL_OSL_UNX_SOCKIMPL_HXX

#include <osl/pipe.h>
#include <osl/socket.h>
#include <osl/interlck.h>

struct oslSocketImpl {
    int                 m_Socket;
    int                 m_nLastError;
    oslInterlockedCount m_nRefCount;
#if defined(LINUX)
    bool                m_bIsAccepting;
    bool                m_bIsInShutdown;
#endif
};

struct oslSocketAddrImpl
{
    sal_Int32 m_nRefCount;
    struct sockaddr m_sockaddr;
};

struct oslPipeImpl {
    int  m_Socket;
    sal_Char m_Name[PATH_MAX + 1];
    oslInterlockedCount m_nRefCount;
    bool m_bClosed;
#if defined(LINUX)
    bool m_bIsAccepting;
    bool m_bIsInShutdown;
#endif
};

oslSocket __osl_createSocketImpl(int Socket);
void __osl_destroySocketImpl(oslSocket pImpl);

#endif

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
