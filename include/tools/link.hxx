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
#ifndef INCLUDED_TOOLS_LINK_HXX
#define INCLUDED_TOOLS_LINK_HXX

#include <tools/toolsdllapi.h>
#include <sal/config.h>
#include <sal/types.h>
#include <tools/solar.h>

typedef sal_IntPtr (*PSTUB)( void*, void* );

#define DECL_LINK( Method, ArgType ) \
    sal_IntPtr Method( ArgType ); \
    static sal_IntPtr LinkStub##Method( void* pThis, void* )

#define DECL_STATIC_LINK( Class, Method, ArgType ) \
    static sal_IntPtr LinkStub##Method( void* pThis, void* ); \
    static sal_IntPtr Method( Class*, ArgType )

#define DECL_DLLPRIVATE_LINK(Method, ArgType) \
    SAL_DLLPRIVATE sal_IntPtr Method(ArgType); \
    SAL_DLLPRIVATE static sal_IntPtr LinkStub##Method(void * pThis, void *)

#define DECL_DLLPRIVATE_STATIC_LINK(Class, Method, ArgType) \
    SAL_DLLPRIVATE static sal_IntPtr LinkStub##Method( void* pThis, void* ); \
    SAL_DLLPRIVATE static sal_IntPtr Method(Class *, ArgType)

#define IMPL_STUB(Class, Method, ArgType) \
    sal_IntPtr Class::LinkStub##Method( void* pThis, void* pCaller) \
    { \
        return static_cast<Class*>(pThis)->Method( static_cast<ArgType>(pCaller) ); \
    }

#define IMPL_STATIC_LINK( Class, Method, ArgType, ArgName ) \
    sal_IntPtr Class::LinkStub##Method( void* pThis, void* pCaller) \
    { \
        return Method( static_cast<Class*>(pThis), static_cast<ArgType>(pCaller) ); \
    } \
    sal_IntPtr Class::Method( Class* pThis, ArgType ArgName )

#define IMPL_STATIC_LINK_NOINSTANCE( Class, Method, ArgType, ArgName ) \
    sal_IntPtr Class::LinkStub##Method( void* pThis, void* pCaller) \
    { \
        return Method( static_cast<Class*>(pThis), static_cast<ArgType>(pCaller) ); \
    } \
    sal_IntPtr Class::Method( SAL_UNUSED_PARAMETER Class*, ArgType ArgName )

#define IMPL_STATIC_LINK_NOINSTANCE_NOARG( Class, Method ) \
    sal_IntPtr Class::LinkStub##Method( void* pThis, void* pCaller) \
    { \
        return Method( static_cast<Class*>(pThis), pCaller ); \
    } \
    sal_IntPtr Class::Method( SAL_UNUSED_PARAMETER Class*, SAL_UNUSED_PARAMETER void* )

#define LINK( Inst, Class, Member ) \
    Link( static_cast<Class*>(Inst), &Class::LinkStub##Member )

#define STATIC_LINK( Inst, Class, Member ) LINK(Inst, Class, Member)

#define IMPL_LINK( Class, Method, ArgType, ArgName ) \
    IMPL_STUB( Class, Method, ArgType ) \
    sal_IntPtr Class::Method( ArgType ArgName )

#define IMPL_LINK_NOARG( Class, Method ) \
    IMPL_STUB( Class, Method, void* ) \
    sal_IntPtr Class::Method( SAL_UNUSED_PARAMETER void* )

#define IMPL_LINK_INLINE_START( Class, Method, ArgType, ArgName ) \
    inline sal_IntPtr Class::Method( ArgType ArgName )

#define IMPL_LINK_INLINE_END( Class, Method, ArgType, ArgName ) \
    IMPL_STUB( Class, Method, ArgType )

#define IMPL_LINK_NOARG_INLINE_START( Class, Method ) \
    inline sal_IntPtr Class::Method( SAL_UNUSED_PARAMETER void* )

#define IMPL_LINK_NOARG_INLINE_END( Class, Method ) \
    IMPL_STUB( Class, Method, void* )

#define IMPL_LINK_INLINE( Class, Method, ArgType, ArgName, Body ) \
    sal_IntPtr Class::Method( ArgType ArgName ) \
    Body \
    IMPL_STUB( Class, Method, ArgType )

#define EMPTYARG

class TOOLS_DLLPUBLIC Link
{
    void*       pInst;
    PSTUB       pFunc;

public:
                Link();
                Link( void* pLinkHdl, PSTUB pMemFunc );

    sal_IntPtr      Call( void* pCaller ) const;

    bool            IsSet() const;
    bool            operator !() const;

    bool            operator==( const Link& rLink ) const;
    bool            operator!=( const Link& rLink ) const
                    { return !(Link::operator==( rLink )); }
    bool            operator<( const Link& rLink ) const
                    { return reinterpret_cast<sal_uIntPtr>(rLink.pFunc) < reinterpret_cast<sal_uIntPtr>(pFunc); }
};

inline Link::Link()
{
    pInst = 0;
    pFunc = 0;
}

inline Link::Link( void* pLinkHdl, PSTUB pMemFunc )
{
    pInst = pLinkHdl;
    pFunc = pMemFunc;
}

inline sal_IntPtr Link::Call(void *pCaller) const
{
    return pFunc ? (*pFunc)(pInst, pCaller) : 0;
}

inline bool Link::IsSet() const
{
    if ( pFunc )
        return true;
    else
        return false;
}

inline bool Link::operator !() const
{
    if ( !pFunc )
        return true;
    else
        return false;
}

#endif

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
