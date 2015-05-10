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
#ifndef INCLUDED_SVTOOLS_XWINDOWITEM_HXX
#define INCLUDED_SVTOOLS_XWINDOWITEM_HXX


#include <svtools/svtdllapi.h>

#include <svl/poolitem.hxx>
#include <toolkit/helper/vclunohelper.hxx>

#include <com/sun/star/awt/XWindow.hpp>

namespace vcl { class Window; }



class SVT_DLLPUBLIC XWindowItem : public SfxPoolItem
{
    ::com::sun::star::uno::Reference< ::com::sun::star::awt::XWindow >      m_xWin;

    XWindowItem & operator = ( const XWindowItem & ) SAL_DELETED_FUNCTION;

public:
    TYPEINFO_OVERRIDE();
    XWindowItem();
    XWindowItem( const XWindowItem &rItem );
    virtual ~XWindowItem();

    virtual SfxPoolItem*    Clone(SfxItemPool* pPool = 0) const SAL_OVERRIDE;
    virtual bool operator == ( const SfxPoolItem& rAttr ) const SAL_OVERRIDE;

    vcl::Window *        GetWindowPtr() const    { return VCLUnoHelper::GetWindow( m_xWin ); }
    com::sun::star::uno::Reference< com::sun::star::awt::XWindow >  GetXWindow() const  { return m_xWin; }
};



#endif

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
