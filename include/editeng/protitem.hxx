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
#ifndef INCLUDED_EDITENG_PROTITEM_HXX
#define INCLUDED_EDITENG_PROTITEM_HXX

#include <svl/poolitem.hxx>
#include <editeng/editengdllapi.h>

class SvXMLUnitConverter;

// class SvxProtectItem --------------------------------------------------


/*  [Description]

    This item describes, if content, size or position should be protected.
*/

class EDITENG_DLLPUBLIC SvxProtectItem : public SfxPoolItem
{
    bool bCntnt :1;     // Content protected
    bool bSize  :1;     // Size protected
    bool bPos   :1;     // Position protected

public:
    TYPEINFO_OVERRIDE();

    explicit inline SvxProtectItem( const sal_uInt16 nId  );
    inline SvxProtectItem &operator=( const SvxProtectItem &rCpy );

    // "pure virtual Methods" from SfxPoolItem
    virtual bool             operator==( const SfxPoolItem& ) const SAL_OVERRIDE;

    virtual bool GetPresentation( SfxItemPresentation ePres,
                                    SfxMapUnit eCoreMetric,
                                    SfxMapUnit ePresMetric,
                                    OUString &rText, const IntlWrapper * = 0 ) const SAL_OVERRIDE;


    virtual SfxPoolItem*     Clone( SfxItemPool *pPool = 0 ) const SAL_OVERRIDE;
    virtual SfxPoolItem*     Create(SvStream &, sal_uInt16) const SAL_OVERRIDE;
    virtual SvStream&        Store(SvStream &, sal_uInt16 nItemVersion ) const SAL_OVERRIDE;

    bool IsCntntProtected() const { return bCntnt; }
    bool IsSizeProtected()  const { return bSize;  }
    bool IsPosProtected()   const { return bPos;   }
    void SetCntntProtect( bool bNew ) { bCntnt = bNew; }
    void SetSizeProtect ( bool bNew ) { bSize  = bNew; }
    void SetPosProtect  ( bool bNew ) { bPos   = bNew; }

    virtual bool            QueryValue( com::sun::star::uno::Any& rVal, sal_uInt8 nMemberId = 0 ) const SAL_OVERRIDE;
    virtual bool            PutValue( const com::sun::star::uno::Any& rVal, sal_uInt8 nMemberId = 0 ) SAL_OVERRIDE;
    void dumpAsXml(struct _xmlTextWriter* pWriter) const SAL_OVERRIDE;
};

inline SvxProtectItem::SvxProtectItem( const sal_uInt16 nId )
    : SfxPoolItem( nId )
{
    bCntnt = bSize = bPos = false;
}

inline SvxProtectItem &SvxProtectItem::operator=( const SvxProtectItem &rCpy )
{
    bCntnt = rCpy.IsCntntProtected();
    bSize  = rCpy.IsSizeProtected();
    bPos   = rCpy.IsPosProtected();
    return *this;
}




#endif

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
