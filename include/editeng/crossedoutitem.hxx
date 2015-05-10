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
#ifndef INCLUDED_EDITENG_CROSSEDOUTITEM_HXX
#define INCLUDED_EDITENG_CROSSEDOUTITEM_HXX

#include <vcl/vclenum.hxx>
#include <svl/eitem.hxx>
#include <editeng/editengdllapi.h>

class SvXMLUnitConverter;

// class SvxCrossedOutItem -----------------------------------------------

/*  [Description]
    This item describes, whether and how it is striked out.
*/

class EDITENG_DLLPUBLIC SvxCrossedOutItem : public SfxEnumItem
{
public:
    TYPEINFO_OVERRIDE();

    SvxCrossedOutItem( const FontStrikeout eSt /*= STRIKEOUT_NONE*/,
                       const sal_uInt16 nId  );

    // "pure virtual Methods" from SfxPoolItem
    virtual bool GetPresentation( SfxItemPresentation ePres,
                                    SfxMapUnit eCoreMetric,
                                    SfxMapUnit ePresMetric,
                                    OUString &rText, const IntlWrapper * = 0 ) const SAL_OVERRIDE;

    virtual SfxPoolItem*    Clone( SfxItemPool *pPool = 0 ) const SAL_OVERRIDE;
    virtual SfxPoolItem*    Create(SvStream &, sal_uInt16) const SAL_OVERRIDE;
    virtual SvStream&       Store(SvStream &, sal_uInt16 nItemVersion) const SAL_OVERRIDE;
    virtual OUString   GetValueTextByPos( sal_uInt16 nPos ) const SAL_OVERRIDE;
    virtual sal_uInt16      GetValueCount() const SAL_OVERRIDE;
    virtual bool            QueryValue( com::sun::star::uno::Any& rVal, sal_uInt8 nMemberId = 0 ) const SAL_OVERRIDE;
    virtual bool            PutValue( const com::sun::star::uno::Any& rVal, sal_uInt8 nMemberId = 0 ) SAL_OVERRIDE;

    // MS VC4.0 messes things up
    void                    SetValue( sal_uInt16 nNewVal )
                                {SfxEnumItem::SetValue(nNewVal); }

    virtual bool            HasBoolValue() const SAL_OVERRIDE;
    virtual bool            GetBoolValue() const SAL_OVERRIDE;
    virtual void            SetBoolValue( bool bVal ) SAL_OVERRIDE;

    inline SvxCrossedOutItem& operator=(const SvxCrossedOutItem& rCross)
        {
            SetValue( rCross.GetValue() );
            return *this;
        }

    // enum cast
    FontStrikeout           GetStrikeout() const
                                { return (FontStrikeout)GetValue(); }
    void                    SetStrikeout( FontStrikeout eNew )
                                { SetValue( (sal_uInt16)eNew ); }
};

#endif // INCLUDED_EDITENG_CROSSEDOUTITEM_HXX

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
