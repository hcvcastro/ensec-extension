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

#ifndef INCLUDED_SOT_FILELIST_HXX
#define INCLUDED_SOT_FILELIST_HXX

#include <sot/sotdllapi.h>
#include <tools/stream.hxx>

#include <vector>

typedef ::std::vector< OUString > FileStringList;

class SOT_DLLPUBLIC FileList : public SvDataCopyStream
{
    FileStringList  aStrList;

protected:

    // SvData-Methoden
    virtual void    Load( SvStream& ) SAL_OVERRIDE;
    virtual void    Save( SvStream& ) SAL_OVERRIDE;
    virtual void    Assign( const SvDataCopyStream& ) SAL_OVERRIDE;

    // Liste loeschen;
    void            ClearAll();

public:

    TYPEINFO_OVERRIDE();
    FileList() {};
    virtual ~FileList();

    // Zuweisungsoperator
    FileList&           operator=( const FileList& rFileList );

    // Im-/Export
    SOT_DLLPUBLIC friend SvStream&  WriteFileList( SvStream& rOStm, const FileList& rFileList );
    SOT_DLLPUBLIC friend SvStream&  ReadFileList( SvStream& rIStm, FileList& rFileList );

    // Liste fuellen/abfragen
    void AppendFile( const OUString& rStr );
    OUString GetFile( size_t i ) const;
    size_t Count() const;

};

#endif // INCLUDED_SOT_FILELIST_HXX

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
