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

#ifndef INCLUDED_UCBHELPER_COMMANDENVIRONMENT_HXX
#define INCLUDED_UCBHELPER_COMMANDENVIRONMENT_HXX

#include <com/sun/star/ucb/XCommandEnvironment.hpp>
#include <ucbhelper/ucbhelperdllapi.h>
#include <cppuhelper/implbase1.hxx>

namespace ucbhelper
{
struct CommandEnvironment_Impl;

/**
  * This class implements the interface
  * com::sun::star::ucb::XCommandEnvironement. Instances of this class can
  * be used to supply environments to commands executed by UCB contents.
  */
class UCBHELPER_DLLPUBLIC CommandEnvironment :
            public cppu::WeakImplHelper1< css::ucb::XCommandEnvironment >
{
    CommandEnvironment_Impl* m_pImpl;

private:
    CommandEnvironment( const CommandEnvironment& ) SAL_DELETED_FUNCTION;
    CommandEnvironment& operator=( const CommandEnvironment& ) SAL_DELETED_FUNCTION;

public:
    /**
      * Constructor.
      *
      * @param rxInteractionHandler is the implementation of an Interaction
      *        Handler or an empty reference.
      * @param rxProgressHandler is the implementation of a Progress
      *        Handler or an empty reference.
      */
    CommandEnvironment( const css::uno::Reference< css::task::XInteractionHandler >& rxInteractionHandler,
                        const css::uno::Reference< css::ucb::XProgressHandler >& rxProgressHandler );
    /**
      * Destructor.
      */
    virtual ~CommandEnvironment();

     // XCommandEnvironemnt
    virtual css::uno::Reference< css::task::XInteractionHandler > SAL_CALL getInteractionHandler()
        throw ( css::uno::RuntimeException,
                std::exception ) SAL_OVERRIDE;

    virtual css::uno::Reference< css::ucb::XProgressHandler > SAL_CALL getProgressHandler()
        throw ( css::uno::RuntimeException,
                std::exception ) SAL_OVERRIDE;
};

} /* namespace ucbhelper */

#endif /* ! INCLUDED_UCBHELPER_COMMANDENVIRONMENT_HXX */

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
