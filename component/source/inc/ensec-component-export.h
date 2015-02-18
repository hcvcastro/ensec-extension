/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*************************************************************************
 *
 * DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
 *
 * Copyright 2014, Bolivia@SHELL:$_ Henry Castro <hcvcastro@gmail.com>.
 *
 * OpenOffice.org - a multi-platform office productivity suite
 *
 * This file is part of Flisol Component.
 *
 * OpenOffice.org is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License version 3
 * only, as published by the Free Software Foundation.
 *
 * OpenOffice.org is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License version 3 for more details
 * (a copy is included in the LICENSE file that accompanied this code).
 *
 * You should have received a copy of the GNU Lesser General Public License
 * version 3 along with OpenOffice.org.  If not, see
 * <http://www.openoffice.org/license.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/

#ifndef _FLISOL_PROTOCOLHANDLER_EXPORT_
#define _FLISOL_PROTOCOLHANDLER_EXPORT_

#include <cppuhelper/component.hxx>

#define ENSEC_PROTOCOLHANDLER_IMPLEMENTATIONNAME "bolivia.shell.ensec.EnsecProtocolHandler"
#define ENSEC_PROTOCOLHANDLER_SERVICENAME "bolivia.shell.ensec.ProtocolHandler"

using namespace ::com::sun::star::uno;

// --------------------------------------------------------------------------------
// get the implementation name
//
::rtl::OUString EnsecProtocolHandler_getImplementationName( )
throw (RuntimeException);

// ------------------------------------------------------------------------------------
// check whether support specific service
//
sal_Bool SAL_CALL EnsecProtocolHandler_supportsService( const ::rtl::OUString& )
throw (RuntimeException);

// ------------------------------------------------------------------------------------
// get a sequences of supported services names
//
Sequence< ::rtl::OUString > SAL_CALL EnsecProtocolHandler_getSupportedServiceNames( )
throw (RuntimeException);

#endif
