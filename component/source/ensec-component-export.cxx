/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*************************************************************************
 *
 * DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
 *
 * Copyright 2014, Bolvia@SHELL:$_ Henry Castro <hcvcastro@gmail.com>.
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

#ifndef _CPPUHELPER_IMPLEMENTATIONENTRY_
#include <cppuhelper/implementationentry.hxx>
#endif

#include <uno/lbnames.h>
#include <cppuhelper/weak.hxx>
#include <cppuhelper/factory.hxx>

#include "ensec-component-export.h"
#include "ensec-protocol-handler.h"


using namespace ::com::sun::star::uno;


// ---------------------------------------------------------------------------------------------
// get the implementation name
//
::rtl::OUString 
EnsecProtocolHandler_getImplementationName () 
  throw (RuntimeException)
{
  return ::rtl::OUString( RTL_CONSTASCII_USTRINGPARAM( ENSEC_PROTOCOLHANDLER_IMPLEMENTATIONNAME ) );
}


// ----------------------------------------------------------------------------------------------
// check wheter support specific service name
//
sal_Bool SAL_CALL 
EnsecProtocolHandler_supportsService( const ::rtl::OUString& ServiceName )
  throw (RuntimeException)
{
  return (
	  ServiceName.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM( ENSEC_PROTOCOLHANDLER_SERVICENAME ) ) ||
	  ServiceName.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM( "com.sun.star.frame.ProtocolHandler" ) )
	  );
}

// -----------------------------------------------------------------------------------------------
// get sequence list of supported service names
//
Sequence< ::rtl::OUString > SAL_CALL 
EnsecProtocolHandler_getSupportedServiceNames(  )
  throw (RuntimeException)
{
  Sequence < ::rtl::OUString > aRet(1);
  aRet[0] = ::rtl::OUString( RTL_CONSTASCII_USTRINGPARAM( ENSEC_PROTOCOLHANDLER_SERVICENAME ) );
  return aRet;
}

// ------------------------------------------------------------------------------------
// create an instance of flisol component
//
Reference< XInterface > SAL_CALL 
EnsecProtocolHandler_createInstance( const Reference< XComponentContext > & rContext )
throw( Exception )
{
  return (cppu::OWeakObject*) new EnsecProtocolHandler( rContext );
}


// ---------------------------------------------------------------------------------------
// Exported symbols
//
extern "C"
{
  static struct ::cppu::ImplementationEntry s_component_entries [] = {
    { EnsecProtocolHandler_createInstance, 
      EnsecProtocolHandler_getImplementationName, 
      EnsecProtocolHandler_getSupportedServiceNames, 
      ::cppu::createSingleComponentFactory, 0, 0 
    },
    { 0, 0, 0, 0, 0, 0 } // terminate with NULL
  };

  // ---------------------------------------------------------------------------
  // get the symbol factory
  //
  SAL_DLLPUBLIC_EXPORT void * SAL_CALL 
  component_getFactory(sal_Char const * implName, 
		       ::com::sun::star::lang::XMultiServiceFactory * xMgr, 
		       ::com::sun::star::registry::XRegistryKey * xRegistry )
  {
    return ::cppu::component_getFactoryHelper(implName, xMgr, xRegistry, s_component_entries );
  }

    // ----------------------------------------------------------------------------
    // get the symbol implementation environment
    //
    SAL_DLLPUBLIC_EXPORT void SAL_CALL 
    component_getImplementationEnvironment( const sal_Char **ppEnvTypeName,
                                            uno_Environment        ** /*ppEnv*/)
    {
        *ppEnvTypeName = CPPU_CURRENT_LANGUAGE_BINDING_NAME;
    }


};
