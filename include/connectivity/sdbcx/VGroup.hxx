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

#ifndef INCLUDED_CONNECTIVITY_SDBCX_VGROUP_HXX
#define INCLUDED_CONNECTIVITY_SDBCX_VGROUP_HXX

#include <osl/diagnose.h>


#include <com/sun/star/sdbcx/XUsersSupplier.hpp>
#include <com/sun/star/sdbcx/XAuthorizable.hpp>
#include <com/sun/star/container/XNamed.hpp>
#include <comphelper/proparrhlp.hxx>
#include <cppuhelper/compbase4.hxx>
#include <comphelper/broadcasthelper.hxx>
#include <connectivity/sdbcx/VCollection.hxx>
#include <comphelper/propertycontainer.hxx>
#include <connectivity/sdbcx/IRefreshable.hxx>
#include <connectivity/sdbcx/VDescriptor.hxx>
#include <connectivity/dbtoolsdllapi.hxx>
#include <com/sun/star/lang/XServiceInfo.hpp>

namespace connectivity
{
    namespace sdbcx
    {
        typedef OCollection OUsers;

        typedef ::cppu::WeakComponentImplHelper4<   ::com::sun::star::sdbcx::XUsersSupplier,
                                                    ::com::sun::star::sdbcx::XAuthorizable,
                                                    ::com::sun::star::container::XNamed,
                                                    ::com::sun::star::lang::XServiceInfo> OGroup_BASE;

        class OOO_DLLPUBLIC_DBTOOLS OGroup :
                        public comphelper::OBaseMutex,
                        public OGroup_BASE,
                        public IRefreshableUsers,
                        public ::comphelper::OPropertyArrayUsageHelper<OGroup>,
                        public ODescriptor
        {
        protected:
            OUsers*         m_pUsers;

            using OGroup_BASE::rBHelper;

            // OPropertyArrayUsageHelper
            virtual ::cppu::IPropertyArrayHelper* createArrayHelper( ) const SAL_OVERRIDE;
            // OPropertySetHelper
            virtual ::cppu::IPropertyArrayHelper & SAL_CALL getInfoHelper() SAL_OVERRIDE;
        public:
            OGroup(bool _bCase);
            OGroup( const OUString& _Name, bool _bCase);
            virtual ~OGroup();
            DECLARE_SERVICE_INFO();

            // XInterface
            virtual void SAL_CALL acquire() throw() SAL_OVERRIDE;
            virtual void SAL_CALL release() throw() SAL_OVERRIDE;

            //XInterface
            virtual ::com::sun::star::uno::Any SAL_CALL queryInterface( const ::com::sun::star::uno::Type & rType ) throw(::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
            //XTypeProvider
            virtual ::com::sun::star::uno::Sequence< ::com::sun::star::uno::Type > SAL_CALL getTypes(  ) throw(::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;

            // ::cppu::OComponentHelper
            virtual void SAL_CALL disposing() SAL_OVERRIDE;
            // XPropertySet
            virtual ::com::sun::star::uno::Reference< ::com::sun::star::beans::XPropertySetInfo > SAL_CALL getPropertySetInfo(  ) throw(::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
            // XUsersSupplier
            virtual ::com::sun::star::uno::Reference< ::com::sun::star::container::XNameAccess > SAL_CALL getUsers(  ) throw(::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
            // XAuthorizable
            virtual sal_Int32 SAL_CALL getPrivileges( const OUString& objName, sal_Int32 objType ) throw(::com::sun::star::sdbc::SQLException, ::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
            virtual sal_Int32 SAL_CALL getGrantablePrivileges( const OUString& objName, sal_Int32 objType ) throw(::com::sun::star::sdbc::SQLException, ::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
            virtual void SAL_CALL grantPrivileges( const OUString& objName, sal_Int32 objType, sal_Int32 objPrivileges ) throw(::com::sun::star::sdbc::SQLException, ::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
            virtual void SAL_CALL revokePrivileges( const OUString& objName, sal_Int32 objType, sal_Int32 objPrivileges ) throw(::com::sun::star::sdbc::SQLException, ::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;

            // XNamed
            virtual OUString SAL_CALL getName(  ) throw(::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
            virtual void SAL_CALL setName( const OUString& aName ) throw(::com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE;
        };
    }
}

#endif // INCLUDED_CONNECTIVITY_SDBCX_VGROUP_HXX


/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
