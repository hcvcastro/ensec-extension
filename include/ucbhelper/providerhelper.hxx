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

#ifndef INCLUDED_UCBHELPER_PROVIDERHELPER_HXX
#define INCLUDED_UCBHELPER_PROVIDERHELPER_HXX

#include <list>
#include <com/sun/star/ucb/XContentProvider.hpp>
#include <com/sun/star/lang/XServiceInfo.hpp>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/lang/XTypeProvider.hpp>
#include <cppuhelper/weak.hxx>

#include <osl/mutex.hxx>
#include <rtl/ref.hxx>
#include <ucbhelper/macros.hxx>
#include <ucbhelper/ucbhelperdllapi.h>



namespace com { namespace sun { namespace star { namespace ucb {
    class XPropertySetRegistry;
    class XPersistentPropertySet;
} } } }

namespace ucbhelper_impl { struct ContentProviderImplHelper_Impl; }

namespace ucbhelper {



class ContentImplHelper;
typedef rtl::Reference< ContentImplHelper > ContentImplHelperRef;
typedef std::list< ContentImplHelperRef > ContentRefList;

/**
  * This is an abstract base class for implementations of the service
  * com.sun.star.ucb.ContentProvider. It provides contents derived from
  * class ucb::ContentImplHelper.
  *
  * Features of the base class implementation:
  * - standard interfaces ( XInterface, XTypeProvider, XServiceInfo )
  * - maintains a set of ContentImplHelper objects, which were created by
  *   the provider implementation. So there will be exactly one object for
  *   one Content Identifier.
  * - Provides access to the Additional Core PropertySet of a content.
  *   ( These set contains the properties added to a content using its
  *   XPropertyContainer interface )
  */
class UCBHELPER_DLLPUBLIC ContentProviderImplHelper : public cppu::OWeakObject,
                                    public com::sun::star::lang::XTypeProvider,
                                    public com::sun::star::lang::XServiceInfo,
                                    public com::sun::star::ucb::XContentProvider
{
    friend class ContentImplHelper;

    ucbhelper_impl::ContentProviderImplHelper_Impl* m_pImpl;

protected:
    osl::Mutex m_aMutex;
    ::com::sun::star::uno::Reference< ::com::sun::star::uno::XComponentContext >  m_xContext;

private:
    UCBHELPER_DLLPRIVATE void removeContent( ContentImplHelper* pContent );

    UCBHELPER_DLLPRIVATE ::com::sun::star::uno::Reference<
        ::com::sun::star::ucb::XPropertySetRegistry >
    getAdditionalPropertySetRegistry();

    UCBHELPER_DLLPRIVATE void cleanupRegisteredContents();

protected:
    /**
      * This method returns a content with the given id, if it already exists.
      * Use this method in your "queryContent" implementation to ensure unique
      * objects.
      *
      * @param  Identifier is the content identifier, for that an existing
      *         content object is requested.
      * @return the content with the given identifier, if it exists or 0, if it
      *         does not exist.
      */
    rtl::Reference< ContentImplHelper >
    queryExistingContent( const ::com::sun::star::uno::Reference<
                   ::com::sun::star::ucb::XContentIdentifier >& Identifier );

    /**
      * This method returns a content with the given URL, if it already exists.
      *
      * @param  rURL is the URL ( content identifier string ), for that an
      *         existing content object is requested.
      * @return the content with the given URL, if it exists or 0, if it
      *         does not exist.
      */
    rtl::Reference< ContentImplHelper >
    queryExistingContent( const OUString& rURL );

    /**
      * This method registers a newly created content instance with the
      * content provider. It should be called directly after creating a new
      * content instance. The provider can reuse a registered instance upon
      * subsedent requests for content instances with an idententifier
      * of a registered instance.
      * Note that the provider does not hold a hard reference on the
      * registered instance. If last external reference is gone, the provider
      * will remove the instance from its inventory of known instances.
      * Nothing will happen in case an already registered instance shall
      * be registered more than once.
      *
      * @param  the content instance that is to be registered.
     */
    void registerNewContent(
        const com::sun::star::uno::Reference<
            ::com::sun::star::ucb::XContent > & xContent );

public:


    // Contsruction/Destruction


    ContentProviderImplHelper(
                const ::com::sun::star::uno::Reference<
                    ::com::sun::star::uno::XComponentContext >& rxContext );
    virtual ~ContentProviderImplHelper();


    // XInterface
    virtual css::uno::Any SAL_CALL queryInterface( const css::uno::Type & rType )
        throw( css::uno::RuntimeException, std::exception ) SAL_OVERRIDE;
    virtual void SAL_CALL acquire()
        throw() SAL_OVERRIDE;
    virtual void SAL_CALL release()
        throw() SAL_OVERRIDE;

    // XTypeProvider


    virtual com::sun::star::uno::Sequence< sal_Int8 > SAL_CALL
    getImplementationId()
        throw( com::sun::star::uno::RuntimeException, std::exception ) SAL_OVERRIDE;
    virtual com::sun::star::uno::Sequence< com::sun::star::uno::Type > SAL_CALL
    getTypes()
        throw( com::sun::star::uno::RuntimeException, std::exception ) SAL_OVERRIDE;

    // XServiceInfo


    virtual OUString SAL_CALL
    getImplementationName()
        throw( ::com::sun::star::uno::RuntimeException, std::exception ) SAL_OVERRIDE = 0;
    virtual sal_Bool SAL_CALL
    supportsService( const OUString& ServiceName )
        throw( ::com::sun::star::uno::RuntimeException, std::exception ) SAL_OVERRIDE;
    virtual ::com::sun::star::uno::Sequence< OUString > SAL_CALL
    getSupportedServiceNames()
        throw( ::com::sun::star::uno::RuntimeException, std::exception ) SAL_OVERRIDE = 0;


    // XContentProvider


    /**
      * This method returns a content with the requested id.
      *
      * The implementation should:
      *
      * - Check, whether the Identifier is valid ( URL syntax ).
      * - Use queryExistingContent(...) to determine, whether there exists
      *   already a content with the given id.
      * - Return the possibly existing content.Create and return a new
      *   content, otherwise
      */
    virtual ::com::sun::star::uno::Reference<
                ::com::sun::star::ucb::XContent > SAL_CALL
    queryContent( const ::com::sun::star::uno::Reference<
                    ::com::sun::star::ucb::XContentIdentifier >& Identifier )
        throw( ::com::sun::star::ucb::IllegalIdentifierException,
               ::com::sun::star::uno::RuntimeException, std::exception ) SAL_OVERRIDE = 0;
    virtual sal_Int32 SAL_CALL
    compareContentIds( const ::com::sun::star::uno::Reference<
                            ::com::sun::star::ucb::XContentIdentifier >& Id1,
                       const ::com::sun::star::uno::Reference<
                               ::com::sun::star::ucb::XContentIdentifier >& Id2 )
        throw( ::com::sun::star::uno::RuntimeException, std::exception ) SAL_OVERRIDE;


    // Non-interface methods.


    /**
      * This method returns a mutex, which protects the content list of the
      * provider. So you can prevent modifications of that list easily.
      *
      * @return the mutex.
      */
    osl::Mutex& getContentListMutex() { return m_aMutex; }

    /**
      * This method fills a list with all contents existing at calling time.
      * Note: You may prevent modifications of the content list at any time
      * by acquiring the content list mutex of the provider.
      *
      * @param  rContents is the list to fill with the children.
      */
    void queryExistingContents( ContentRefList& rContents );

    /**
      * This method returns the propertyset containing the Additional Core
      * Properties of a content.
      *
      * @param  rKey is the key for the propertyset.
      * @param  bCreate is a flag indicating whether the propertyset shall
      *         be created in case it does not exist.
      * @return the propertyset containing the Additional Core Properties.
      */
    ::com::sun::star::uno::Reference<
        com::sun::star::ucb::XPersistentPropertySet >
    getAdditionalPropertySet( const OUString& rKey, bool bCreate );

    /**
      * This method renames the propertyset containing the Additional Core
      * Properties of a content.
      *
      * @param  rOldKey is the old key of the propertyset.
      * @param  rNewKey is the new key for the propertyset.
      * @param  bRecursive is a flag indicating whether propertysets for
      *         children described by rOldKey shall be renamed, too.
      * @return True, if the operation succeeded - False, otherwise.
      */
    bool renameAdditionalPropertySet( const OUString& rOldKey,
                                          const OUString& rNewKey,
                                          bool bRecursive );

    /**
      * This method copies the propertyset containing the Additional Core
      * Properties of a content.
      *
      * @param  rSourceKey is the key of the source propertyset.
      * @param  rTargetKey is the key of the target propertyset.
      * @param  bRecursive is a flag indicating whether propertysets for
      *         children described by rSourceKey shall be copied, too.
      * @return True, if the operation succeeded - False, otherwise.
      */
    bool copyAdditionalPropertySet( const OUString& rSourceKey,
                                        const OUString& rTargetKey,
                                        bool bRecursive );

    /**
      * This method removes the propertyset containing the Additional Core
      * Properties of a content.
      *
      * @param  rKey is the key of the propertyset.
      * @param  bRecursive is a flag indicating whether propertysets for
      *         children described by rOldKey shall be removed, too.
      * @return True, if the operation succeeded - False, otherwise.
      */
    bool removeAdditionalPropertySet( const OUString& rKey,
                                          bool bRecursive );
};

} // namespace ucbhelper

#endif /* ! INCLUDED_UCBHELPER_PROVIDERHELPER_HXX */

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
