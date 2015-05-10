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
#ifndef INCLUDED_CPPUHELPER_IMPLBASE10_HXX
#define INCLUDED_CPPUHELPER_IMPLBASE10_HXX

#include <cppuhelper/implbase_ex.hxx>
#include <rtl/instance.hxx>

namespace cppu
{
    /// @cond INTERNAL

    struct class_data10
    {
        sal_Int16 m_nTypes;
        sal_Bool m_storedTypeRefs;
        sal_Bool m_storedId;
        sal_Int8 m_id[ 16 ];
        type_entry m_typeEntries[ 10 + 1 ];
    };

    template< typename Ifc1, typename Ifc2, typename Ifc3, typename Ifc4, typename Ifc5, typename Ifc6, typename Ifc7, typename Ifc8, typename Ifc9, typename Ifc10, typename Impl >
    struct ImplClassData10
    {
        class_data* operator ()()
        {
            static class_data10 s_cd =
            {
                10 +1, sal_False, sal_False,
                { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 },
                {
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc1),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc2),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc3),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc4),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc5),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc6),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc7),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc8),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc9),
                    CPPUHELPER_DETAIL_TYPEENTRY(Ifc10),
                    CPPUHELPER_DETAIL_TYPEENTRY(com::sun::star::lang::XTypeProvider)
                }
            };
            return reinterpret_cast< class_data * >(&s_cd);
        }
    };

    /// @endcond

    /** Implementation helper implementing interface com::sun::star::lang::XTypeProvider
        and method XInterface::queryInterface(), but no reference counting.

        @derive
        Inherit from this class giving your interface(s) to be implemented as template argument(s).
        Your sub class defines method implementations for these interface(s) including acquire()/
        release() and delegates incoming queryInterface() calls to this base class.
    */
    template< class Ifc1, class Ifc2, class Ifc3, class Ifc4, class Ifc5, class Ifc6, class Ifc7, class Ifc8, class Ifc9, class Ifc10 >
    class SAL_NO_VTABLE SAL_DLLPUBLIC_TEMPLATE ImplHelper10
        : public com::sun::star::lang::XTypeProvider
        , public Ifc1, public Ifc2, public Ifc3, public Ifc4, public Ifc5, public Ifc6, public Ifc7, public Ifc8, public Ifc9, public Ifc10
    {
        struct cd : public rtl::StaticAggregate< class_data, ImplClassData10< Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10, ImplHelper10<Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10> > > {};
    public:
        virtual com::sun::star::uno::Any SAL_CALL queryInterface( com::sun::star::uno::Type const & rType ) throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return ImplHelper_query( rType, cd::get(), this ); }
        virtual com::sun::star::uno::Sequence< com::sun::star::uno::Type > SAL_CALL getTypes() throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return ImplHelper_getTypes( cd::get() ); }
        virtual com::sun::star::uno::Sequence< sal_Int8 > SAL_CALL getImplementationId() throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return ImplHelper_getImplementationId( cd::get() ); }

#if !defined _MSC_VER // public -> protected changes mangled names there
    protected:
#endif
        ~ImplHelper10() throw () {}
    };
    /** Implementation helper implementing interfaces com::sun::star::lang::XTypeProvider and
        com::sun::star::uno::XInterface which supports weak mechanism to be held weakly
        (supporting com::sun::star::uno::XWeak through ::cppu::OWeakObject).

        @derive
        Inherit from this class giving your interface(s) to be implemented as template argument(s).
        Your sub class defines method implementations for these interface(s).
    */
    template< class Ifc1, class Ifc2, class Ifc3, class Ifc4, class Ifc5, class Ifc6, class Ifc7, class Ifc8, class Ifc9, class Ifc10 >
    class SAL_NO_VTABLE SAL_DLLPUBLIC_TEMPLATE WeakImplHelper10
        : public OWeakObject
        , public com::sun::star::lang::XTypeProvider
        , public Ifc1, public Ifc2, public Ifc3, public Ifc4, public Ifc5, public Ifc6, public Ifc7, public Ifc8, public Ifc9, public Ifc10
    {
        struct cd : public rtl::StaticAggregate< class_data, ImplClassData10< Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10, WeakImplHelper10<Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10> > > {};
    public:
        virtual com::sun::star::uno::Any SAL_CALL queryInterface( com::sun::star::uno::Type const & rType ) throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return WeakImplHelper_query( rType, cd::get(), this, static_cast<OWeakObject *>(this) ); }
        virtual void SAL_CALL acquire() throw () SAL_OVERRIDE
            { OWeakObject::acquire(); }
        virtual void SAL_CALL release() throw () SAL_OVERRIDE
            { OWeakObject::release(); }
        virtual com::sun::star::uno::Sequence< com::sun::star::uno::Type > SAL_CALL getTypes() throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return WeakImplHelper_getTypes( cd::get() ); }
        virtual com::sun::star::uno::Sequence< sal_Int8 > SAL_CALL getImplementationId() throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return ImplHelper_getImplementationId( cd::get() ); }
    };
    /** Implementation helper implementing interfaces com::sun::star::lang::XTypeProvider and
        com::sun::star::uno::XInterface which supports weak mechanism to be held weakly
        (supporting com::sun::star::uno::XWeak through ::cppu::OWeakAggObject).
        In addition, it supports also aggregation meaning object of this class can be aggregated
        (com::sun::star::uno::XAggregation through ::cppu::OWeakAggObject).
        If a delegator is set (this object is aggregated), then incoming queryInterface()
        calls are delegated to the delegator object. If the delegator does not support the
        demanded interface, it calls queryAggregation() on its aggregated objects.

        @derive
        Inherit from this class giving your interface(s) to be implemented as template argument(s).
        Your sub class defines method implementations for these interface(s).
    */
    template< class Ifc1, class Ifc2, class Ifc3, class Ifc4, class Ifc5, class Ifc6, class Ifc7, class Ifc8, class Ifc9, class Ifc10 >
    class SAL_NO_VTABLE SAL_DLLPUBLIC_TEMPLATE WeakAggImplHelper10
        : public OWeakAggObject
        , public com::sun::star::lang::XTypeProvider
        , public Ifc1, public Ifc2, public Ifc3, public Ifc4, public Ifc5, public Ifc6, public Ifc7, public Ifc8, public Ifc9, public Ifc10
    {
        struct cd : public rtl::StaticAggregate< class_data, ImplClassData10< Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10, WeakAggImplHelper10<Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10> > > {};
    public:
        virtual com::sun::star::uno::Any SAL_CALL queryInterface( com::sun::star::uno::Type const & rType ) throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return OWeakAggObject::queryInterface( rType ); }
        virtual com::sun::star::uno::Any SAL_CALL queryAggregation( com::sun::star::uno::Type const & rType ) throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return WeakAggImplHelper_queryAgg( rType, cd::get(), this, static_cast<OWeakAggObject *>(this) ); }
        virtual void SAL_CALL acquire() throw () SAL_OVERRIDE
            { OWeakAggObject::acquire(); }
        virtual void SAL_CALL release() throw () SAL_OVERRIDE
            { OWeakAggObject::release(); }
        virtual com::sun::star::uno::Sequence< com::sun::star::uno::Type > SAL_CALL getTypes() throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return WeakAggImplHelper_getTypes( cd::get() ); }
        virtual com::sun::star::uno::Sequence< sal_Int8 > SAL_CALL getImplementationId() throw (com::sun::star::uno::RuntimeException, std::exception) SAL_OVERRIDE
            { return ImplHelper_getImplementationId( cd::get() ); }
    };
    /** Implementation helper implementing interfaces com::sun::star::lang::XTypeProvider and
        com::sun::star::uno::XInterface inherting from a BaseClass.
        All acquire() and release() calls are delegated to the BaseClass. Upon queryInterface(),
        if a demanded interface is not supported by this class directly, the request is
        delegated to the BaseClass.

        @attention
        The BaseClass has to be complete in a sense, that com::sun::star::uno::XInterface
        and com::sun::star::lang::XTypeProvider are implemented properly.  The
        BaseClass must have at least one ctor that can be called with six or
        fewer arguments, of which none is of non-const reference type.

        @derive
        Inherit from this class giving your additional interface(s) to be implemented as
        template argument(s). Your sub class defines method implementations for these interface(s).
    */
    template< class BaseClass, class Ifc1, class Ifc2, class Ifc3, class Ifc4, class Ifc5, class Ifc6, class Ifc7, class Ifc8, class Ifc9, class Ifc10 >
    class SAL_NO_VTABLE SAL_DLLPUBLIC_TEMPLATE ImplInheritanceHelper10
        : public BaseClass
        , public Ifc1, public Ifc2, public Ifc3, public Ifc4, public Ifc5, public Ifc6, public Ifc7, public Ifc8, public Ifc9, public Ifc10
    {
        struct cd : public rtl::StaticAggregate< class_data, ImplClassData10< Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10, ImplInheritanceHelper10<BaseClass, Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10> > > {};
    protected:
        template< typename T1 >
        explicit ImplInheritanceHelper10(T1 const & arg1): BaseClass(arg1) {}
        template< typename T1, typename T2 >
        ImplInheritanceHelper10(T1 const & arg1, T2 const & arg2):
            BaseClass(arg1, arg2) {}
        template< typename T1, typename T2, typename T3 >
        ImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3):
            BaseClass(arg1, arg2, arg3) {}
        template< typename T1, typename T2, typename T3, typename T4 >
        ImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3, T4 const & arg4):
            BaseClass(arg1, arg2, arg3, arg4) {}
        template<
            typename T1, typename T2, typename T3, typename T4, typename T5 >
        ImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3, T4 const & arg4,
            T5 const & arg5):
            BaseClass(arg1, arg2, arg3, arg4, arg5) {}
        template<
            typename T1, typename T2, typename T3, typename T4, typename T5,
            typename T6 >
        ImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3, T4 const & arg4,
            T5 const & arg5, T6 const & arg6):
            BaseClass(arg1, arg2, arg3, arg4, arg5, arg6) {}
    public:
        ImplInheritanceHelper10() {}
        virtual com::sun::star::uno::Any SAL_CALL queryInterface( com::sun::star::uno::Type const & rType ) throw (com::sun::star::uno::RuntimeException) SAL_OVERRIDE
            {
                com::sun::star::uno::Any aRet( ImplHelper_queryNoXInterface( rType, cd::get(), this ) );
                if (aRet.hasValue())
                    return aRet;
                return BaseClass::queryInterface( rType );
            }
        virtual void SAL_CALL acquire() throw () SAL_OVERRIDE
            { BaseClass::acquire(); }
        virtual void SAL_CALL release() throw () SAL_OVERRIDE
            { BaseClass::release(); }
        virtual com::sun::star::uno::Sequence< com::sun::star::uno::Type > SAL_CALL getTypes() throw (com::sun::star::uno::RuntimeException) SAL_OVERRIDE
            { return ImplInhHelper_getTypes( cd::get(), BaseClass::getTypes() ); }
        virtual com::sun::star::uno::Sequence< sal_Int8 > SAL_CALL getImplementationId() throw (com::sun::star::uno::RuntimeException) SAL_OVERRIDE
            { return ImplHelper_getImplementationId( cd::get() ); }
    };
    /** Implementation helper implementing interfaces com::sun::star::lang::XTypeProvider and
        com::sun::star::uno::XInterface inherting from a BaseClass.
        All acquire(),  release() and queryInterface() calls are delegated to the BaseClass.
        Upon queryAggregation(), if a demanded interface is not supported by this class directly,
        the request is delegated to the BaseClass.

        @attention
        The BaseClass has to be complete in a sense, that com::sun::star::uno::XInterface,
        com::sun::star::uno::XAggregation and com::sun::star::lang::XTypeProvider
        are implemented properly.  The BaseClass must have at least one ctor
        that can be called with six or fewer arguments, of which none is of
        non-const reference type.

        @derive
        Inherit from this class giving your additional interface(s) to be implemented as
        template argument(s). Your sub class defines method implementations for these interface(s).
    */
    template< class BaseClass, class Ifc1, class Ifc2, class Ifc3, class Ifc4, class Ifc5, class Ifc6, class Ifc7, class Ifc8, class Ifc9, class Ifc10 >
    class SAL_NO_VTABLE SAL_DLLPUBLIC_TEMPLATE AggImplInheritanceHelper10
        : public BaseClass
        , public Ifc1, public Ifc2, public Ifc3, public Ifc4, public Ifc5, public Ifc6, public Ifc7, public Ifc8, public Ifc9, public Ifc10
    {
        struct cd : public rtl::StaticAggregate< class_data, ImplClassData10< Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10, AggImplInheritanceHelper10<BaseClass, Ifc1, Ifc2, Ifc3, Ifc4, Ifc5, Ifc6, Ifc7, Ifc8, Ifc9, Ifc10> > > {};
    protected:
        template< typename T1 >
        explicit AggImplInheritanceHelper10(T1 const & arg1): BaseClass(arg1) {}
        template< typename T1, typename T2 >
        AggImplInheritanceHelper10(T1 const & arg1, T2 const & arg2):
            BaseClass(arg1, arg2) {}
        template< typename T1, typename T2, typename T3 >
        AggImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3):
            BaseClass(arg1, arg2, arg3) {}
        template< typename T1, typename T2, typename T3, typename T4 >
        AggImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3, T4 const & arg4):
            BaseClass(arg1, arg2, arg3, arg4) {}
        template<
            typename T1, typename T2, typename T3, typename T4, typename T5 >
        AggImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3, T4 const & arg4,
            T5 const & arg5):
            BaseClass(arg1, arg2, arg3, arg4, arg5) {}
        template<
            typename T1, typename T2, typename T3, typename T4, typename T5,
            typename T6 >
        AggImplInheritanceHelper10(
            T1 const & arg1, T2 const & arg2, T3 const & arg3, T4 const & arg4,
            T5 const & arg5, T6 const & arg6):
            BaseClass(arg1, arg2, arg3, arg4, arg5, arg6) {}
    public:
        AggImplInheritanceHelper10() {}
        virtual com::sun::star::uno::Any SAL_CALL queryInterface( com::sun::star::uno::Type const & rType ) throw (com::sun::star::uno::RuntimeException) SAL_OVERRIDE
            { return BaseClass::queryInterface( rType ); }
        virtual com::sun::star::uno::Any SAL_CALL queryAggregation( com::sun::star::uno::Type const & rType ) throw (com::sun::star::uno::RuntimeException) SAL_OVERRIDE
            {
                com::sun::star::uno::Any aRet( ImplHelper_queryNoXInterface( rType, cd::get(), this ) );
                if (aRet.hasValue())
                    return aRet;
                return BaseClass::queryAggregation( rType );
            }
        virtual void SAL_CALL acquire() throw () SAL_OVERRIDE
            { BaseClass::acquire(); }
        virtual void SAL_CALL release() throw () SAL_OVERRIDE
            { BaseClass::release(); }
        virtual com::sun::star::uno::Sequence< com::sun::star::uno::Type > SAL_CALL getTypes() throw (com::sun::star::uno::RuntimeException) SAL_OVERRIDE
            { return ImplInhHelper_getTypes( cd::get(), BaseClass::getTypes() ); }
        virtual com::sun::star::uno::Sequence< sal_Int8 > SAL_CALL getImplementationId() throw (com::sun::star::uno::RuntimeException) SAL_OVERRIDE
            { return ImplHelper_getImplementationId( cd::get() ); }
    };
}

#endif

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
