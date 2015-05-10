/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

/*
 This file has been autogenerated by update_pch.sh . It is possible to edit it
 manually (such as when an include file has been moved/renamed/removed. All such
 manual changes will be rewritten by the next run of update_pch.sh (which presumably
 also fixes all possible problems, so it's usually better to use it).
*/

#include <algorithm>
#include <boost/noncopyable.hpp>
#include <boost/scoped_array.hpp>
#include <boost/shared_ptr.hpp>
#include <boost/weak_ptr.hpp>
#include <cassert>
#include <com/sun/star/beans/NamedValue.hpp>
#include <com/sun/star/beans/Property.hpp>
#include <com/sun/star/beans/PropertyAttribute.hpp>
#include <com/sun/star/beans/PropertyChangeEvent.hpp>
#include <com/sun/star/beans/PropertyValue.hpp>
#include <com/sun/star/beans/PropertyVetoException.hpp>
#include <com/sun/star/beans/UnknownPropertyException.hpp>
#include <com/sun/star/beans/XFastPropertySet.hpp>
#include <com/sun/star/beans/XPropertyAccess.hpp>
#include <com/sun/star/beans/XPropertyChangeListener.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/beans/XPropertySetInfo.hpp>
#include <com/sun/star/beans/XVetoableChangeListener.hpp>
#include <com/sun/star/bridge/UnoUrlResolver.hpp>
#include <com/sun/star/bridge/XUnoUrlResolver.hpp>
#include <com/sun/star/connection/SocketPermission.hpp>
#include <com/sun/star/container/ElementExistException.hpp>
#include <com/sun/star/container/NoSuchElementException.hpp>
#include <com/sun/star/container/XEnumeration.hpp>
#include <com/sun/star/container/XHierarchicalNameAccess.hpp>
#include <com/sun/star/container/XNameContainer.hpp>
#include <com/sun/star/io/FilePermission.hpp>
#include <com/sun/star/lang/DisposedException.hpp>
#include <com/sun/star/lang/EventObject.hpp>
#include <com/sun/star/lang/IllegalAccessException.hpp>
#include <com/sun/star/lang/IllegalArgumentException.hpp>
#include <com/sun/star/lang/WrappedTargetException.hpp>
#include <com/sun/star/lang/WrappedTargetRuntimeException.hpp>
#include <com/sun/star/lang/XComponent.hpp>
#include <com/sun/star/lang/XEventListener.hpp>
#include <com/sun/star/lang/XInitialization.hpp>
#include <com/sun/star/lang/XMultiComponentFactory.hpp>
#include <com/sun/star/lang/XServiceInfo.hpp>
#include <com/sun/star/lang/XSingleComponentFactory.hpp>
#include <com/sun/star/lang/XSingleServiceFactory.hpp>
#include <com/sun/star/loader/CannotActivateFactoryException.hpp>
#include <com/sun/star/loader/XImplementationLoader.hpp>
#include <com/sun/star/reflection/InvalidTypeNameException.hpp>
#include <com/sun/star/reflection/NoSuchTypeNameException.hpp>
#include <com/sun/star/reflection/TypeDescriptionSearchDepth.hpp>
#include <com/sun/star/reflection/XCompoundTypeDescription.hpp>
#include <com/sun/star/reflection/XConstantTypeDescription.hpp>
#include <com/sun/star/reflection/XConstantsTypeDescription.hpp>
#include <com/sun/star/reflection/XEnumTypeDescription.hpp>
#include <com/sun/star/reflection/XIdlClass.hpp>
#include <com/sun/star/reflection/XIdlField2.hpp>
#include <com/sun/star/reflection/XIndirectTypeDescription.hpp>
#include <com/sun/star/reflection/XInterfaceAttributeTypeDescription2.hpp>
#include <com/sun/star/reflection/XInterfaceMemberTypeDescription.hpp>
#include <com/sun/star/reflection/XInterfaceMethodTypeDescription.hpp>
#include <com/sun/star/reflection/XInterfaceTypeDescription2.hpp>
#include <com/sun/star/reflection/XMethodParameter.hpp>
#include <com/sun/star/reflection/XModuleTypeDescription.hpp>
#include <com/sun/star/reflection/XPublished.hpp>
#include <com/sun/star/reflection/XServiceTypeDescription2.hpp>
#include <com/sun/star/reflection/XSingletonTypeDescription2.hpp>
#include <com/sun/star/reflection/XStructTypeDescription.hpp>
#include <com/sun/star/reflection/XTypeDescription.hpp>
#include <com/sun/star/reflection/theCoreReflection.hpp>
#include <com/sun/star/registry/CannotRegisterImplementationException.hpp>
#include <com/sun/star/registry/InvalidRegistryException.hpp>
#include <com/sun/star/security/RuntimePermission.hpp>
#include <com/sun/star/security/XAccessController.hpp>
#include <com/sun/star/uno/Any.hxx>
#include <com/sun/star/uno/DeploymentException.hpp>
#include <com/sun/star/uno/Exception.hpp>
#include <com/sun/star/uno/Reference.hxx>
#include <com/sun/star/uno/RuntimeException.hpp>
#include <com/sun/star/uno/Sequence.hxx>
#include <com/sun/star/uno/Type.hxx>
#include <com/sun/star/uno/TypeClass.hpp>
#include <com/sun/star/uno/XComponentContext.hpp>
#include <com/sun/star/uno/XInterface.hpp>
#include <com/sun/star/uno/XUnloadingPreference.hpp>
#include <com/sun/star/util/XMacroExpander.hpp>
#include <config_features.h>
#include <config_folders.h>
#include <cppu/unotype.hxx>
#include <cstddef>
#include <cstdlib>
#include <cstring>
#include <exception>
#include <map>
#include <memory>
#include <osl/diagnose.h>
#include <osl/doublecheckedlocking.h>
#include <osl/file.hxx>
#include <osl/module.h>
#include <osl/module.hxx>
#include <osl/mutex.hxx>
#include <osl/security.hxx>
#include <osl/thread.hxx>
#include <registry/registry.hxx>
#include <rtl/alloc.h>
#include <rtl/bootstrap.hxx>
#include <rtl/byteseq.hxx>
#include <rtl/instance.hxx>
#include <rtl/malformeduriexception.hxx>
#include <rtl/process.h>
#include <rtl/random.h>
#include <rtl/ref.hxx>
#include <rtl/strbuf.hxx>
#include <rtl/string.h>
#include <rtl/string.hxx>
#include <rtl/textenc.h>
#include <rtl/unload.h>
#include <rtl/uri.h>
#include <rtl/uri.hxx>
#include <rtl/ustrbuf.hxx>
#include <rtl/ustring.h>
#include <rtl/ustring.hxx>
#include <rtl/uuid.h>
#include <sal/alloca.h>
#include <sal/config.h>
#include <sal/log.hxx>
#include <sal/macros.h>
#include <sal/types.h>
#include <salhelper/simplereferenceobject.hxx>
#include <set>
#include <stack>
#include <string.h>
#include <typelib/typedescription.h>
#include <uno/dispatcher.hxx>
#include <uno/environment.hxx>
#include <uno/lbnames.h>
#include <uno/mapping.hxx>
#include <unoidl/unoidl.hxx>
#include <unordered_map>
#include <vector>

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
