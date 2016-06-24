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

#ifndef _FLISOL_PROTOCOL_HANDLER_HXX
#define _FLISOL_PROTOCOL_HANDLER_HXX

#ifndef _CPPUHELPER_IMPLBASE4_HXX_
#include <cppuhelper/implbase4.hxx>
#endif

#ifndef _COM_SUN_STAR_FRAME_XFRAME_HPP_
#include <com/sun/star/frame/XFrame.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_TOOLKIT_HPP_
#include <com/sun/star/awt/Toolkit.hpp>
#endif

#ifndef _COM_SUN_STAR_LANG_XINITIALIZATION_HPP_
#include <com/sun/star/lang/XInitialization.hpp>
#endif

#ifndef _COM_SUN_STAR_FRAME_XDISPATCHPROVIDER_HPP_
#include <com/sun/star/frame/XDispatchProvider.hpp>
#endif

#ifndef _COM_SUN_STAR_LANG_XSERVICEINFO_HPP_
#include <com/sun/star/lang/XServiceInfo.hpp>
#endif

#ifndef _COM_SUN_STAR_BEANS_PROPERTYVALUE_HPP_
#include <com/sun/star/beans/PropertyValue.hpp>
#endif

#ifndef _COM_SUN_STAR_SHEET_XSPREADSHEETDOCUMENT_HPP_
#include "com/sun/star/sheet/XSpreadsheetDocument.hpp"
#endif

#ifndef _COM_SUN_STAR_SDBC_XDATASOURCE_HPP_
#include "com/sun/star/sdbc/XDataSource.hpp"
#endif

#ifndef _COM_SUN_STAR_SHEET_XSPREADSHEET_HPP_
#include "com/sun/star/sheet/XSpreadsheet.hpp"
#endif

#ifndef _COM_SUN_STAR_SDBC_XROW_HPP_
#include <com/sun/star/sdbc/XRow.hpp>
#endif

#ifndef _COM_SUN_STAR_SDB_XDOCUMENTDATASOURCE_HPP_
#include "com/sun/star/sdb/XDocumentDataSource.hpp"
#endif

#ifndef _COM_SUN_STAR_UNO_XCOMPONENTCONTEXT_HPP_
#include <com/sun/star/uno/XComponentContext.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XMESSAGEBOX_HPP_
#include <com/sun/star/awt/XMessageBox.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_MESSAGEBOXBUTTONS_HPP_
#include <com/sun/star/awt/MessageBoxButtons.hpp>
#endif

#define FILE_OPEN_SERVICE_NAME_OOO  OUString("com.sun.star.ui.dialogs.OfficeFilePicker")
#define FILE_PICKER                 OUString("com.sun.star.ui.dialogs.FilePicker")


using namespace ::com::sun::star::uno;
using namespace ::rtl;

class EnsecProtocolHandler : public cppu::WeakImplHelper4
<
  com::sun::star::lang::XInitialization,
  com::sun::star::lang::XServiceInfo,
  com::sun::star::frame::XDispatchProvider,
  com::sun::star::frame::XDispatch
>
{
private:
  ::com::sun::star::uno::Reference< ::com::sun::star::uno::XComponentContext > mxContext;
  ::com::sun::star::uno::Reference< ::com::sun::star::frame::XFrame > mxFrame;
  ::com::sun::star::uno::Reference< ::com::sun::star::awt::XToolkit > mxToolkit;

protected:
    Reference< ::com::sun::star::sdbc::XDataSource > getDataSource( ) throw ( ::com::sun::star::uno::RuntimeException );
    Reference< ::com::sun::star::sdb::XOfficeDatabaseDocument > getDatabaseDocument( ) throw ( ::com::sun::star::uno::RuntimeException );
    void checkTable( Reference< ::com::sun::star::sdbc::XDataSource > &xDataSource);
    void checkColumnHeader( Reference < ::com::sun::star::sheet::XSpreadsheetDocument > &xDocument );
    void openForm();
    void openSheet();
    void createForm();
    //    void showMessageBox(const ::rtl::OUString & strMessage);
    void showMessageBox(const Reference< ::com::sun::star::awt::XToolkit >& rToolkit,
                        const Reference< ::com::sun::star::frame::XFrame >& rFrame,
                        const ::rtl::OUString& aTitle,
                        const ::rtl::OUString& aMsgText );
    void generarCronogramaTrabajo();
    void ingresarNotas();
    void registrarInscripcion();
    void reporteNotas();
    void reporteAnualNotas();
    void reporteInscriptos();
    void sheetNotas();
    void sheetNotas(const Reference< ::com::sun::star::sdbc::XConnection >&,
                    const Reference< ::com::sun::star::sheet::XSpreadsheet >&,
                    sal_Int32,
                    const ::rtl::OUString &,
                    sal_Int32,
                    const ::rtl::OUString &);

    sal_Int32 sheetNotasHeader ( const Reference< ::com::sun::star::sdbc::XConnection>& ,
                            const Reference< ::com::sun::star::sheet::XSpreadsheet >& ,
                            sal_Int32,
                            const ::rtl::OUString& ,
                                 sal_Int32,
                                 const ::rtl::OUString&);
    void sheetNotasRow( const Reference< ::com::sun::star::sdbc::XConnection>& ,
                        const Reference< ::com::sun::star::sheet::XSpreadsheet >& ,
                        sal_Int32 ,
                        const ::rtl::OUString& ,
                        sal_Int32 ,
                        sal_Int32 );

    void sheetNotasEstudiante ( const Reference< ::com::sun::star::sdbc::XConnection>& ,
                                const Reference< ::com::sun::star::sheet::XSpreadsheet >& ,
                                sal_Int32 ,
                                const ::rtl::OUString& ,
                                sal_Int32 ,
                                sal_Int32 ,
                                sal_Int32 ,
                                sal_Int32 );

    sal_Int32 findColumn ( const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                           const ::rtl::OUString& strColumn );


    Reference< ::com::sun::star::sheet::XSpreadsheet > getCurrentSheet();
    ::rtl::OUString  getAsignatura(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection, sal_Int32 nGestion);
    ::rtl::OUString  getPlantilla(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection);
    sal_Int32 generarEncabezado( const Reference< ::com::sun::star::sdbc::XDataSource >& xDataSource,
                           const Reference< ::com::sun::star::sheet::XSpreadsheet > & xSheet,
                           ::rtl::OUString &,
			   sal_Int32 nGestion  );
    void generarCalendario( const Reference< ::com::sun::star::sdbc::XDataSource >& xDataSource,
                            const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet,
                            const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                            const ::rtl::OUString &,
                            sal_Int32 nCalendarRow,
			    sal_Int32 nGestion );
    sal_Int32 generarMes( const Reference< ::com::sun::star::sdbc::XDataSource >& xDataSource,
                    const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet,
                    const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                    const ::rtl::OUString &,
                    sal_Int32 nCalendarRow,
                    sal_Int32 nMonth,
		    sal_Int32 nGestion );
    void generarHeaderRow ( const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                            sal_Int32 nRow );
    void generarRow( const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet,
                     const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                     const Reference< ::com::sun::star::sdbc::XRow >& xRow,
                     sal_Int32 nRow );
    Reference< ::com::sun::star::sheet::XSpreadsheetDocument > getDocumentSheet();
    void mergeMonth( const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet,
                                  const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                  sal_Int32 nMonth,
                                  sal_Int32 nStartRow,
                                  sal_Int32 nEndRow
                     );
    sal_Int32 getGestion(const Reference< ::com::sun::star::sdbc::XConnection>& xConnection);
    sal_Int32 getPeriodo(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection);


    void formularioNotas(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection,
                         sal_Int32 nGetsion,
                         sal_Int32 nPeriodo,
                         ::rtl::OUString&  strAsignatura);
    void formularioInscripcion(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection,
                         sal_Int32 nGetsion,
                         sal_Int32 nPeriodo,
                         ::rtl::OUString&  strAsignatura);
    void habilitarTrimestre();
    void ingresarTrimestre(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection,
			 sal_Int32 nGestion,
			 sal_Int32 nPeriodo,
			 ::rtl::OUString& strAsignatura);
    void UpdateCalendar();
    OUString PickCSVFile();

    Reference<com::sun::star::uno::XInterface>
    CreateInstance(const OUString& sService)
    {
        return mxContext->getServiceManager()->createInstanceWithContext(sService, mxContext);
    }

    short ShowMessageBox( ::com::sun::star::awt::MessageBoxType eType,
                          long                                  nButtons,
                          const OUString&                       sTitle,
                          const OUString&                       sMessage )
    {
        Reference<com::sun::star::awt::XWindowPeer> xWindowPeer( mxFrame->getComponentWindow(), UNO_QUERY_THROW );
        Reference<com::sun::star::awt::XToolkit2> xToolkit( mxToolkit, UNO_QUERY_THROW );
        Reference<com::sun::star::awt::XMessageBox> xMsgBox (
                xToolkit->createMessageBox(  xWindowPeer,
                                             eType,
                                             nButtons,
                                             sTitle,
                                             sMessage), UNO_QUERY_THROW);
        return xMsgBox->execute();
    }

public:
    EnsecProtocolHandler(const ::com::sun::star::uno::Reference< ::com::sun::star::uno::XComponentContext > &rxContext) :
    mxContext( rxContext ) {}
    ~EnsecProtocolHandler();


  // XInitialization
  virtual void SAL_CALL
    initialize( const ::com::sun::star::uno::Sequence< ::com::sun::star::uno::Any >& aArguments )
    throw (::com::sun::star::uno::Exception, ::com::sun::star::uno::RuntimeException);

  // XServiceInfo
  virtual ::rtl::OUString SAL_CALL
    getImplementationName(  )
    throw (::com::sun::star::uno::RuntimeException);

  virtual sal_Bool SAL_CALL
    supportsService( const ::rtl::OUString& ServiceName )
    throw (::com::sun::star::uno::RuntimeException);

  virtual ::com::sun::star::uno::Sequence< ::rtl::OUString > SAL_CALL
    getSupportedServiceNames(  )
    throw (::com::sun::star::uno::RuntimeException);

  // XDispatchProvider
  virtual ::com::sun::star::uno::Reference< ::com::sun::star::frame::XDispatch >  SAL_CALL
    queryDispatch(const ::com::sun::star::util::URL& aURL,
		  const ::rtl::OUString& sTargetFrameName,
		  sal_Int32 nSearchFlags )
    throw( ::com::sun::star::uno::RuntimeException );


  virtual ::com::sun::star::uno::Sequence < ::com::sun::star::uno::Reference< ::com::sun::star::frame::XDispatch > > SAL_CALL
    queryDispatches(const ::com::sun::star::uno::Sequence < ::com::sun::star::frame::DispatchDescriptor >& seqDescriptor )
    throw( ::com::sun::star::uno::RuntimeException );

  // XDipatch
  virtual void SAL_CALL
    dispatch( const ::com::sun::star::util::URL& aURL,
	      const Sequence < ::com::sun::star::beans::PropertyValue >& lArgs )
    throw (RuntimeException);

  virtual void SAL_CALL
    addStatusListener( const Reference< ::com::sun::star::frame::XStatusListener >& xControl,
		       const ::com::sun::star::util::URL& aURL )
    throw (RuntimeException);

  virtual void SAL_CALL
    removeStatusListener( const Reference< ::com::sun::star::frame::XStatusListener >& xControl,
			  const ::com::sun::star::util::URL& aURL )
    throw (RuntimeException);

};

#endif // _FLISOL_PROTOCOL_HANDLER_HXX
