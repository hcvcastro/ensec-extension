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

#include <stdio.h>
#include "osl/module.hxx"
#include "cppuhelper/bootstrap.hxx"
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include "ensec-component-export.h"
#include "ensec-protocol-handler.h"

#ifndef _COM_SUN_STAR_AWT_XMESSAGEBOX_HPP_
#include "com/sun/star/awt/XMessageBox.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_XMESSAGEBOXFACTORY_HPP_
#include "com/sun/star/awt/XMessageBoxFactory.hpp"
#endif

#ifndef _COM_SUN_STAR_FRAME_XDESKTOP_HPP_
#include "com/sun/star/frame/XDesktop.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_MESSAGEBOXBUTTONS_HPP_
#include "com/sun/star/awt/MessageBoxButtons.hpp"
#endif

#ifndef _COM_SUN_STAR_LANG_XMULTICOMPONENTFACTORY_HPP_
#include "com/sun/star/lang/XMultiComponentFactory.hpp"
#endif

#ifndef _COM_SUN_STAR_SHEET_XSPREADSHEET_HPP_
#include "com/sun/star/sheet/XSpreadsheet.hpp"
#endif

#ifndef _COM_SUN_STAR_CONTAINER_XNAMED_HPP_
#include "com/sun/star/container/XNamed.hpp"
#endif

#ifndef _COM_SUN_STAR_UTIL_DATETIME_HPP_
#include "com/sun/star/util/DateTime.hpp"
#endif

#ifndef _COM_SUN_STAR_LANG_XLOCALIZABLE_HPP_
#include "com/sun/star/lang/XLocalizable.hpp"
#endif

#ifndef _COM_SUN_STAR_UTIL_XNUMBERFORMATSSUPPLIER_HPP_
#include "com/sun/star/util/XNumberFormatsSupplier.hpp"
#endif

#ifndef _COM_SUN_STAR_UTIL_XNUMBERFORMATTYPES_HPP_
#include "com/sun/star/util/XNumberFormatTypes.hpp"
#endif

#ifndef _COM_SUN_STAR_UTIL_NUMBERFORMAT_HPP_
#include <com/sun/star/util/NumberFormat.hpp>
#endif

#ifndef _COM_SUN_STAR_BEANS_XPROPERTYSET_HPP_
#include "com/sun/star/beans/XPropertySet.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_XCONTROL_HPP_
#include "com/sun/star/awt/XControl.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_XDIALOG_HPP_
#include "com/sun/star/awt/XDialog.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_XBUTTON_HPP_
#include "com/sun/star/awt/XButton.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_TOOLKIT_HPP_
#include <com/sun/star/awt/Toolkit.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XCONTROLCONTAINER_HPP_
#include "com/sun/star/awt/XControlContainer.hpp"
#endif

#ifndef _COM_SUN_STAR_SHEET_XCELLRANGESQUERY_HPP_
#include "com/sun/star/sheet/XCellRangesQuery.hpp"
#endif

#ifndef _COM_SUN_STAR_SHEET_CELLFLAGS_HPP_
#include "com/sun/star/sheet/CellFlags.hpp"
#endif

#ifndef _COM_SUN_STAR_LANG_XSINGLESERVICEFACTORY_HPP_
#include "com/sun/star/lang/XSingleServiceFactory.hpp"
#endif

#ifndef _COM_SUN_STAR_UNO_XNAMINGSERVICE_HPP_
#include "com/sun/star/uno/XNamingService.hpp"
#endif

#ifndef _COM_SUN_STAR_FRAME_XSTORABLE_HPP_
#include "com/sun/star/frame/XStorable.hpp"
#endif

#ifndef _COM_SUN_STAR_UNO_RUNTIMEEXCEPTION_HPP_
#include "com/sun/star/uno/RuntimeException.hpp"
#endif

#ifndef _COM_SUN_STAR_CONTAINER_NOSUCHELEMENTEXCEPTION_HPP_
#include "com/sun/star/container/NoSuchElementException.hpp"
#endif

#ifndef _COM_SUN_STAR_SDB_XDOCUMENTDATASOURCE_HPP_
#include "com/sun/star/sdb/XDocumentDataSource.hpp"
#endif

#ifndef _COM_SUN_STAR_SDBC_XCONNECTION_HPP_
#include "com/sun/star/sdbc/XConnection.hpp"
#endif

#ifndef _COM_SUN_STAR_SDBCX_XTABLESSUPPLIER_HPP_
#include <com/sun/star/sdbcx/XTablesSupplier.hpp>
#endif

#ifndef _COM_SUN_STAR_FRAME_XCOMPONENTLOADER_HPP_
#include <com/sun/star/frame/XComponentLoader.hpp>
#endif

#ifndef _COM_SUN_STAR_BEANS_XFASTPROPERTYSET_HPP_
#include <com/sun/star/beans/XFastPropertySet.hpp>
#endif

#ifndef _COM_SUN_STAR_FORM_XFORM_HPP_
#include <com/sun/star/form/XForm.hpp>
#endif

#ifndef _COM_SUN_STAR_FORM_NAVIGATIONBARMODE_HPP_
#include <com/sun/star/form/NavigationBarMode.hpp>
#endif

#ifndef _COM_SUN_STAR_SDBCX_XCOLUMNSSUPPLIER_HPP_
#include <com/sun/star/sdbcx/XColumnsSupplier.hpp>
#endif

#ifndef _COM_SUN_STAR_SDB_XCOLUMN_HPP_
#include <com/sun/star/sdb/XColumn.hpp>
#endif

#ifndef _COM_SUN_STAR_TABLE_XTABLECOLUMNS_HPP_
#include <com/sun/star/table/XTableColumns.hpp>
#endif

#ifndef _COM_SUN_STAR_TABLE_XCOLUMNROWRANGE_HPP_
#include <com/sun/star/table/XColumnRowRange.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_FONTWEIGHT_HPP_
#include <com/sun/star/awt/FontWeight.hpp>
#endif

#ifndef _COM_SUN_STAR_TABLE_CELLHORIJUSTIFY_HPP_
#include <com/sun/star/table/CellHoriJustify.hpp>
#endif

#ifndef _COM_SUN_STAR_SDBC_XROW_HPP_
#include <com/sun/star/sdbc/XRow.hpp>
#endif

#ifndef _COM_SUN_STAR_SHEET_XSPREADSHEETVIEW_HPP_
#include "com/sun/star/sheet/XSpreadsheetView.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_XITEMLIST_HPP_
#include "com/sun/star/awt/XItemList.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_XLISTBOX_HPP_
#include "com/sun/star/awt/XListBox.hpp"
#endif

#ifndef _COM_SUN_STAR_UTIL_XMERGEABLE_HPP_
#include "com/sun/star/util/XMergeable.hpp"
#endif

#ifndef _COM_SUN_STAR_TABLE_BORDERLINE_HPP_
#include <com/sun/star/table/BorderLine.hpp>
#endif

#ifndef _COM_SUN_STAR_TABLE_TABLEBORDER_HPP_
#include <com/sun/star/table/TableBorder.hpp>
#endif

#ifndef _COM_SUN_STAR_TABLE_TABLEBORDERDISTANCES_HPP_
#include <com/sun/star/table/TableBorderDistances.hpp>
#endif

#ifndef _COM_SUN_STAR_TABLE_CELLVERTJUSTIFY_HPP_
#include <com/sun/star/table/CellVertJustify.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XTEXTLISTENER_HPP_
#include <com/sun/star/awt/XTextListener.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XTEXTCOMPONENT_HPP_
#include <com/sun/star/awt/XTextComponent.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XFIXEDTEXT_HPP_
#include <com/sun/star/awt/XFixedText.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XUNOCONTROLCONTAINER_HPP_
#include <com/sun/star/awt/XUnoControlContainer.hpp>
#endif

#ifndef _COM_SUN_STAR_SDBC_XRESULTSETUPDATE_HPP_
#include <com/sun/star/sdbc/XResultSetUpdate.hpp>
#endif

#ifndef _COM_SUN_STAR_SDBC_XROWUPDATE_HPP_
#include <com/sun/star/sdbc/XRowUpdate.hpp>
#endif

#ifndef _COM_SUN_STAR_SDB_XQUERYDEFINITIONSSUPPLIER_HPP_
#include <com/sun/star/sdb/XQueryDefinitionsSupplier.hpp>
#endif

#ifndef _COM_SUN_STAR_SDBC_XCOLUMNLOCATE_HPP_
#include <com/sun/star/sdbc/XColumnLocate.hpp>
#endif

#ifndef _COM_SUN_STAR_SDBC_XPARAMETERS_HPP_
#include <com/sun/star/sdbc/XParameters.hpp>
#endif

#ifndef _COM_SUN_STAR_XML_ATTRIBUTEDATA_HPP_
#include <com/sun/star/xml/AttributeData.hpp>
#endif

#ifndef _COM_SUN_STAR_TABLE_CELLCONTENTTYPE_HPP_
#include <com/sun/star/table/CellContentType.hpp>
#endif

#ifndef _COM_SUN_STAR_SHEET_XCELLADDRESSABLE_HPP_
#include <com/sun/star/sheet/XCellAddressable.hpp>
#endif

#ifndef _COM_SUN_STAR_UNO_XCOMPONENTCONTEXT_HPP_
#include <com/sun/star/uno/XComponentContext.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_WINDOWATTRIBUTE_HPP_
#include <com/sun/star/awt/WindowAttribute.hpp>
#endif

#ifndef _COM_SUN_STAR_DEPLOYMENT_PACKAGEINFORMATIONPROVIDER_HPP_
#include <com/sun/star/deployment/PackageInformationProvider.hpp>
#endif

#ifndef _COM_SUN_STAR_UTIL_URLTransformer_HPP_
#include <com/sun/star/util/URLTransformer.hpp>
#endif

#ifndef _COM_SUN_STAR_UTIL_XURLTransformer_HPP_
#include <com/sun/star/util/XURLTransformer.hpp>
#endif

#ifndef _COM_SUN_STAR_UCB_SIMPLEFILEACCESS_HPP_
#include <com/sun/star/ucb/SimpleFileAccess.hpp>
#endif


#include <com/sun/star/xml/sax/XExtendedDocumentHandler.hpp>

using namespace com::sun::star::uno;
using namespace com::sun::star::frame;
using com::sun::star::util::URL;



typedef ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XDocumentHandler > pFuncImportDialogModel (
    ::com::sun::star::uno::Reference<
    ::com::sun::star::container::XNameContainer > const & xDialogModel,
    ::com::sun::star::uno::Reference<
    ::com::sun::star::uno::XComponentContext > const & xContext,
    ::com::sun::star::uno::Reference<
        ::com::sun::star::frame::XModel > const & xDocument );



typedef ::cppu::WeakImplHelper1 < ::com::sun::star::awt::XActionListener > ClickHandlerBase;
typedef ::cppu::WeakImplHelper1 < ::com::sun::star::awt::XTextListener > EditHandlerBase;

class ClickHandler : public ClickHandlerBase
{
private:
  Reference < ::com::sun::star::awt::XDialog > mxDialog;
public:
  ClickHandler( Reference < ::com::sun::star::awt::XDialog > &aDialog ) : mxDialog(aDialog) {};
protected:
  // XActionListener
  virtual void SAL_CALL actionPerformed( const ::com::sun::star::awt::ActionEvent& rEvent ) throw (RuntimeException);
  // XEventListener
  virtual void SAL_CALL disposing( const ::com::sun::star::lang::EventObject& Source ) throw (RuntimeException);
};

void SAL_CALL ClickHandler::actionPerformed( const ::com::sun::star::awt::ActionEvent& rEvent ) throw (RuntimeException)
{
  if (rEvent.ActionCommand.compareToAscii("OK")==0) {
    mxDialog->endExecute();
  }
  else
    mxDialog->endExecute();
}

void SAL_CALL ClickHandler::disposing( const ::com::sun::star::lang::EventObject& /*Source*/ ) throw (RuntimeException)
{
  // not interested in
}


class UpdateHandler : public ClickHandlerBase
{
private:
    Reference < ::com::sun::star::sdbc::XConnection > mxConnection;
    Reference < ::com::sun::star::awt::XDialog > mxDialog;
    sal_Int32 mnGestion;
    sal_Int32 mnPeriodo;
    ::rtl::OUString mstrAsignatura;
public:
    UpdateHandler( const Reference< ::com::sun::star::sdbc::XConnection >& aConnection, 
                   const Reference < ::com::sun::star::awt::XDialog >& aDialog,
                   sal_Int32 aGestion,
                   sal_Int32 aPeriodo,
                   ::rtl::OUString astrAsignatura ) : mxConnection(aConnection), 
                                                      mxDialog(aDialog), 
                                                      mnGestion(aGestion), 
                                                      mnPeriodo(aPeriodo), 
                                                      mstrAsignatura(astrAsignatura)
    {}
protected:
    // XActionListener
    virtual void SAL_CALL actionPerformed( const ::com::sun::star::awt::ActionEvent& rEvent ) throw (RuntimeException);
    // XEventListener
    virtual void SAL_CALL disposing( const ::com::sun::star::lang::EventObject& Source ) throw (RuntimeException);
};

void SAL_CALL UpdateHandler::actionPerformed( const ::com::sun::star::awt::ActionEvent& rEvent ) throw (RuntimeException)
{
    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( mxDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XControl > xControl ( rEvent.Source , ::com::sun::star::uno::UNO_QUERY );
    Reference < ::com::sun::star::awt::XTextComponent > xText;
    Reference < ::com::sun::star::awt::XFixedText > xFixedText;
    Reference < ::com::sun::star::beans::XPropertySet > xPropFixedText;  
    Reference < ::com::sun::star::awt::XButton > xButton ( xControl , ::com::sun::star::uno::UNO_QUERY );
    Reference < ::com::sun::star::beans::XPropertySet > xPropButton ( xControl->getModel(), ::com::sun::star::uno::UNO_QUERY );
    ::rtl::OUString strLabel;

    xPropButton->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label"))) >>= strLabel;

    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT TITULO FROM EVALUACION WHERE GESTION=")) +
        ::rtl::OUString::number(mnGestion) +
        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
        mstrAsignatura +
        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO =")) +
        ::rtl::OUString::number(mnPeriodo);


    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( mxConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XStatement > xUpdate ( mxConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    ::rtl::OUString strTag;
    ::rtl::OUString strInsert;
    ::rtl::OUString strUpdate;
    ::rtl::OUString strEstudiante;
    ::rtl::OUString strEvaluacion;
    ::rtl::OString strDebug;
    

    if ( strLabel.compareToAscii("Insert") == 0 ) {
        while (xResultSet->next()){
            strInsert = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("INSERT INTO NOTA VALUES ( NULL,")) +
                ::rtl::OUString::number(mnGestion) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(", '")) +
                mstrAsignatura + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("', ")) +
                ::rtl::OUString::number(mnPeriodo) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(", ")) ;

            xControl.set (xControlContainer->getControl( xRow->getString(1) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_label"))), 
                          ::com::sun::star::uno::UNO_QUERY );
            xFixedText.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel() , ::com::sun::star::uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                strInsert += strTag + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(", ")) ;
            }

            xControl.set (xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))), 
                          ::com::sun::star::uno::UNO_QUERY );
            xFixedText.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel() , ::com::sun::star::uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                strInsert += strTag + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(", ")) ;
            }

            // 7 Fecha
            strInsert += ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NULL, ")) ;

            xControl.set (xControlContainer->getControl( xRow->getString(1)), 
                          ::com::sun::star::uno::UNO_QUERY );
            xText.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
            if ( xControl.is() && xText.is() ) {
                strTag = xText->getText();
                strInsert += strTag + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" )")) ;                
            }

            strDebug = ::rtl::OUStringToOString(strInsert, RTL_TEXTENCODING_UTF8);
            xUpdate->executeUpdate( strInsert );
        }
        xButton->setLabel(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Done")));
        xPropButton->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), ::com::sun::star::uno::makeAny(sal_False));
    }
    else if ( strLabel.compareToAscii("Update") == 0 ) {
        while (xResultSet->next()){

            strUpdate = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("UPDATE NOTA SET FECHA=NULL, NOTA="));

            xControl.set (xControlContainer->getControl(xRow->getString(1) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_label"))), 
                          ::com::sun::star::uno::UNO_QUERY );
            xFixedText.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel() , ::com::sun::star::uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strEvaluacion;
            }

            xControl.set (xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))), 
                          ::com::sun::star::uno::UNO_QUERY );
            xFixedText.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel() , ::com::sun::star::uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strEstudiante;
            }

            xControl.set (xControlContainer->getControl(xRow->getString(1)), 
                          ::com::sun::star::uno::UNO_QUERY );
            xText.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
            if ( xControl.is() && xText.is() ) {
                strUpdate += xText->getText() + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) ;                
            }

            strUpdate += ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("  WHERE GESTION=")) +
                ::rtl::OUString::number(mnGestion) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +                
                mstrAsignatura + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO=")) +
                ::rtl::OUString::number(mnPeriodo) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND EVALUACION=")) + 
                strEvaluacion + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ESTUDIANTE=")) + 
                strEstudiante;

            strDebug = ::rtl::OUStringToOString(strUpdate, RTL_TEXTENCODING_UTF8);
            xUpdate->executeUpdate( strUpdate );
            
        }
        xButton->setLabel(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Done")));
        xPropButton->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), ::com::sun::star::uno::makeAny(sal_False));
    }
}

void SAL_CALL UpdateHandler::disposing( const ::com::sun::star::lang::EventObject& /*Source*/ ) throw (RuntimeException)
{
  // not interested in
}



class EditHandler : public EditHandlerBase
{
private:
    Reference < ::com::sun::star::sdbc::XConnection > mxConnection;
    Reference < ::com::sun::star::awt::XDialog > mxDialog;
    sal_Int32 mnGestion;
    sal_Int32 mnPeriodo;
    ::rtl::OUString mstrAsignatura;

public:
    EditHandler( const Reference< ::com::sun::star::sdbc::XConnection >& aConnection, 
                 const Reference < ::com::sun::star::awt::XDialog >& aDialog,
                 sal_Int32 aGestion,
                 sal_Int32 aPeriodo,
                 ::rtl::OUString astrAsignatura ) : mxConnection(aConnection), mxDialog(aDialog), mnGestion(aGestion), mnPeriodo(aPeriodo), mstrAsignatura(astrAsignatura) {};
protected:
  // XTextListener           -> modify setzen
  virtual void SAL_CALL textChanged(const  ::com::sun::star::awt::TextEvent& rEvent) throw( ::com::sun::star::uno::RuntimeException );
  // XEventListener
  virtual void SAL_CALL disposing( const ::com::sun::star::lang::EventObject& Source ) throw (RuntimeException);

};
 

void SAL_CALL EditHandler::textChanged(const  ::com::sun::star::awt::TextEvent& rEvent) throw( ::com::sun::star::uno::RuntimeException )
{
    ::com::sun::star::util::Color clrGreenColor = 0x0000FF00;
    ::com::sun::star::util::Color clrRedColor = 0x00FF0000;
    //::com::sun::star::util::Color clrBlueColor = 0x000000FF;
    sal_Int32 nIDEstudiante = -1;
    bool bNotas = false;

    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( mxDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XTextComponent > xText ( rEvent.Source , ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XButton > xButton;
    Reference < ::com::sun::star::awt::XTextComponent > xEditText;
    Reference < ::com::sun::star::awt::XFixedText > xFixedText (xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XControl > xControl ( xFixedText, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::beans::XPropertySet > xPropTextModel ( xControl->getModel() , ::com::sun::star::uno::UNO_QUERY_THROW );  


    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.NOMBRE, ESTUDIANTE.ESTUDIANTE_ID FROM INSCRIPCION INNER JOIN ESTUDIANTE ON INSCRIPCION.GESTION=")) +
          ::rtl::OUString::number(mnGestion) +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND INSCRIPCION.ASIGNATURA='")) +
          mstrAsignatura +
           ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND INSCRIPCION.ESTUDIANTE = ESTUDIANTE.ESTUDIANTE_ID AND ESTUDIANTE.NOMBRE LIKE '%")) + 
          xText->getText().toAsciiUpperCase() +
           ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("%'"));

    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( mxConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    ::rtl::OUString strFound = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ESTUDIANTE NO ENCONTRADO!"));
    sal_Int32 nCounter = 0;

    while ( xResultSet->next() ) {
        if ( xResultSet->isFirst())
            strFound = xRow->getString(1);
        else
            strFound += ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) + xRow->getString(1);

        nIDEstudiante = xRow->getInt(2);
        nCounter++;
    }


    if ( nCounter == 1 ) {
        strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT TITULO FROM EVALUACION WHERE GESTION=")) +
            ::rtl::OUString::number(mnGestion) +
            ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
            mstrAsignatura +
            ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO =")) +
            ::rtl::OUString::number(mnPeriodo);

        //xStatement.set ( mxConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
        xResultSet.set ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
        xRow.set ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  


        while ( xResultSet->next() ) {
            xEditText.set (xControlContainer->getControl( xRow->getString(1) ), ::com::sun::star::uno::UNO_QUERY );
            if ( xEditText.is() )
                xEditText->setText( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("0")));
        }


        strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT EVALUACION.TITULO, NOTA.* FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION=EVALUACION.ID AND NOTA.GESTION=")) +
            ::rtl::OUString::number(mnGestion) +
            ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA='")) +
            mstrAsignatura +
            ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO =")) + 
            ::rtl::OUString::number(mnPeriodo) +
            ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ESTUDIANTE =")) + 
            ::rtl::OUString::number(nIDEstudiante);

        //xStatement.set ( mxConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
        xResultSet.set ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
        xRow.set ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  


        while ( xResultSet->next() ) {
            bNotas = true;
            xEditText.set (xControlContainer->getControl( xRow->getString(1) ), ::com::sun::star::uno::UNO_QUERY );
            if ( xEditText.is() )
                xEditText->setText( ::rtl::OUString::number(xRow->getInt(9)));
        }
    } else {
        strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT TITULO FROM EVALUACION WHERE GESTION=")) +
            ::rtl::OUString::number(mnGestion) +
            ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
            mstrAsignatura +
            ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO =")) +
            ::rtl::OUString::number(mnPeriodo);

        //xStatement.set ( mxConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
        xResultSet.set ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
        xRow.set ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  


        while ( xResultSet->next() ) {
            xEditText.set (xControlContainer->getControl( xRow->getString(1) ), ::com::sun::star::uno::UNO_QUERY );
            if ( xEditText.is() )
                xEditText->setText( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("0")));
        }
    }

    ::rtl::OUString strTag;
    Reference < ::com::sun::star::beans::XPropertySet > xPropControl;

    if (nCounter == 1) {
        xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TextColor")), ::com::sun::star::uno::makeAny(clrGreenColor) );
        Sequence< Reference< ::com::sun::star::awt::XControl > > controls = xControlContainer->getControls();

        for( sal_Int32 nIterator = 0; nIterator < controls.getLength(); nIterator++) {
            xPropControl.set (controls[nIterator]->getModel(), ::com::sun::star::uno::UNO_QUERY );
            if (xPropControl.is()) {
                xPropControl->getPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                if ( strTag.getLength() != 0 ) {
                    xPropControl->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), ::com::sun::star::uno::makeAny( sal_True ) );
                }
            }
        }
    }
    else  {
        xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TextColor")), ::com::sun::star::uno::makeAny(clrRedColor) );
        Sequence< Reference< ::com::sun::star::awt::XControl > > controls = xControlContainer->getControls();

        for( sal_Int32 nIterator = 0; nIterator < controls.getLength(); nIterator++) {
            xPropControl.set (controls[nIterator]->getModel(), ::com::sun::star::uno::UNO_QUERY );
            if (xPropControl.is()) {
                xPropControl->getPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                if ( strTag.getLength() != 0 )
                    xPropControl->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), ::com::sun::star::uno::makeAny( sal_False ) );
            }
        }
    }

    xButton.set (xControlContainer->getControl( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate"))), ::com::sun::star::uno::UNO_QUERY);
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), ::com::sun::star::uno::makeAny(::rtl::OUString::number(nIDEstudiante)) );
    xFixedText->setText(strFound);

    if ( bNotas )
        xButton->setLabel (::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Update")));
    else
        xButton->setLabel (::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Insert")));

}

void SAL_CALL EditHandler::disposing( const ::com::sun::star::lang::EventObject& /*Source*/ ) throw (RuntimeException)
{
  // not interested in
}


class NotaHandler : public EditHandlerBase
{
private:
    Reference < ::com::sun::star::awt::XDialog > mxDialog;

public:
    NotaHandler( const Reference < ::com::sun::star::awt::XDialog >& aDialog ) : mxDialog(aDialog) {};
protected:
    // XTextListener           -> modify setzen
    virtual void SAL_CALL textChanged(const  ::com::sun::star::awt::TextEvent& rEvent) throw( ::com::sun::star::uno::RuntimeException );
    // XEventListener
    virtual void SAL_CALL disposing( const ::com::sun::star::lang::EventObject& Source ) throw (RuntimeException);
    static void updateNotaFinal( const Reference < ::com::sun::star::awt::XDialog >& );

};

void SAL_CALL NotaHandler::textChanged(const  ::com::sun::star::awt::TextEvent& rEvent) throw( ::com::sun::star::uno::RuntimeException )
{
    float fPonderacion = 0.0;
    float fNota = 0.0;
    ::rtl::OUString strTag;
    ::rtl::OUString strNota;

    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( mxDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XTextComponent > xTextComponent ( rEvent.Source , ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XControl > xControlText ( xTextComponent , ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::beans::XPropertySet > xPropTextModel ( xControlText->getModel() , ::com::sun::star::uno::UNO_QUERY_THROW );  

    xPropTextModel->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
    Reference < ::com::sun::star::awt::XFixedText > xFixedText (xControlContainer->getControl(strTag), ::com::sun::star::uno::UNO_QUERY );
    
    if ( xFixedText.is() ) {
        Reference < ::com::sun::star::awt::XControl > xControlFixedText ( xFixedText, ::com::sun::star::uno::UNO_QUERY );
        Reference < ::com::sun::star::beans::XPropertySet > xPropFixedTextModel ( xControlFixedText->getModel() , ::com::sun::star::uno::UNO_QUERY_THROW );  

        xPropFixedTextModel->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
        fPonderacion = strTag.toFloat();
        strNota = xTextComponent->getText();
        fNota = strNota.toFloat();
        fNota = fNota * fPonderacion;
    }

    xFixedText->setText( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" = ")) + ::rtl::OUString::number( fNota ) );

    NotaHandler::updateNotaFinal( mxDialog );

}


void NotaHandler::updateNotaFinal(const Reference < ::com::sun::star::awt::XDialog >& aDialog )
{
    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( aDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Sequence< Reference< ::com::sun::star::awt::XControl > > controls = xControlContainer->getControls();
    Reference< ::com::sun::star::awt::XControl > xControl;
    Reference< ::com::sun::star::awt::XTextComponent > xEditControl;
    Reference < ::com::sun::star::beans::XPropertySet > xPropEditControl;  
    Reference < ::com::sun::star::awt::XFixedText > xFixedText;
    Reference < ::com::sun::star::beans::XPropertySet > xPropFixedControl;  
    float fPonderacion;
    float fNota;
    float fNotaFinal = 0.0;
    ::rtl::OUString strTag;

    for( sal_Int32 nIterator = 0; nIterator < controls.getLength(); nIterator++) {
        fNota = 0.0;
        fPonderacion = 0.0;
        xControl.set ( controls[nIterator], ::com::sun::star::uno::UNO_QUERY );
        xEditControl.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
        if (xEditControl.is()) {
            xPropEditControl.set ( xControl->getModel(), ::com::sun::star::uno::UNO_QUERY );
            fNota = xEditControl->getText().toFloat();            
            xPropEditControl->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
            xControl.set (xControlContainer->getControl (strTag), ::com::sun::star::uno::UNO_QUERY );
            xFixedText.set ( xControl, ::com::sun::star::uno::UNO_QUERY );
            if (xControl.is() && xFixedText.is()) {
                xPropFixedControl.set (xControl->getModel(), ::com::sun::star::uno::UNO_QUERY );
                xPropFixedControl->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                fPonderacion = strTag.toFloat();
                fNota = fNota * fPonderacion;
            }
            else fNota = 0.0;
        }
        fNotaFinal += fNota;
    }

    xFixedText.set( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("nota_final"))), ::com::sun::star::uno::UNO_QUERY );
    xFixedText->setText(::rtl::OUString::number( fNotaFinal));
}


void SAL_CALL NotaHandler::disposing( const ::com::sun::star::lang::EventObject& /*Source*/ ) throw (RuntimeException)
{
  // not interested in
}

EnsecProtocolHandler::~EnsecProtocolHandler( ) 
{
}



// XInitialization
// ---------------------------------------------------------------------------------
// Initializes the object.
//
void SAL_CALL 
EnsecProtocolHandler::initialize( const Sequence< Any >& aArguments ) 
  throw ( Exception, RuntimeException )
{
  Reference < XFrame > xFrame;
  if ( aArguments.getLength() ) {
    aArguments[0] >>= xFrame;
    mxFrame = xFrame;
  }

  // Create the toolkit to have access to it later
  mxToolkit = Reference< ::com::sun::star::awt::XToolkit >( ::com::sun::star::awt::Toolkit::create(mxContext), UNO_QUERY_THROW );
}


// XDispatchProvider
// -------------------------------------------------------------------------------------------
// Searches for an XDispatch for the specified URL within the specified target frame. 
//
Reference< XDispatch > SAL_CALL 
EnsecProtocolHandler::queryDispatch( const URL& aURL, 
				      const ::rtl::OUString& sTargetFrameName, 
				      sal_Int32 nSearchFlags )
  throw( RuntimeException )
{
  Reference < XDispatch > xRet;

  if ( sTargetFrameName.isEmpty() && nSearchFlags ) {
  }
      
  OSL_ENSURE(sTargetFrameName.isEmpty(), "sTargetFrameName is empty");
  OSL_ENSURE(nSearchFlags==0, "nSearchFlags no flags");

  if ( aURL.Protocol.compareToAscii("bolivia@shell.ensec:") == 0 ) {
    xRet = this;
  }
  return xRet;
}

// ----------------------------------------------------------------------------------------------
// Actually this method is redundant to XDispatchProvider::queryDispatch() to avoid multiple remote calls.
//
Sequence < Reference< XDispatch > > SAL_CALL 
EnsecProtocolHandler::queryDispatches( const Sequence < DispatchDescriptor >& seqDescripts )
  throw( RuntimeException )
{
  sal_Int32 nCount = seqDescripts.getLength();
  Sequence < Reference < XDispatch > > lDispatcher( nCount );
 
  for( sal_Int32 i=0; i<nCount; ++i )
    lDispatcher[i] = queryDispatch( seqDescripts[i].FeatureURL, seqDescripts[i].FrameName, seqDescripts[i].SearchFlags );
 
  return lDispatcher;
}


// XServiceInfo
// ---------------------------------------------------------------
// Provides the implementation name of the service implementation.
//
::rtl::OUString SAL_CALL 
EnsecProtocolHandler::getImplementationName(  )
  throw (RuntimeException)
{
  return EnsecProtocolHandler_getImplementationName();
}

// -------------------------------------------------------------------------
// Tests whether the specified service is supported, i.e. implemented by the implementation.
//
sal_Bool SAL_CALL 
EnsecProtocolHandler::supportsService( const ::rtl::OUString& rServiceName )
  throw (RuntimeException)
{
  return EnsecProtocolHandler_supportsService( rServiceName );
}

// -------------------------------------------------------------------------------------------------
// Provides the supported service names of the implementation, including also indirect service names.
//
Sequence< ::rtl::OUString > SAL_CALL 
EnsecProtocolHandler::getSupportedServiceNames(  )
  throw (RuntimeException)
{
  return EnsecProtocolHandler_getSupportedServiceNames();
}

// XDispatch
// ----------------------------------------------------------------------------------------------------
// dispatch to process commands
//
void SAL_CALL 
EnsecProtocolHandler::dispatch( const URL& aURL, 
				 const Sequence < ::com::sun::star::beans::PropertyValue >& lArgs ) 
  throw (RuntimeException)
{
    lArgs.getLength();

    /*const Reference< ::com::sun::star::deployment::XPackageInformationProvider > xPackageInfo = 
        ::com::sun::star::deployment::PackageInformationProvider::get(mxContext);
    Reference < ::com::sun::star::util::XURLTransformer > xTransformer ( ::com::sun::star::util::URLTransformer::create( mxContext ) );
    Reference < ::com::sun::star::ucb::XSimpleFileAccess3 > xSFI = ::com::sun::star::ucb::SimpleFileAccess::create(mxContext);

    ::rtl::OUString strRoot = xPackageInfo->getPackageLocation(::rtl::OUString("bolivia@shell:Ensec"));
    ::com::sun::star::util::URL oURL;

    oURL.Complete = strRoot + ::rtl::OUString("/dialogo.xdl");
    xTransformer->parseStrict(oURL);

    //showMessageBox (mxToolkit, mxFrame, ::rtl::OUString("Path"), oURL.Path);
    //showMessageBox (mxToolkit, mxFrame, ::rtl::OUString("Main"), oURL.Main);
    //showMessageBox (mxToolkit, mxFrame, ::rtl::OUString("Complete"), oURL.Complete);

    try {

        oslGenericFunction pSym = NULL;
        pFuncImportDialogModel * ptrFuncImportDialogModel;
        Reference< ::com::sun::star::io::XInputStream > xInput; // = xSFI->openFileRead( oURL.Complete );
        
        if (!xInput.is()) {
            ::rtl::OUString uri = "vnd.sun.star.expand:$LO_LIB_DIR/libxmlscriptlo.so";
            uri = cppu::bootstrap_expandUri(uri);
            ::rtl::OUString moduleUri(uri);
            oslModule lib = osl_loadModule(moduleUri.pData, SAL_LOADMODULE_LAZY | SAL_LOADMODULE_GLOBAL );

            if (! lib) {
                ::rtl::OUString const msg("loading component library failed: " + moduleUri);
                showMessageBox (mxToolkit, mxFrame, "error" , strRoot);                            
            }

            ::rtl::OUString aGetImportName = ::rtl::OUString("xmlscript::importDialogModel(com::sun::star::uno::Reference<com::sun::star::io::XInputStream> const&, com::sun::star::uno::Reference<com::sun::star::container::XNameContainer> const&, com::sun::star::uno::Reference<com::sun::star::uno::XComponentContext> const&, com::sun::star::uno::Reference<com::sun::star::frame::XModel> const&)");
            //::rtl::OUString aGetImportName = ::rtl::OUString("_ZN9xmlscript17importDialogModelERKN3com3sun4star3uno9ReferenceINS2_2io12XInputStreamEEERKNS4_INS2_9container14XNameContainerEEERKNS4_INS3_17XComponentContextEEERKNS4_INS2_5frame6XModelEEE");



            if ( lib )
                pSym = osl_getFunctionSymbol( lib, aGetImportName.pData );

            if (pSym != 0) {
                ptrFuncImportDialogModel = (pFuncImportDialogModel*) pSym;
                showMessageBox (mxToolkit, mxFrame, "Success" , aGetImportName);             }
            else
                showMessageBox (mxToolkit, mxFrame, "Failed" , aGetImportName);
        }
            
    }
    catch( Exception& ) {
        showMessageBox (mxToolkit, mxFrame, ::rtl::OUString("Excepcion"), strRoot);            
    }


    showMessageBox (mxToolkit, mxFrame, ::rtl::OUString("Terminado"), strRoot); */



    /*Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
      Reference < ::com::sun::star::uno::XInterface > xInter(xServiceManager->createInstanceWithContext(::rtl::OUString("com.sun.star.sheet.SpreadsheetDocument"),mxContext), ::com::sun::star::uno::UNO_QUERY_THROW);*/

  if ( aURL.Protocol.compareToAscii("bolivia@shell.ensec:") == 0 ) {
    if ( aURL.Path.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM("Calendario" ) ) ) {
        generarCronogramaTrabajo();
    }
    else if ( aURL.Path.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM("Reporte" ) ) ) {
      openSheet();
    }
    else if ( aURL.Path.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM("Notas" ) ) ) {
      ingresarNotas();
    }
    else if ( aURL.Path.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM("ReporteNotas" ) ) ) {
        reporteNotas();
    }
    else if ( aURL.Path.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM("ReporteAnualNotas" ) ) ) {
        reporteAnualNotas();
    }
    else if ( aURL.Path.equalsAsciiL( RTL_CONSTASCII_STRINGPARAM("SheetNotas" ) ) ) {
        sheetNotas();
    }
  }
}


Reference< ::com::sun::star::container::XNameContainer > lcl_createControlModel(
  const Reference< XComponentContext >& xContext) {
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xSMgr_( xContext->getServiceManager(), UNO_QUERY_THROW );
    Reference< ::com::sun::star::container::XNameContainer > xControlModel( xSMgr_->createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", xContext ), UNO_QUERY_THROW );
    return xControlModel;
}


Reference< ::com::sun::star::container::XNameContainer > lcl_createDialogModel(
        const Reference< XComponentContext >& xContext,
        const Reference< ::com::sun::star::io::XInputStream >& xInput,
        const Reference< ::com::sun::star::frame::XModel >& xModel
        ) throw ( Exception ) {
    
    Reference< ::com::sun::star::container::XNameContainer > xDialogModel( lcl_createControlModel(xContext) );

    //::xmlscript::importDialogModel( xInput, xDialogModel, xContext, xModel );
    xInput.is();
    xModel.is();
    xContext.is();
    return xDialogModel;
}



void 
EnsecProtocolHandler::sheetNotas()
{
    // get Data Source ensec
    Reference< ::com::sun::star::sdbc::XDataSource > xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }

    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    ::rtl::OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    // get periodo data
    sal_Int32 nPeriodo = getPeriodo(xConnection);
    if (nPeriodo == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("El periodo no ha sido seleccionado!")));
        return;
    }


    // get Plantilla
    ::rtl::OUString strPlantilla = getPlantilla(xConnection);
    if ( strPlantilla.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Plantilla no ha sido seleccionada!")));
        return;
    }


    // get Document Sheet 
    Reference< ::com::sun::star::sheet::XSpreadsheetDocument > xDocSheet = getDocumentSheet();
    if (!xDocSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Documento Sheet Actual!")));
        return;
    }

    // get Current Sheet 
    Reference< ::com::sun::star::sheet::XSpreadsheet > xSheet = getCurrentSheet();

    if (!xSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Sheet Actual!")));
        return;
    }

    sheetNotas(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, strPlantilla );

    showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                    ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Sheet Notas generado!")));

}

void
EnsecProtocolHandler::sheetNotas ( const Reference< ::com::sun::star::sdbc::XConnection>& xConnection,
                                   const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                   sal_Int32 nGestion,
                                   const ::rtl::OUString& strAsignatura,
                                   sal_Int32 nPeriodo,
                                   const ::rtl::OUString& strPlantilla)
{

    sal_Int32 nRow = sheetNotasHeader(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, strPlantilla);
    sheetNotasRow(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, nRow);
}



void
EnsecProtocolHandler::sheetNotasEstudiante ( const Reference< ::com::sun::star::sdbc::XConnection>& xConnection, 
                                             const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                             sal_Int32 nGestion,                                         
                                             const ::rtl::OUString& strAsignatura,                                         
                                             sal_Int32 nPeriodo,
                                             sal_Int32 nEstudiante,
                                             sal_Int32 nCol,
                                             sal_Int32 nRow)
{
    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT EVALUACION.TITULO, NOTA.NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.GESTION = EVALUACION.GESTION AND NOTA.ASIGNATURA = EVALUACION.ASIGNATURA AND NOTA.PERIODO = EVALUACION.PERIODO AND NOTA.EVALUACION = EVALUACION.ID  WHERE NOTA.GESTION = ? AND NOTA.ASIGNATURA = ? AND NOTA.PERIODO = ? AND NOTA.ESTUDIANTE = ?"));

    Reference< ::com::sun::star::sdbc::XPreparedStatement > xPreparedStatement ( xConnection->prepareStatement(strSQL), ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::sdbc::XParameters > xParameters ( xPreparedStatement, ::com::sun::star::uno::UNO_QUERY );

    xParameters->setInt(1, nGestion);
    xParameters->setString(2, strAsignatura);
    xParameters->setInt(3, nPeriodo);
    xParameters->setInt(4, nEstudiante);

    Reference< ::com::sun::star::sdbc::XResultSet > xResult = xPreparedStatement->executeQuery();
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResult, ::com::sun::star::uno::UNO_QUERY );  
    Reference< ::com::sun::star::sdbc::XColumnLocate > xColumn ( xResult, ::com::sun::star::uno::UNO_QUERY );

    Reference< ::com::sun::star::table::XCell > xCell;
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp; 

    while (xResult->next()) {
        nCol = findColumn(xSheet, xRow->getString(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TITULO")))));
        if ( nCol != -1 ) {
            xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
            xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
            xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                        ::com::sun::star::uno::makeAny(sal_Int32(2)) );
            xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                        ::com::sun::star::uno::makeAny(sal_Int32(2)) );
            xCell->setValue(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NOTA")))));
        }
    }

    /*nCol = findColumn(xSheet, ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formula_Practico")));
    if ( nCol != -1 ) {
        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=C")) + 
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" * 3")));
    }

    nCol = findColumn(xSheet, ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formula_DPS")));
    if ( nCol != -1 ) {
        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=SUM(E")) + 
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";F")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";G")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";H")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";I")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(") * (2/10)")));
    }

    nCol = findColumn(xSheet, ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formula_Parcial")));
    if ( nCol != -1 ) {
        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=SUM(K")) + 
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";L")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";M")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(") * (20/100)")));
    }

    nCol = findColumn(xSheet, ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formula_DPS_Parcial")));
    if ( nCol != -1 ) {
        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=SUM(J")) + 
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";N")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(")")));
    }

    nCol = findColumn(xSheet, ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formula_Trimestre")));
    if ( nCol != -1 ) {
        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=SUM(D")) + 
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";O")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(")")));
    }

    nCol = findColumn(xSheet, ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formula_Final")));
    if ( nCol != -1 ) {
        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=SUM(P")) + 
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(";Q")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(")")));
    }


    nCol = findColumn(xSheet, ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formula_Aprobado")));
    if ( nCol != -1 ) {
        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(1)) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                    ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=IF(R")) + 
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("=0;\"SIN NOTA\";IF(R")) +
                          ::rtl::OUString::valueOf(nRow + 1) +
                          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(">=51;\"APROBADO\";\"REPROBADO\"))")));
                          }*/
}

sal_Int32 EnsecProtocolHandler::findColumn ( const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                             const ::rtl::OUString& strColumn )
{
    sal_Int32 nCol = 0;
    sal_Int32 nRet = -1;
    Reference< ::com::sun::star::sheet::XCellAddressable > xCellAddress;
    Reference< ::com::sun::star::table::XCell > xCellFound;
    Reference< ::com::sun::star::table::XCell > xCell;
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp; 
    Reference< ::com::sun::star::container::XNameContainer > xAttributes;
    ::com::sun::star::xml::AttributeData rAttribute;
    ::com::sun::star::table::CellAddress rCellAddress;

    xCell.set (xSheet->getCellByPosition(nCol, 0), ::com::sun::star::uno::UNO_QUERY);
    xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);

    while (xCell->getType() != ::com::sun::star::table::CellContentType_EMPTY) {
        xCellProp->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("UserDefinedAttributes"))) >>= xAttributes;        
        if (xAttributes.is() && xAttributes->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")))) {
            xAttributes->getByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= rAttribute;
            if ( rAttribute.Value.compareTo(strColumn)==0) {
                xCellFound = xCell;
                break;
            }
        }
        
        xCell.set (xSheet->getCellByPosition(++nCol, 0), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
    }

    if ( xCellFound.is () ) {
        xCellAddress.set (xCellFound, ::com::sun::star::uno::UNO_QUERY);
        rCellAddress = xCellAddress->getCellAddress();
        nRet = rCellAddress.Column;
    }

    return nRet;
}


void
EnsecProtocolHandler::sheetNotasRow( const Reference< ::com::sun::star::sdbc::XConnection>& xConnection, 
                                     const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                     sal_Int32 nGestion,                                         
                                     const ::rtl::OUString& strAsignatura,                                         
                                     sal_Int32 nPeriodo,
                                     sal_Int32 nRow)
{
    sal_Int32 nCol = 0;
    sal_Int32 nIterator = 1;
    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.* FROM INSCRIPCION INNER JOIN ESTUDIANTE ON INSCRIPCION.ESTUDIANTE = ESTUDIANTE.ESTUDIANTE_ID WHERE INSCRIPCION.GESTION = ? AND INSCRIPCION.ASIGNATURA = ? ORDER BY ESTUDIANTE.NOMBRE"));

    Reference< ::com::sun::star::sdbc::XPreparedStatement > xPreparedStatement ( xConnection->prepareStatement(strSQL), ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::sdbc::XParameters > xParameters ( xPreparedStatement, ::com::sun::star::uno::UNO_QUERY );

    xParameters->setInt(1, nGestion);
    xParameters->setString(2, strAsignatura);

    Reference< ::com::sun::star::sdbc::XResultSet > xResult = xPreparedStatement->executeQuery();
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResult, ::com::sun::star::uno::UNO_QUERY );  
    Reference< ::com::sun::star::sdbc::XColumnLocate > xColumn ( xResult, ::com::sun::star::uno::UNO_QUERY );

    Reference< ::com::sun::star::table::XCell > xCell;
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp; 

    while (xResult->next()) {
        xCell.set (xSheet->getCellByPosition(nCol++, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                     ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                     ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(::rtl::OUString::number(nIterator++));

        xCell.set (xSheet->getCellByPosition(nCol++, nRow), ::com::sun::star::uno::UNO_QUERY);
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCell->setFormula(xRow->getString(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NOMBRE")))));

        sheetNotasEstudiante(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, 
                             xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ESTUDIANTE_ID")))),
                             nCol, nRow);
        nRow++;
        nCol = 0;
    }
}



sal_Int32
EnsecProtocolHandler::sheetNotasHeader ( const Reference< ::com::sun::star::sdbc::XConnection>& xConnection, 
                                         const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                         sal_Int32 nGestion,                                         
                                         const ::rtl::OUString& strAsignatura,                                         
                                         sal_Int32 nPeriodo,
                                         const ::rtl::OUString& strPlantilla)
{

    if (nGestion && strAsignatura.isEmpty() && nPeriodo )
    {
    }

    sal_Int32 nRow = 0;
    sal_Int32 nCol = 0;
    ::com::sun::star::util::Color clrBlackColor = 0x00000000;
    ::com::sun::star::table::BorderLine aBorderLine;
    ::com::sun::star::table::TableBorder aTableBorder;

    aBorderLine.Color = clrBlackColor;
    aBorderLine.OuterLineWidth = 1;
    aBorderLine.InnerLineWidth = 0;
    aBorderLine.LineDistance = 0;

    aTableBorder.IsTopLineValid = aTableBorder.IsBottomLineValid = aTableBorder.IsLeftLineValid = aTableBorder.IsRightLineValid = aTableBorder.IsDistanceValid = sal_True;
    aTableBorder.TopLine = aTableBorder.BottomLine = aTableBorder.LeftLine = aTableBorder.RightLine = aBorderLine;
    aTableBorder.Distance = 0;


    /*::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT EVALUACION.CATEGORIA, EVALUACION.TITULO, EVALUACION.ASIGNATURA FROM EVALUACION INNER JOIN GESTION ON EVALUACION.GESTION=GESTION.CODIGO INNER JOIN ASIGNATURA ON EVALUACION.ASIGNATURA=ASIGNATURA.CODIGO INNER JOIN PERIODO ON EVALUACION.PERIODO=PERIODO.PERIODO_ID WHERE EVALUACION.GESTION = ? AND EVALUACION.ASIGNATURA = ? AND EVALUACION.PERIODO = ? "));*/
    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM COLUMNA WHERE PLANTILLA ='")) + strPlantilla + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' ORDER BY CODIGO"));

    Reference< ::com::sun::star::sdbc::XPreparedStatement > xPreparedStatement ( xConnection->prepareStatement(strSQL), ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::sdbc::XParameters > xParameters ( xPreparedStatement, ::com::sun::star::uno::UNO_QUERY );

    //xParameters->setInt(1, nGestion);
    //xParameters->setString(2, strAsignatura);
    //xParameters->setInt(3, nPeriodo);

    Reference< ::com::sun::star::sdbc::XResultSet > xResult = xPreparedStatement->executeQuery();
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResult, ::com::sun::star::uno::UNO_QUERY );  
    Reference< ::com::sun::star::sdbc::XColumnLocate > xColumn ( xResult, ::com::sun::star::uno::UNO_QUERY );

    Reference< ::com::sun::star::table::XColumnRowRange > xColsRows( xSheet, ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::table::XTableColumns > xTableColumns ( xColsRows->getColumns(), ::com::sun::star::uno::UNO_QUERY );    
    Reference< ::com::sun::star::table::XTableRows > xTableRows ( xColsRows->getRows(), ::com::sun::star::uno::UNO_QUERY );    
    Reference< ::com::sun::star::beans::XPropertySet > xColumnProp;
    Reference< ::com::sun::star::beans::XPropertySet > xRowProp; 



    Reference< ::com::sun::star::table::XCell > xCell;
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp; 
    Reference< ::com::sun::star::container::XNameContainer > xAttributes;
    //::com::sun::star::xml::AttributeData rAttribute;


    xRowProp.set (xTableRows->getByIndex(nRow), ::com::sun::star::uno::UNO_QUERY );
    xRowProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny( sal_Int32(5000)));


    while (xResult->next()) {
        ::com::sun::star::xml::AttributeData rAttribute;
        xColumnProp.set (xTableColumns->getByIndex(nCol), ::com::sun::star::uno::UNO_QUERY );
        xColumnProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), 
                                      ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("WIDTH"))))) );

        xCell.set (xSheet->getCellByPosition(nCol, nRow), ::com::sun::star::uno::UNO_QUERY );
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->getPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("UserDefinedAttributes"))) >>= xAttributes;
        rAttribute.Type = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CDATA"));
        rAttribute.Value = xRow->getString(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TIPO"))));
        xAttributes->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), ::com::sun::star::uno::makeAny(rAttribute));
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("UserDefinedAttributes")), ::com::sun::star::uno::makeAny(xAttributes) );
        xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("RotateAngle")), 
                                    ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ANGULO"))))) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("IsTextWrapped")), 
                                     ::com::sun::star::uno::makeAny( sal_True ) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                     ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HORIZONTAL"))))) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                     ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VERTICAL"))))) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                     ::com::sun::star::uno::makeAny(aTableBorder) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ParaTopMargin")), 
                                     ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("MARGEN_TOP"))))));
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ParaBottomMargin")), 
                                     ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("MARGEN_BOTTOM"))))));
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharHeight")), 
                                     ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("FONT_HEIGHT"))))));
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     ::com::sun::star::uno::makeAny(xRow->getInt(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("FONT_WEIGHT"))))));



        xCell->setFormula(xRow->getString(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TITULO")))));

        xCell.set (xSheet->getCellByPosition(nCol, nRow + 1), ::com::sun::star::uno::UNO_QUERY );
        xCellProp.set (xCell, ::com::sun::star::uno::UNO_QUERY);
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                     ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                     ::com::sun::star::uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                     ::com::sun::star::uno::makeAny(aTableBorder) );
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharHeight")), 
                                     ::com::sun::star::uno::makeAny(sal_Int32(8)));
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD));

        xCell->setFormula(xRow->getString(xColumn->findColumn(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SUBTITULO")))));

        nCol++;
    }

    return nRow + 3;
}



void 
EnsecProtocolHandler::reporteNotas()
{
    // get Data Source ensec
    Reference< ::com::sun::star::sdbc::XDataSource > xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference< ::com::sun::star::sdb::XOfficeDatabaseDocument > xOfficeDbD  = getDatabaseDocument();

    if (!xOfficeDbD.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("El documento base de datos ensec no esta registrado!")));
        return;
    }

    Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }

    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    ::rtl::OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    // get periodo data
    sal_Int32 nPeriodo = getPeriodo(xConnection);
    if (nPeriodo == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("El periodo no ha sido seleccionado!")));
        return;
    }

    Reference < ::com::sun::star::sdb::XQueryDefinitionsSupplier > xQueryDS ( xDataSource, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::beans::XPropertySet > xPropQueryDefinition ( xQueryDS->getQueryDefinitions()->getByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA"))), ::com::sun::star::uno::UNO_QUERY_THROW );
        
    
    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.NOMBRE, CALC.NOTA FROM ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) AS NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          ::rtl::OUString::number(nGestion) + 
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +
          strAsignatura + 
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = ")) + 
          ::rtl::OUString::number(nPeriodo) +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) CALC INNER JOIN ESTUDIANTE ON CALC.ESTUDIANTE = ESTUDIANTE.ESTUDIANTE_ID ORDER BY CALC.NOTA DESC"));

    xPropQueryDefinition->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Command")), ::com::sun::star::uno::makeAny( strSQL ) );

    Reference< ::com::sun::star::sdb::XReportDocumentsSupplier > xReportDS ( xOfficeDbD, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::container::XNameAccess > xContainer ( xReportDS->getReportDocuments(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::frame::XComponentLoader > xLoader ( xReportDS->getReportDocuments(), ::com::sun::star::uno::UNO_QUERY_THROW );
    
    if ( xContainer->hasElements() && xContainer->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA"))) ) {
      Sequence < ::com::sun::star::beans::PropertyValue > aArgs(1);      
      aArgs[0].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ActiveConnection"));
      aArgs[0].Value <<= xConnection;
      /*aArgs[1].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CommandType"));
      aArgs[1].Value <<= ::com::sun::star::uno::makeAny(2);
      aArgs[2].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Command"));
      aArgs[2].Value <<= strSQL;*/


      xLoader->loadComponentFromURL( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA")), 
				     ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_blank")),
				     0,
				     aArgs );
    }
}

void 
EnsecProtocolHandler::reporteAnualNotas()
{
    // get Data Source ensec
    Reference< ::com::sun::star::sdbc::XDataSource > xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference< ::com::sun::star::sdb::XOfficeDatabaseDocument > xOfficeDbD  = getDatabaseDocument();

    if (!xOfficeDbD.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("El documento base de datos ensec no esta registrado!")));
        return;
    }

    Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }

    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    ::rtl::OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    Reference < ::com::sun::star::sdb::XQueryDefinitionsSupplier > xQueryDS ( xDataSource, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::beans::XPropertySet > xPropQueryDefinition ( xQueryDS->getQueryDefinitions()->getByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA_FINAL"))), ::com::sun::star::uno::UNO_QUERY_THROW );
        
    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.NOMBRE, COALESCE ( PRIMERO.NOTA, 0 ) PRIMERO, COALESCE ( SEGUNDO.NOTA, 0 ) SEGUNDO, COALESCE ( TERCERO.NOTA, 0 ) TERCERO, ( ( COALESCE ( PRIMERO.NOTA, 0 ) + COALESCE ( SEGUNDO.NOTA, 0 ) + COALESCE ( TERCERO.NOTA, 0 ) ) / 3 ) \"NOTA FINAL\" FROM ESTUDIANTE LEFT OUTER JOIN ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          ::rtl::OUString::number(nGestion) + 
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +
          strAsignatura + 
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = 1 GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) PRIMERO ON ESTUDIANTE.ESTUDIANTE_ID = PRIMERO.ESTUDIANTE LEFT OUTER JOIN ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          ::rtl::OUString::number(nGestion) +           
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +
          strAsignatura +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = 2 GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) SEGUNDO ON PRIMERO.ESTUDIANTE = SEGUNDO.ESTUDIANTE LEFT OUTER JOIN ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          ::rtl::OUString::number(nGestion) +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +          
          strAsignatura +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = 3 GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) TERCERO ON PRIMERO.ESTUDIANTE = TERCERO.ESTUDIANTE WHERE PRIMERO.GESTION = ")) + 
          ::rtl::OUString::number(nGestion);


    xPropQueryDefinition->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Command")), ::com::sun::star::uno::makeAny( strSQL ) );

    Reference< ::com::sun::star::sdb::XReportDocumentsSupplier > xReportDS ( xOfficeDbD, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::container::XNameAccess > xContainer ( xReportDS->getReportDocuments(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::frame::XComponentLoader > xLoader ( xReportDS->getReportDocuments(), ::com::sun::star::uno::UNO_QUERY_THROW );
    
    if ( xContainer->hasElements() && xContainer->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA"))) ) {
      Sequence < ::com::sun::star::beans::PropertyValue > aArgs(1);      
      aArgs[0].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ActiveConnection"));
      aArgs[0].Value <<= xConnection;
      /*aArgs[1].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CommandType"));
      aArgs[1].Value <<= ::com::sun::star::uno::makeAny(2);
      aArgs[2].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Command"));
      aArgs[2].Value <<= strSQL;*/

    
      xLoader->loadComponentFromURL( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA_FINAL")), 
				     ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_blank")),
				     0,
				     aArgs );
                     }
}



void
EnsecProtocolHandler::ingresarNotas()
{
    // get Data Source ensec
    Reference< ::com::sun::star::sdbc::XDataSource > xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }


    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    ::rtl::OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    // get periodo data
    sal_Int32 nPeriodo = getPeriodo(xConnection);
    if (nPeriodo == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("El periodo no ha sido seleccionado!")));
        return;
    }

    formularioNotas( xConnection, nGestion, nPeriodo, strAsignatura );

    //showMessageBox (::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Notas Ingresado! ")) + strAsignatura); // ::rtl::OUString::valueOf(nGestion));
}

void 
EnsecProtocolHandler::generarCronogramaTrabajo()
{
    // get Data Source ensec
    Reference< ::com::sun::star::sdbc::XDataSource > xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }


    // get Document Sheet 
    Reference< ::com::sun::star::sheet::XSpreadsheetDocument > xDocSheet = getDocumentSheet();
    if (!xDocSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Documento Sheet Actual!")));
        return;
    }

    // get Current Sheet 
    Reference< ::com::sun::star::sheet::XSpreadsheet > xSheet = getCurrentSheet();

    if (!xSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Sheet Actual!")));
        return;
    }

    // get Asignatura
    ::rtl::OUString strAsignatura = getAsignatura( xConnection, 2014 );
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    sal_Int32 nCalendarRow = generarEncabezado( xDataSource, xSheet, strAsignatura );

    generarCalendario( xDataSource, xDocSheet, xSheet, strAsignatura, nCalendarRow );


    showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                    ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Cronograma de Trabajo generado!")));
}


void
EnsecProtocolHandler::generarCalendario( const Reference< ::com::sun::star::sdbc::XDataSource >& xDataSource,
                                         const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet,
                                         const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                         const ::rtl::OUString & strAsignatura,
                                         sal_Int32 nCalendarRow ) 
{
    // 12 meses.
    for ( sal_Int32 nMonth = 1; nMonth <= 12; nMonth++ ) {
        nCalendarRow = generarMes( xDataSource, xDocSheet, xSheet, strAsignatura, nCalendarRow, nMonth );
    }
}

sal_Int32 
EnsecProtocolHandler::generarMes( const Reference< ::com::sun::star::sdbc::XDataSource >& xDataSource,
                                  const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet,                                  
                                  const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                  const ::rtl::OUString & strAsignatura,
                                  sal_Int32 nCalendarRow,
                                  sal_Int32 nMonth )
{
    char strBuffer[700];
    ::rtl::OString strParam = ::rtl::OUStringToOString(strAsignatura, RTL_TEXTENCODING_UTF8);
    sprintf(strBuffer, "SELECT STARTDATE, DESCRIPTION, CATEGORY FROM CALENDARIO WHERE MONTH(STARTDATE)=%d UNION SELECT STARTDATE, DESCRIPTION, CATEGORY FROM CRONOGRAMA WHERE MONTH(STARTDATE)=%d AND ASIGNATURA='%s'  ORDER BY STARTDATE, DESCRIPTION", nMonth, nMonth, strParam.getStr());

    Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );
    //::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM CALENDARIO"));
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(::rtl::OUString::createFromAscii(strBuffer)), 
                                                                 ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    sal_Int32 nStartRow = -1;
    sal_Int32 nEndRow = -1;

    while ( xResultSet->next() ) {
        if ( xResultSet->isFirst() ) {
            generarHeaderRow( xSheet, nCalendarRow++ );
            nStartRow = nCalendarRow;
        }

        generarRow ( xDocSheet, xSheet, xRow, nCalendarRow++ );
	}

    nEndRow = nCalendarRow - 1;

    mergeMonth(xDocSheet, xSheet, nMonth, nStartRow, nEndRow);

    return ++nCalendarRow;
}

void 
EnsecProtocolHandler::mergeMonth( const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet, 
                                  const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet, 
                                  sal_Int32 nMonth,
                                  sal_Int32 nStartRow,
                                  sal_Int32 nEndRow
                                  )
{
    xDocSheet.is();
    ::com::sun::star::util::Color clrBlackColor = 0x00000000;
    ::com::sun::star::table::BorderLine aBorderLine;
    ::com::sun::star::table::TableBorder aTableBorder;

    aBorderLine.Color = clrBlackColor;
    aBorderLine.OuterLineWidth = 1;
    aBorderLine.InnerLineWidth = 0;
    aBorderLine.LineDistance = 0;

    aTableBorder.IsTopLineValid = aTableBorder.IsBottomLineValid = aTableBorder.IsLeftLineValid = aTableBorder.IsRightLineValid = aTableBorder.IsDistanceValid = sal_True;
    aTableBorder.TopLine = aTableBorder.BottomLine = aTableBorder.LeftLine = aTableBorder.RightLine = aBorderLine;
    aTableBorder.Distance = 0;

    if ( nStartRow == -1 || nEndRow == -1 )
        return;

    Reference < ::com::sun::star::table::XCellRange > xRange ( xSheet->getCellRangeByPosition( 0, nStartRow, 0, nEndRow ), ::com::sun::star::uno::UNO_QUERY);
    Reference < ::com::sun::star::util::XMergeable > xMerge ( xRange, ::com::sun::star::uno::UNO_QUERY);
    Reference< ::com::sun::star::table::XCell > xCell (xRange->getCellByPosition(0,0), ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp ( xCell, ::com::sun::star::uno::UNO_QUERY );

    ::rtl::OUString strMes;

    switch( nMonth ) {
    case 1 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ENERO"));
        break;
    case 2 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("FEBRERO"));
        break;
    case 3 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("MARZO"));
        break;
    case 4 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ABRIL"));
        break;
    case 5 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("MAYO"));
        break;
    case 6 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("JUNIO"));
        break;
    case 7 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("JULIO"));
        break;
    case 8 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("AGOSTO"));
        break;
    case 9 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SEPTIEMBRE"));
        break;
    case 10 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OCTUBRE"));
        break;
    case 11 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NOVIEMBRE"));
        break;
    case 12 :
        strMes = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("DICIEMBRE"));
        break;
    }

    xCell->setFormula( strMes );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                 ::com::sun::star::uno::makeAny(sal_Int16(2)) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 ::com::sun::star::uno::makeAny(aTableBorder) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("RotateAngle")), 
                                 ::com::sun::star::uno::makeAny(sal_Int16(9000)) );

    xMerge->merge( sal_True );


}


void
EnsecProtocolHandler::generarHeaderRow ( const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet,
                                         sal_Int32 nRow )
{
    ::com::sun::star::util::Color clrLightGreyColor = 0x00C0C0C0;
    ::com::sun::star::util::Color clrBlackColor = 0x00000000;
    ::com::sun::star::table::BorderLine aBorderLine;
    ::com::sun::star::table::TableBorder aTableBorder;

    aBorderLine.Color = clrBlackColor;
    aBorderLine.OuterLineWidth = 1;
    aBorderLine.InnerLineWidth = 0;
    aBorderLine.LineDistance = 0;

    aTableBorder.IsTopLineValid = aTableBorder.IsBottomLineValid = aTableBorder.IsLeftLineValid = aTableBorder.IsRightLineValid = aTableBorder.IsDistanceValid = sal_True;
    aTableBorder.TopLine = aTableBorder.BottomLine = aTableBorder.LeftLine = aTableBorder.RightLine = aBorderLine;
    aTableBorder.Distance = 0;
    
    Reference< ::com::sun::star::table::XCell > xCell (xSheet->getCellByPosition(0, nRow), ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("MES")) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CellBackColor")), 
                                 ::com::sun::star::uno::makeAny(clrLightGreyColor) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 ::com::sun::star::uno::makeAny(aTableBorder) );

    xCell.set (xSheet->getCellByPosition(1, nRow), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("FECHA")) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CellBackColor")), 
                                 ::com::sun::star::uno::makeAny(clrLightGreyColor) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 ::com::sun::star::uno::makeAny(aTableBorder) );


    xCell.set (xSheet->getCellByPosition(2, nRow), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ACTIVIDAD")) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CellBackColor")), 
                                 ::com::sun::star::uno::makeAny(clrLightGreyColor) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 ::com::sun::star::uno::makeAny(aTableBorder) );
}

void
EnsecProtocolHandler::generarRow( const Reference< ::com::sun::star::sheet::XSpreadsheetDocument >& xDocSheet, 
                                  const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet, 
                                  const Reference< ::com::sun::star::sdbc::XRow >& xRow,
                                  sal_Int32 nRow )
{
    ::com::sun::star::util::Color clrBlackColor = 0x00000000;
    ::com::sun::star::table::BorderLine aBorderLine;
    ::com::sun::star::table::TableBorder aTableBorder;
    ::com::sun::star::table::TableBorderDistances aTableBorderDistances;
    ::com::sun::star::util::Date aDate;
    ::com::sun::star::lang::Locale olocale;
    sal_Int32 nNewIndex = 0;

    aBorderLine.Color = clrBlackColor;
    aBorderLine.OuterLineWidth = 1;
    aBorderLine.InnerLineWidth = 0;
    aBorderLine.LineDistance = 0;

    aTableBorder.IsTopLineValid = aTableBorder.IsBottomLineValid = aTableBorder.IsLeftLineValid = aTableBorder.IsRightLineValid = aTableBorder.IsDistanceValid = sal_True;
    aTableBorder.TopLine = aTableBorder.BottomLine = aTableBorder.LeftLine = aTableBorder.RightLine = aBorderLine;
    aTableBorder.Distance = 0;

    aTableBorderDistances.LeftDistance = 10;
    aTableBorderDistances.IsLeftDistanceValid = sal_True;

    ::rtl::OUString strCategory = xRow->getString(3);

    Reference < ::com::sun::star::util::XNumberFormatsSupplier > xFormatsSupplier ( xDocSheet, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::util::XNumberFormats > xNumberFormats ( xFormatsSupplier->getNumberFormats(), ::com::sun::star::uno::UNO_QUERY_THROW); 
    //Reference < ::com::sun::star::util::XNumberFormatTypes > xFormatTypes ( xFormatsSupplier->getNumberFormats(), ::com::sun::star::uno::UNO_QUERY_THROW); 

    Reference< ::com::sun::star::table::XCell > xCell (xSheet->getCellByPosition(1, nRow), ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp ( xCell, ::com::sun::star::uno::UNO_QUERY );

    //int nKeyFormat = xFormatTypes->getStandardFormat(::com::sun::star::util::NumberFormat::DATE, olocale);
    
    //Reference< ::com::sun::star::beans::XPropertySet > xFormatProp =  xNumberFormats->getByKey( nKeyFormat );
    //xFormatProp->getPropertyValue( ::rtl::OUString( RTL_CONSTASCII_USTRINGPARAM( "Locale" ) ) ) >>= olocale;

    nNewIndex = xNumberFormats->queryKey( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NN D")), olocale, false );
    if ( nNewIndex == -1 ) // format not defined
        nNewIndex = xNumberFormats->addNew( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NN D")), olocale );


    char strBuffer[20];
    aDate = xRow->getDate(1);
    sprintf(strBuffer, "%d/%d/%d", aDate.Year, aDate.Month, aDate.Day);

    xCell->setFormula( ::rtl::OUString::createFromAscii(strBuffer) );
    if ( strCategory.getLength() != 0 && strCategory.compareToAscii("Normal") != 0 ) 
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                 ::com::sun::star::uno::makeAny(sal_Int16(2)) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_LEFT) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ParaIndent")), 
                                 ::com::sun::star::uno::makeAny(sal_Int16(400)) );
    xCellProp->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NumberFormat")), ::com::sun::star::uno::makeAny(nNewIndex));
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 ::com::sun::star::uno::makeAny(aTableBorder) );
    /*xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorderDistances")), 
      ::com::sun::star::uno::makeAny(aTableBorderDistances) );*/


    xCell.set (xSheet->getCellByPosition(2, nRow), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( xRow->getString(2) );
    if ( strCategory.getLength() != 0 && strCategory.compareToAscii("Normal") != 0 ) 
        xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("IsTextWrapped")), 
                                 ::com::sun::star::uno::makeAny( sal_True ) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_LEFT) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ParaIndent")), 
                                 ::com::sun::star::uno::makeAny(sal_Int16(300)) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                 ::com::sun::star::uno::makeAny(sal_Int16(2)) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 ::com::sun::star::uno::makeAny(aTableBorder) );


}


sal_Int32
EnsecProtocolHandler::generarEncabezado( const Reference< ::com::sun::star::sdbc::XDataSource >& xDataSource, 
                                         const Reference< ::com::sun::star::sheet::XSpreadsheet >& xSheet, 
                                         ::rtl::OUString& strAsignatura )
{

    Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ASIGNATURA.CODIGO, CARRERA.CARRERA, ASIGNATURA.NOMBRE, DOCENTE.NOMBRE, ASIGNATURA.CURSO, ASIGNATURA.GESTION  FROM ASIGNATURA INNER JOIN CARRERA ON ASIGNATURA.NOMBRE='")) + 
        strAsignatura + 
        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("' AND ASIGNATURA.GESTION=2014 AND CARRERA.CODIGO = ASIGNATURA.CARRERA INNER JOIN DOCENTE ON ASIGNATURA.DOCENTE = DOCENTE.CODIGO"));

    ::rtl::OString strDebug = ::rtl::OUStringToOString(strSQL, RTL_TEXTENCODING_UTF8);
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xAsignatura ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    if ( !xResultSet->next() )
        return 8;

    strAsignatura = xAsignatura->getString(1);

    Reference< ::com::sun::star::table::XColumnRowRange > xColsRows( xSheet, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::table::XTableColumns > xTableColumns ( xColsRows->getColumns(), ::com::sun::star::uno::UNO_QUERY_THROW );    
    Reference< ::com::sun::star::beans::XPropertySet > xColumnProp (xTableColumns->getByIndex( 0 ), ::com::sun::star::uno::UNO_QUERY_THROW );
	xColumnProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(2470)) );

    xColumnProp.set (xTableColumns->getByIndex( 1 ), ::com::sun::star::uno::UNO_QUERY_THROW );
	xColumnProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(2470)) );
    
    xColumnProp.set (xTableColumns->getByIndex( 2 ), ::com::sun::star::uno::UNO_QUERY_THROW );
	xColumnProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(13840)) );

    Reference < ::com::sun::star::table::XCellRange > xHeader ( xSheet->getCellRangeByPosition( 0, 0, 2, 0 ), ::com::sun::star::uno::UNO_QUERY);
    Reference < ::com::sun::star::util::XMergeable > xMerge ( xHeader, ::com::sun::star::uno::UNO_QUERY);
    Reference< ::com::sun::star::table::XCell > xCell (xHeader->getCellByPosition(0,0), ::com::sun::star::uno::UNO_QUERY );
    Reference< ::com::sun::star::beans::XPropertySet > xCellProp ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(3) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    xMerge->merge( sal_True );

    xHeader.set ( xSheet->getCellRangeByPosition( 0, 1, 2, 1 ), ::com::sun::star::uno::UNO_QUERY );
    xMerge.set ( xHeader, ::com::sun::star::uno::UNO_QUERY );
    xCell.set ( xHeader->getCellByPosition(0,0), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CRONOGRAMA DE TRABAJO")) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    xMerge->merge( sal_True );


    xHeader.set ( xSheet->getCellRangeByPosition( 0, 2, 2, 2 ), ::com::sun::star::uno::UNO_QUERY );
    xMerge.set ( xHeader, ::com::sun::star::uno::UNO_QUERY );
    xCell.set ( xHeader->getCellByPosition(0,0), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("GESTION ")) + xAsignatura->getString(6) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    xMerge->merge( sal_True );


    xCell.set ( xSheet->getCellByPosition( 0, 4 ), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CARRERA :")) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );


    xCell.set ( xSheet->getCellByPosition( 1, 4 ), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(2) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );


    xCell.set ( xSheet->getCellByPosition( 0, 5 ), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("DOCENTE :")) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );

    xCell.set ( xSheet->getCellByPosition( 1, 5 ), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(4) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );


    xCell.set ( xSheet->getCellByPosition( 0, 6 ), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CURSO :")) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );

    xCell.set ( xSheet->getCellByPosition( 1, 6 ), ::com::sun::star::uno::UNO_QUERY );
    xCellProp.set ( xCell, ::com::sun::star::uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(5) );
    xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );

    return 8;
}

void 
EnsecProtocolHandler::formularioNotas(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection, 
                                      sal_Int32 nGestion,
                                      sal_Int32 nPeriodo,
                                      ::rtl::OUString&  strAsignatura)
{
    ::rtl::OUString strTitle;
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::uno::XInterface > xDialogModel (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::beans::XPropertySet > xPropDlg ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::container::XNameContainer > xNameContainer ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );

    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(250)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(250)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Formulario Notas"))) );

    Reference< ::com::sun::star::lang::XMultiServiceFactory > xSMDialog ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW);

    // Label 
    Reference< ::com::sun::star::uno::XInterface > xTextModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropTextModel ( xTextModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

    strTitle = ::rtl::OUString::number(nGestion) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) +
        strAsignatura + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) +
        ::rtl::OUString::number(nPeriodo);

    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(10)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(70)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(strTitle) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("txtAsignatura"))) );

    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("txtAsignatura")), ::com::sun::star::uno::makeAny(xTextModel));

    // Edit Ok
    Reference< ::com::sun::star::uno::XInterface > xEditModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlEditModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropEditModel ( xEditModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

    xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(25)) );
    xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(70)) );
    xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("editEstudiante"))) );

    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("editEstudiante")), ::com::sun::star::uno::makeAny(xEditModel));

    // Label 
    xTextModel.set (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    xPropTextModel.set ( xTextModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(45)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(150)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(strAsignatura) );
    xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))) );

    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante")), ::com::sun::star::uno::makeAny(xTextModel));


    // GroupBox OK
    Reference< ::com::sun::star::uno::XInterface > xGroupBoxModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlGroupBoxModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropGroupBoxModel (xGroupBoxModel, ::com::sun::star::uno::UNO_QUERY_THROW);  
								     
    xPropGroupBoxModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropGroupBoxModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(55)) );
    xPropGroupBoxModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(200)) );
    xPropGroupBoxModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
    xPropGroupBoxModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), 
                                           ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("containerNotas"))) );
    xPropGroupBoxModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Notas")))  );

    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("grpNotas")), ::com::sun::star::uno::makeAny(xGroupBoxModel));


    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM EVALUACION WHERE GESTION=")) + 
          ::rtl::OUString::number(nGestion) +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND PERIODO=")) +
          ::rtl::OUString::number(nPeriodo) +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
          strAsignatura +
          ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("'"));
    
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    sal_Int32 nYStart = 65;
    sal_Int32 nYPos = nYStart;
    sal_Int32 nXPos = 37;
    sal_Int32 nHeight = 10;
    sal_Int32 nWidth = 100;
    sal_Int32 nYOffset = 10;
    sal_Int32 nXOffset = 5;
    sal_Int32 nEditWidth = 20;
    sal_Int32 nCalcWidth = 50;
    
    ::rtl::OUString strLabel;
    
    while ( xResultSet->next() ) {
        xTextModel.set (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
        xPropTextModel.set (xTextModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

        xEditModel.set (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlEditModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
        xPropEditModel.set ( xEditModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

        strLabel = xRow->getString(4) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) + ::rtl::OUString::number(xRow->getFloat(6)) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("-")) + ::rtl::OUString::number(xRow->getFloat(7));            

        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(nXPos)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(nYPos)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(nWidth)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(nHeight)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(strLabel));
        xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), ::com::sun::star::uno::makeAny(::rtl::OUString::number(xRow->getInt(1))));

        xNameContainer->insertByName(xRow->getString(4) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_label")), ::com::sun::star::uno::makeAny(xTextModel));

        xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny( nXPos + nWidth + nXOffset ) );
        xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny( nYPos ) );
        xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny( nEditWidth ) );
        xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny( nHeight ) );
        xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), ::com::sun::star::uno::makeAny( sal_False ) );
        xPropEditModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), ::com::sun::star::uno::makeAny(xRow->getString(4) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_nota")) ));

        xNameContainer->insertByName(xRow->getString(4), ::com::sun::star::uno::makeAny(xEditModel));

        xTextModel.set (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
        xPropTextModel.set (xTextModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16( nXPos + nWidth + nEditWidth + nXOffset)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(nYPos)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(nCalcWidth)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(nHeight)));
        xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(" = "))));
        xPropTextModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), ::com::sun::star::uno::makeAny(::rtl::OUString::number( xRow->getFloat(7) / xRow->getFloat(6) )));


        xNameContainer->insertByName(xRow->getString(4) + ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_nota")) , ::com::sun::star::uno::makeAny(xTextModel));

        nYPos += nYOffset;

	}

    xTextModel.set (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), 
                    ::com::sun::star::uno::UNO_QUERY_THROW);
    xPropTextModel.set (xTextModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(nXPos)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(nYPos)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(nWidth)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(nHeight)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), 
                                     ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Nota Final"))));

    xNameContainer->insertByName( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("nota_final_texto")),
                                  ::com::sun::star::uno::makeAny(xTextModel));

    xTextModel.set (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), 
                    ::com::sun::star::uno::UNO_QUERY_THROW);
    xPropTextModel.set (xTextModel , ::com::sun::star::uno::UNO_QUERY_THROW );  

    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(nXPos + nWidth + nXOffset)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(nYPos)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(nEditWidth)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(nHeight)));
    xPropTextModel->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), 
                                     ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("0"))));

    xNameContainer->insertByName( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("nota_final")),
                                  ::com::sun::star::uno::makeAny(xTextModel));

    nYPos += nYOffset;

    xPropGroupBoxModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(nYPos + nHeight - nYStart) );    
    
    nYPos += nYOffset;

    // Button Update OK
    Reference< ::com::sun::star::uno::XInterface > xButtonModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropButtonModel ( xButtonModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(nYPos)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate"))) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("None"))) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Update"))));
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), ::com::sun::star::uno::makeAny( sal_False ) );

    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate")), ::com::sun::star::uno::makeAny(xButtonModel));


    // Button OK
    xButtonModel.set (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    xPropButtonModel.set ( xButtonModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35 + 50 + 5)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(nYPos)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), ::com::sun::star::uno::makeAny(xButtonModel));


    Reference < ::com::sun::star::uno::XInterface > xDialog (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XControl > xControl ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XControlModel > xControlModel ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference < ::com::sun::star::awt::XToolkit > xToolkit (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XWindow > xWindow ( xControl, ::com::sun::star::uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XButton > xButton ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XButton > xUpdate ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate"))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XTextComponent > xText ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("editEstudiante"))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XTextComponent > xTextNota;

    Reference < ::com::sun::star::awt::XDialog > xDlg ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::awt::XActionListener > xEnsureDelete( new ClickHandler( xDlg ) );
    Reference< ::com::sun::star::awt::XTextListener > xNotaHandler( new NotaHandler( xDlg ) );
    Reference< ::com::sun::star::awt::XTextListener > xTextHandler( new EditHandler( xConnection, xDlg, nGestion, nPeriodo, strAsignatura ) );
    Reference< ::com::sun::star::awt::XActionListener > xUpdateHandler( new UpdateHandler( xConnection, xDlg, nGestion, nPeriodo, strAsignatura ) );

    xButton->addActionListener( xEnsureDelete );
    xUpdate->addActionListener( xUpdateHandler );
    xText->addTextListener( xTextHandler );

    xResultSet.set ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    xRow.set ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    while ( xResultSet->next() ) {
        xTextNota.set ( xControlContainer->getControl(xRow->getString(4)), ::com::sun::star::uno::UNO_QUERY );        
        if ( xTextNota.is() )
            xTextNota->addTextListener( xNotaHandler );
    }

    xDlg->execute();

    Reference < ::com::sun::star::lang::XComponent > xDispose ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    xDispose->dispose();


}



sal_Int32  
EnsecProtocolHandler::getGestion(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection)
{
    sal_Int32 nResult = 0;
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::uno::XInterface > xDialogModel (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::beans::XPropertySet > xPropDlg ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );

    //xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
    //xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Gestion"))) );

    Reference< ::com::sun::star::lang::XMultiServiceFactory > xSMDialog ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::uno::XInterface > xComboModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropComboModel ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    Reference < ::com::sun::star::awt::XItemList > xItemList ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), ::com::sun::star::uno::makeAny(sal_True) );
    //xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ReadOnly")), ::com::sun::star::uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), ::com::sun::star::uno::makeAny(sal_Int8(10)) );


    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM GESTION"));
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));			
        xItemList->setItemData( nIndex, ::com::sun::star::uno::makeAny(xRow->getInt(1)));
        nIndex = xItemList->getItemCount();
	}

    Reference< ::com::sun::star::uno::XInterface > xButtonModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropButtonModel ( xButtonModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference < ::com::sun::star::container::XNameContainer > xNameContainer ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), ::com::sun::star::uno::makeAny(xComboModel));
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), ::com::sun::star::uno::makeAny(xButtonModel));

    Reference < ::com::sun::star::uno::XInterface > xDialog (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XControl > xControl ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XControlModel > xControlModel ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference < ::com::sun::star::awt::XToolkit > xToolkit (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XWindow > xWindow ( xControl, ::com::sun::star::uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XButton > xButton ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XListBox > xListBox ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), ::com::sun::star::uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference < ::com::sun::star::awt::XDialog > xDlg ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::awt::XActionListener > xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //   sal_Int32 clipformat;
    //    aSysDataType >>= clipformat;
 
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= nResult;

    Reference < ::com::sun::star::lang::XComponent > xDispose ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return nResult;

}


sal_Int32  
EnsecProtocolHandler::getPeriodo(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection)
{
    sal_Int32 nResult = 0;
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::uno::XInterface > xDialogModel (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::beans::XPropertySet > xPropDlg ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );

    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Periodo"))) );

    Reference< ::com::sun::star::lang::XMultiServiceFactory > xSMDialog ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::uno::XInterface > xComboModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropComboModel ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    Reference < ::com::sun::star::awt::XItemList > xItemList ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), ::com::sun::star::uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), ::com::sun::star::uno::makeAny(sal_Int8(10)) );


    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM PERIODO"));
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));			
        xItemList->setItemData( nIndex, ::com::sun::star::uno::makeAny(xRow->getInt(1)));
        nIndex = xItemList->getItemCount();
	}

    Reference< ::com::sun::star::uno::XInterface > xButtonModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropButtonModel ( xButtonModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference < ::com::sun::star::container::XNameContainer > xNameContainer ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), ::com::sun::star::uno::makeAny(xComboModel));
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), ::com::sun::star::uno::makeAny(xButtonModel));

    Reference < ::com::sun::star::uno::XInterface > xDialog (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XControl > xControl ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XControlModel > xControlModel ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference < ::com::sun::star::awt::XToolkit > xToolkit (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XWindow > xWindow ( xControl, ::com::sun::star::uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XButton > xButton ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XListBox > xListBox ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), ::com::sun::star::uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference < ::com::sun::star::awt::XDialog > xDlg ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::awt::XActionListener > xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //   sal_Int32 clipformat;
    //    aSysDataType >>= clipformat;
 
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= nResult;

    Reference < ::com::sun::star::lang::XComponent > xDispose ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return nResult;

}


::rtl::OUString 
EnsecProtocolHandler::getAsignatura(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection, sal_Int32 nGestion)
{
    ::rtl::OUString strResult;
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::uno::XInterface > xDialogModel (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::beans::XPropertySet > xPropDlg ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );

    //xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
    //xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Asignatura"))) );

    Reference< ::com::sun::star::lang::XMultiServiceFactory > xSMDialog ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::uno::XInterface > xComboModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropComboModel ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    Reference < ::com::sun::star::awt::XItemList > xItemList ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), ::com::sun::star::uno::makeAny(sal_True) );
    //xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ReadOnly")), ::com::sun::star::uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), ::com::sun::star::uno::makeAny(sal_Int8(10)) );

    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM ASIGNATURA WHERE GESTION=")) + ::rtl::OUString::number((sal_Int32) nGestion);
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));
        xItemList->setItemData( nIndex, ::com::sun::star::uno::makeAny(xRow->getString(1)));			
        nIndex = xItemList->getItemCount();
	}

    Reference< ::com::sun::star::uno::XInterface > xButtonModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropButtonModel ( xButtonModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference < ::com::sun::star::container::XNameContainer > xNameContainer ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), ::com::sun::star::uno::makeAny(xComboModel));
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), ::com::sun::star::uno::makeAny(xButtonModel));

    Reference < ::com::sun::star::uno::XInterface > xDialog (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XControl > xControl ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XControlModel > xControlModel ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference < ::com::sun::star::awt::XToolkit > xToolkit (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XWindow > xWindow ( xControl, ::com::sun::star::uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XButton > xButton ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XListBox > xListBox ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), ::com::sun::star::uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference < ::com::sun::star::awt::XDialog > xDlg ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::awt::XActionListener > xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //strResult = xListBox->getSelectedItem();
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= strResult;

    Reference < ::com::sun::star::lang::XComponent > xDispose ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return strResult;

}

::rtl::OUString 
EnsecProtocolHandler::getPlantilla(const Reference< ::com::sun::star::sdbc::XConnection >& xConnection)
{
    ::rtl::OUString strResult;
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::uno::XInterface > xDialogModel (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::beans::XPropertySet > xPropDlg ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );

    //xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
    //xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Plantilla"))) );

    Reference< ::com::sun::star::lang::XMultiServiceFactory > xSMDialog ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::uno::XInterface > xComboModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropComboModel ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    Reference < ::com::sun::star::awt::XItemList > xItemList ( xComboModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), ::com::sun::star::uno::makeAny(sal_True) );
    //xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ReadOnly")), ::com::sun::star::uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), ::com::sun::star::uno::makeAny(sal_Int8(10)) );

    ::rtl::OUString strSQL = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT PLANTILLA FROM COLUMNA GROUP BY PLANTILLA"));
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery(strSQL), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));
        xItemList->setItemData( nIndex, ::com::sun::star::uno::makeAny(xRow->getString(1)));			
        nIndex = xItemList->getItemCount();
	}

    Reference< ::com::sun::star::uno::XInterface > xButtonModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference < ::com::sun::star::beans::XPropertySet > xPropButtonModel ( xButtonModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference < ::com::sun::star::container::XNameContainer > xNameContainer ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), ::com::sun::star::uno::makeAny(xComboModel));
    xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), ::com::sun::star::uno::makeAny(xButtonModel));

    Reference < ::com::sun::star::uno::XInterface > xDialog (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XControl > xControl ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XControlModel > xControlModel ( xDialogModel, ::com::sun::star::uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference < ::com::sun::star::awt::XToolkit > xToolkit (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference < ::com::sun::star::awt::XWindow > xWindow ( xControl, ::com::sun::star::uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XButton > xButton ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference < ::com::sun::star::awt::XListBox > xListBox ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), ::com::sun::star::uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference < ::com::sun::star::awt::XDialog > xDlg ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::awt::XActionListener > xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //strResult = xListBox->getSelectedItem();
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= strResult;

    Reference < ::com::sun::star::lang::XComponent > xDispose ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return strResult;

}


Reference< ::com::sun::star::sheet::XSpreadsheetDocument >
EnsecProtocolHandler::getDocumentSheet()
{
    Reference< ::com::sun::star::sheet::XSpreadsheetDocument > xDocSheet;
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::frame::XDesktop > xDesktop ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW);    
    Reference< ::com::sun::star::frame::XFrame > xFrame ( xDesktop->getCurrentFrame(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::frame::XController > xController ( xFrame->getController(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::frame::XModel > xModel ( xController->getModel(), ::com::sun::star::uno::UNO_QUERY_THROW);

    xDocSheet.set ( xModel, ::com::sun::star::uno::UNO_QUERY_THROW );

    return xDocSheet;
}




Reference< ::com::sun::star::sheet::XSpreadsheet >
EnsecProtocolHandler::getCurrentSheet()
{
    Reference< ::com::sun::star::sheet::XSpreadsheet > xSheet;
    Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::frame::XDesktop > xDesktop ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW);    
    Reference< ::com::sun::star::frame::XFrame > xFrame ( xDesktop->getCurrentFrame(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::sheet::XSpreadsheetView > xSheetView ( xFrame->getController(), ::com::sun::star::uno::UNO_QUERY);    


    if ( !xSheetView.is() )
        return xSheet;

    xSheet.set( xSheetView->getActiveSheet(), ::com::sun::star::uno::UNO_QUERY );

    return xSheet;
}

void
EnsecProtocolHandler::createForm()
{
  Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager (mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::frame::XDesktop > xDesktop ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::frame::XFrame > xFrame ( xDesktop->getCurrentFrame(), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::frame::XController > xController ( xFrame->getController(), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::frame::XModel > xModel ( xController->getModel(), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::sheet::XSpreadsheetDocument > xSpreadsheetDocument ( xModel, ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::sheet::XSpreadsheets > xSpreadsheets ( xSpreadsheetDocument->getSheets(), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::container::XIndexAccess > xIndexAccess ( xSpreadsheets, ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::sheet::XSpreadsheet > xSheet ( xIndexAccess->getByIndex(0), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::container::XNamed > xNamed ( xSheet, ::com::sun::star::uno::UNO_QUERY_THROW);
  //::rtl::OUString strName = xNamed->getName();
  //Reference< ::com::sun::star::table::XCell > xSheetCell (xSheet->getCellByPosition(0,0), ::com::sun::star::uno::UNO_QUERY_THROW );
  //::com::sun::star::util::DateTime oDate, oDate0;

  //oDate.Year=2013;
  //oDate.Month=2;
  //oDate.Day=1;

  //xSheetCell->setValue( 0.0 );
  //xSheetCell->setFormula(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("2014/05/16")));

  //::com::sun::star::lang::Locale olocale;

  //Reference < ::com::sun::star::util::XNumberFormatsSupplier > xFormatsSupplier ( xSpreadsheetDocument, ::com::sun::star::uno::UNO_QUERY_THROW);
  //Reference < ::com::sun::star::util::XNumberFormatTypes > xFormatTypes ( xFormatsSupplier->getNumberFormats(), ::com::sun::star::uno::UNO_QUERY_THROW);
  //int nFormat = xFormatTypes->getStandardFormat(::com::sun::star::util::NumberFormat::DATE, olocale);
  //Reference < ::com::sun::star::beans::XPropertySet > xPropSet ( xSheetCell, ::com::sun::star::uno::UNO_QUERY_THROW);
  //xPropSet->setPropertyValue(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("NumberFormat")), ::com::sun::star::uno::makeAny(nFormat) );


  Reference < ::com::sun::star::uno::XInterface > xIDialog (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference < ::com::sun::star::beans::XPropertySet > xPropDlg ( xIDialog, ::com::sun::star::uno::UNO_QUERY_THROW );

  xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
  xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
  xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(150)) );
  xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
  xPropDlg->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Dialogo"))) );

  Reference< ::com::sun::star::lang::XMultiServiceFactory > xSMDialog ( xIDialog, ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::uno::XInterface > objButtonModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference < ::com::sun::star::beans::XPropertySet > xButtonModel ( objButtonModel , ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(20)) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(70)) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Button1"))) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Click Me"))) );

  Reference< ::com::sun::star::uno::XInterface > objLabelModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference < ::com::sun::star::beans::XPropertySet > xLabelModel ( objLabelModel, ::com::sun::star::uno::UNO_QUERY_THROW );  

  xLabelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(40)) );
  xLabelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(30)) );
  xLabelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(100)) );
  xLabelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
  xLabelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label1"))) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(1)) );
  xButtonModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Snake"))) );

  Reference< ::com::sun::star::uno::XInterface > objCancelModel (xSMDialog->createInstance(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference < ::com::sun::star::beans::XPropertySet > xCancelModel ( objCancelModel, ::com::sun::star::uno::UNO_QUERY_THROW );  
								     
  xCancelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), ::com::sun::star::uno::makeAny(sal_Int16(80)) );
  xCancelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), ::com::sun::star::uno::makeAny(sal_Int16(70)) );
  xCancelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(50)) );
  xCancelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), ::com::sun::star::uno::makeAny(sal_Int16(14)) );
  xCancelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel"))) );
  xCancelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), ::com::sun::star::uno::makeAny(sal_Int8(0)) );
  xCancelModel->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), ::com::sun::star::uno::makeAny(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel"))) );

  Reference < ::com::sun::star::container::XNameContainer > xNameContainer ( xIDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
  xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Button1")), ::com::sun::star::uno::makeAny(objButtonModel));
  xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Label1")), ::com::sun::star::uno::makeAny(objLabelModel));
  xNameContainer->insertByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel")), ::com::sun::star::uno::makeAny(objCancelModel));

  Reference < ::com::sun::star::uno::XInterface > xDialog (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
  
  Reference < ::com::sun::star::awt::XControl > xControl ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference < ::com::sun::star::awt::XControlModel > xControlModel ( xIDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
  xControl->setModel(xControlModel);

  Reference < ::com::sun::star::awt::XToolkit > xToolkit (xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

  Reference < ::com::sun::star::awt::XWindow > xWindow ( xControl, ::com::sun::star::uno::UNO_QUERY_THROW );
  xWindow->setVisible( sal_True );
  xControl->createPeer( xToolkit, 0 );

  Reference < ::com::sun::star::awt::XControlContainer > xControlContainer ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );

  Reference < ::com::sun::star::awt::XButton > xButton ( xControlContainer->getControl(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel"))), ::com::sun::star::uno::UNO_QUERY_THROW );
  xButton->setActionCommand(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OK")));

  Reference < ::com::sun::star::awt::XDialog > xDlg ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::awt::XActionListener > xEnsureDelete( new ClickHandler( xDlg ) );
  xButton->addActionListener( xEnsureDelete );

  xDlg->execute();

  Reference < ::com::sun::star::lang::XComponent > xDispose ( xDialog, ::com::sun::star::uno::UNO_QUERY_THROW );
  xDispose->dispose();
}

/*void 
EnsecProtocolHandler::showMessageBox(const ::rtl::OUString& strMessage)
{
 ::rtl::OUString strName = rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Informacion"));
    ::com::sun::star::uno::Reference< ::com::sun::star::lang::XMultiComponentFactory > smgr(mxContext->getServiceManager(), ::com::sun::star::uno::UNO_SET_THROW);
    ::com::sun::star::uno::Reference< ::com::sun::star::awt::XMessageBox > box( 
    ::com::sun::star::uno::Reference< ::com::sun::star::awt::XMessageBoxFactory >( smgr->createInstanceWithContext( rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW)->createMessageBox(
	::com::sun::star::uno::Reference< ::com::sun::star::awt::XWindowPeer >(::com::sun::star::uno::Reference< ::com::sun::star::frame::XFrame >(::com::sun::star::uno::Reference< ::com::sun::star::frame::XDesktop >(smgr->createInstanceWithContext(rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext),  ::com::sun::star::uno::UNO_QUERY_THROW)->getCurrentFrame(), ::com::sun::star::uno::UNO_SET_THROW)->getComponentWindow(), ::com::sun::star::uno::UNO_QUERY_THROW),
	::com::sun::star::awt::Rectangle(), rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("infobox")), 
    ::com::sun::star::awt::MessageBoxButtons::BUTTONS_OK, strName, strMessage), ::com::sun::star::uno::UNO_SET_THROW);

    box->execute();

    ::com::sun::star::uno::Reference< ::com::sun::star::lang::XComponent >(box, ::com::sun::star::uno::UNO_QUERY_THROW)->dispose();

    }*/

void
EnsecProtocolHandler::showMessageBox( const Reference< ::com::sun::star::awt::XToolkit >& rToolkit, const Reference< ::com::sun::star::frame::XFrame >& rFrame, const ::rtl::OUString& aTitle, const ::rtl::OUString& aMsgText )
{
    if ( rFrame.is() && rToolkit.is() )
    {
        // describe window properties.
        ::com::sun::star::awt::WindowDescriptor                aDescriptor;
        aDescriptor.Type              = ::com::sun::star::awt::WindowClass_MODALTOP;
        aDescriptor.WindowServiceName = "infobox";
        aDescriptor.ParentIndex       = -1;
        aDescriptor.Parent            = Reference< ::com::sun::star::awt::XWindowPeer >( rFrame->getContainerWindow(), UNO_QUERY );
        aDescriptor.Bounds            = ::com::sun::star::awt::Rectangle(0,0,300,200);
        aDescriptor.WindowAttributes  = ::com::sun::star::awt::WindowAttribute::BORDER |
            ::com::sun::star::awt::WindowAttribute::MOVEABLE |
            ::com::sun::star::awt::WindowAttribute::CLOSEABLE;

        Reference< ::com::sun::star::awt::XWindowPeer > xPeer = rToolkit->createWindow( aDescriptor );
        if ( xPeer.is() )
        {
            Reference< ::com::sun::star::awt::XMessageBox > xMsgBox( xPeer, UNO_QUERY );
            if ( xMsgBox.is() )
            {
                xMsgBox->setCaptionText( aTitle );
                xMsgBox->setMessageText( aMsgText );
                xMsgBox->execute();
            }
        }
    }
}


// -------------------------------------------------------------------------------------------------
// helper to get flisol database
//
Reference< ::com::sun::star::sdbc::XDataSource >
EnsecProtocolHandler::getDataSource( ) throw ( ::com::sun::star::uno::RuntimeException )
{
  Reference< ::com::sun::star::sdbc::XDataSource > xDataSource;
  Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::container::XNameAccess > xNameAccess ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );

  if ( xNameAccess->hasElements() && 
       xNameAccess->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ensec")))) 
      xDataSource.set ( xNameAccess->getByName( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ensec"))), ::com::sun::star::uno::UNO_QUERY_THROW );

  return xDataSource;
}

Reference< ::com::sun::star::sdb::XOfficeDatabaseDocument >
EnsecProtocolHandler::getDatabaseDocument( ) throw ( ::com::sun::star::uno::RuntimeException )
{
  Reference< ::com::sun::star::sdb::XOfficeDatabaseDocument > xDatabaseDocument;
  Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::container::XNameAccess > xNameAccess ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::sdb::XDocumentDataSource > xDocumentDataSource;

  if ( xNameAccess->hasElements() && 
       xNameAccess->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ensec"))))  {

      xDocumentDataSource.set ( xNameAccess->getByName( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ensec"))), ::com::sun::star::uno::UNO_QUERY_THROW );
      xDatabaseDocument.set ( xDocumentDataSource->getDatabaseDocument(), ::com::sun::star::uno::UNO_QUERY_THROW );
  }
  
  return xDatabaseDocument;
}



// ---------------------------------------------------------------------------------
// helper to check flisol table definition
//
void
EnsecProtocolHandler::checkTable( Reference< ::com::sun::star::sdbc::XDataSource > &aDataSource)
{
    aDataSource.is();
  bool bTableName = false;
  Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::container::XNameAccess > xDatabaseContext ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::sdbc::XDataSource > xDataSource ( xDatabaseContext->getByName( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Bibliography"))), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );

  Reference< ::com::sun::star::sdbcx::XTablesSupplier> xTSupplier( xConnection, UNO_QUERY);

  if ( xTSupplier.is() ){
    Reference< ::com::sun::star::container::XNameAccess > xNameAccess ( xTSupplier->getTables(), ::com::sun::star::uno::UNO_QUERY_THROW );    
    if (xNameAccess->hasElements() &&  xNameAccess->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("registro"))) ) {
      bTableName = true;
    }
  }

  if ( !bTableName ) {
    ::rtl::OUString sTable = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CREATE TABLE registro ( ID INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY, NOMBRE VARCHAR, APELLIDO VARCHAR, PROFESION VARCHAR, UNIVERSIDAD VARCHAR, PUBLICIDAD VARCHAR )"));
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW );
    xStatement->executeQuery( sTable );
  }
}

// -----------------------------------------------------------------------------------------------
// helper to check flisol column definitions
//
void 
EnsecProtocolHandler::checkColumnHeader( Reference < ::com::sun::star::sheet::XSpreadsheetDocument > & xDocument )
{
  Reference< ::com::sun::star::sheet::XSpreadsheets > xSheets ( xDocument->getSheets(), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::container::XIndexAccess > xIndex ( xSheets, ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::sheet::XSpreadsheet > xSheet ( xIndex->getByIndex(0), ::com::sun::star::uno::UNO_QUERY_THROW );

  if ( xSheet.is() ) {
    Reference< ::com::sun::star::sheet::XCellRangesQuery > xQuery ( xSheet, ::com::sun::star::uno::UNO_QUERY_THROW );

    const sal_Int16 nContentFlags = ::com::sun::star::sheet::CellFlags::STRING | 
      ::com::sun::star::sheet::CellFlags::VALUE | 
      ::com::sun::star::sheet::CellFlags::DATETIME | 
      ::com::sun::star::sheet::CellFlags::FORMULA | 
      ::com::sun::star::sheet::CellFlags::ANNOTATION;

    Reference< ::com::sun::star::sheet::XSheetCellRanges > xSheetCellRanges ( xQuery->queryContentCells( nContentFlags ), ::com::sun::star::uno::UNO_QUERY_THROW );    
    if ( xSheetCellRanges->getCount() > 0 ) {
      //Reference< ::com::sun::star::sheet::XCellRange > xCellRange ( xSheetCellRanges->getByIndex(0), ::com::sun::star::uno::UNO_QUERY_THROW );
      

    }
    else {


    }


  }
}

// ------------------------------------------------------------------------
// open flisol database form
//
void 
EnsecProtocolHandler::openForm( )
{
  Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager ( mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::uno::XInterface > xIDatabaseContext ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::container::XNameAccess > xNameAccess ( xIDatabaseContext, ::com::sun::star::uno::UNO_QUERY_THROW );

  if (xNameAccess->hasElements() && xNameAccess->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014"))) ) {
    Reference< ::com::sun::star::sdb::XDocumentDataSource > xDocumentDataSource (xNameAccess->getByName( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014")) ), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::sdb::XOfficeDatabaseDocument > xOfficeDbD ( xDocumentDataSource->getDatabaseDocument(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::sdb::XFormDocumentsSupplier > xFormDS ( xOfficeDbD, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::container::XNameAccess > xContainer ( xFormDS->getFormDocuments(), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::frame::XComponentLoader > xLoader ( xFormDS->getFormDocuments(), ::com::sun::star::uno::UNO_QUERY_THROW );
    
    if ( xContainer->hasElements() && xContainer->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO"))) ) {
      //Reference< ::com::sun::star::beans::XFastPropertySet > xFastProperty ( xContainer->getByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO"))), ::com::sun::star::uno::UNO_QUERY_THROW );
      //Reference< ::com::sun::star::lang::XTypeProvider > xTypeProvider ( xContainer->getByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO"))), ::com::sun::star::uno::UNO_QUERY_THROW );      
      Reference< ::com::sun::star::sdbc::XDataSource > xDataSource ( xOfficeDbD->getDataSource(), ::com::sun::star::uno::UNO_QUERY_THROW );
      Reference< ::com::sun::star::sdbc::XConnection> xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );

      Sequence < ::com::sun::star::beans::PropertyValue > aArgs(2);      
      aArgs[0].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("ActiveConnection"));
      aArgs[0].Value <<= xConnection;
      aArgs[1].Name = ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("OpenMode"));
      aArgs[1].Value <<= ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("open"));;

      //::com::sun::star::form::NavigationBarMode mode = ::com::sun::star::form::NavigationBarMode_CURRENT;

      //xFastProperty->setFastPropertyValue( 13, ::com::sun::star::uno::makeAny( ::com::sun::star::form::NavigationBarMode_CURRENT ) );
      //xFastProperty->setFastPropertyValue( 15, ::com::sun::star::uno::makeAny( sal_True ) );
      //xFastProperty->setFastPropertyValue( 16, ::com::sun::star::uno::makeAny( sal_True ) );
      //xFastProperty->setFastPropertyValue( 17, ::com::sun::star::uno::makeAny( sal_True ) );
      
      
      /*::com::sun::star::uno::Sequence < ::com::sun::star::uno::Type > seqTypes = xTypeProvider->getTypes();
      
            ::rtl::OString ouTypeName;
      for ( int nIterator = 0; nIterator < seqTypes.getLength(); nIterator++ ) {
	  ouTypeName = OUStringToOString( seqTypes[nIterator].getTypeName(), RTL_TEXTENCODING_ASCII_US );
	  }*/

      xLoader->loadComponentFromURL( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO")), 
				     ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("_blank")),
				     0,
				     aArgs );

    }
    else showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                         ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La tabla REGISTRO no encontrado")));
  }
  else showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                       ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Base de Datos Flisol2014 no encontrado")));
}

// -----------------------------------------------------------------------------
// open flisol sheet 
//
void 
EnsecProtocolHandler::openSheet()
{
  Reference< ::com::sun::star::lang::XMultiComponentFactory > xServiceManager (mxContext->getServiceManager(), ::com::sun::star::uno::UNO_QUERY_THROW);
  Reference< ::com::sun::star::uno::XInterface > iDatabaseContext ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW );
  Reference< ::com::sun::star::container::XNameAccess > xNameAccess ( iDatabaseContext, ::com::sun::star::uno::UNO_QUERY_THROW );

  if (xNameAccess->hasElements() && xNameAccess->hasByName(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014"))) ) {

    Reference< ::com::sun::star::sdbc::XDataSource >  xDataSource ( xNameAccess->getByName( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014")) ), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::sdbc::XConnection > xConnection ( xDataSource->getConnection( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("")), ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM(""))), ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::sdbc::XStatement > xStatement ( xConnection->createStatement(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::sdbc::XResultSet > xResultSet ( xStatement->executeQuery( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM REGISTRO"))), ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference< ::com::sun::star::frame::XDesktop > xDesktop ( xServiceManager->createInstanceWithContext(::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), ::com::sun::star::uno::UNO_QUERY_THROW);    
    Reference< ::com::sun::star::frame::XFrame > xFrame ( xDesktop->getCurrentFrame(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::frame::XController > xController ( xFrame->getController(), ::com::sun::star::uno::UNO_QUERY_THROW);    
    Reference< ::com::sun::star::frame::XModel > xModel ( xController->getModel(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::sheet::XSpreadsheetDocument > xSpreadsheetDocument ( xModel, ::com::sun::star::uno::UNO_QUERY );

    if ( !xSpreadsheetDocument.is() ) {
        showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede abrir un documento de Calc")));
        return;
    }

    Reference< ::com::sun::star::sheet::XSpreadsheets > xSpreadsheets ( xSpreadsheetDocument->getSheets(), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::container::XIndexAccess > xIndexAccess ( xSpreadsheets, ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::sheet::XSpreadsheet > xSheet ( xIndexAccess->getByIndex(0), ::com::sun::star::uno::UNO_QUERY_THROW);
    Reference< ::com::sun::star::table::XColumnRowRange > xColsRows( xSheet, ::com::sun::star::uno::UNO_QUERY_THROW );

    Reference< ::com::sun::star::sdbcx::XColumnsSupplier > xColumns ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );
    Reference< ::com::sun::star::container::XNameAccess > xColumnNames ( xColumns->getColumns(), ::com::sun::star::uno::UNO_QUERY_THROW );    
    Reference< ::com::sun::star::table::XTableColumns > xTableColumns ( xColsRows->getColumns(), ::com::sun::star::uno::UNO_QUERY_THROW );    
    Sequence < ::rtl::OUString > arrayNames = xColumnNames->getElementNames();

    sal_Int32 nRowColumn = 0;
    sal_Int32 nRow = 1;

    for ( sal_Int32 nIterator = 0; nIterator < arrayNames.getLength(); nIterator++ ) {
	//Reference< ::com::sun::star::sdb::XColumn > xColumn ( xColumnIndex->getByIndex( nIterator ), ::com::sun::star::uno::UNO_QUERY_THROW );
      Reference< ::com::sun::star::table::XCell > xCell ( xSheet->getCellByPosition( nIterator, nRowColumn ), ::com::sun::star::uno::UNO_QUERY_THROW );
      //Reference< ::com::sun::star::table::XTableColumn > xTableColumn ( xTableColumns->getByIndex( nIterator ), ::com::sun::star::uno::UNO_QUERY_THROW );
      Reference< ::com::sun::star::beans::XPropertySet > xColumnProp ( xTableColumns->getByIndex( nIterator ), ::com::sun::star::uno::UNO_QUERY_THROW );
      Reference< ::com::sun::star::beans::XPropertySet > xCellProp ( xCell, ::com::sun::star::uno::UNO_QUERY_THROW );
      if ( arrayNames[nIterator].compareToAscii("ID") == 0) {
	xCell->setFormula( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("No.")) );
	xColumnProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(2000)) );
      }
      else {
	xCell->setFormula( arrayNames[nIterator] );	
	xColumnProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), ::com::sun::star::uno::makeAny(sal_Int16(6000)) );
      }
      xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
				   ::com::sun::star::uno::makeAny(::com::sun::star::awt::FontWeight::BOLD) );
      xCellProp->setPropertyValue( ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
				   ::com::sun::star::uno::makeAny(::com::sun::star::table::CellHoriJustify_CENTER) );
    }

    Reference< ::com::sun::star::sdbc::XRow > xRow ( xResultSet, ::com::sun::star::uno::UNO_QUERY_THROW );  

    while ( xResultSet->next() ) {
      for ( sal_Int32 nIterator = 1; nIterator <= arrayNames.getLength(); nIterator++) {
	Reference< ::com::sun::star::table::XCell > xCell ( xSheet->getCellByPosition( nIterator - 1, nRow ), ::com::sun::star::uno::UNO_QUERY_THROW );	      
	Reference< ::com::sun::star::beans::XPropertySet > xCellProp ( xCell, ::com::sun::star::uno::UNO_QUERY_THROW );	      
	if ( arrayNames[nIterator - 1].compareToAscii("ID") == 0) {
	  xCell->setValue( nRow );
	}
	else {
	  xCell->setFormula( xRow->getString(nIterator) );			
	}
      }
      nRow++;
    }
  }
  else showMessageBox (mxToolkit, mxFrame, 
                        ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                       ::rtl::OUString(RTL_CONSTASCII_USTRINGPARAM("La Base de Datos Flisol2014 no encontrado")));
}

// ------------------------------------------------------------------------------------------------
//
//
void SAL_CALL 
EnsecProtocolHandler::addStatusListener( const Reference< XStatusListener >& xControl, 
					  const URL& aURL )
  throw (RuntimeException)
{
    xControl.is();
    aURL.Complete.isEmpty();
}

// -------------------------------------------------------------------------------------------------
//
//
void SAL_CALL 
EnsecProtocolHandler::removeStatusListener( const Reference< XStatusListener >& xControl, 
					     const URL& aURL ) 
  throw (RuntimeException)
{
    xControl.is();
    aURL.Complete.isEmpty();
}
