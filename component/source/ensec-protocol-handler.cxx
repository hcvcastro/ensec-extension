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

#ifndef _COM_SUN_STAR_XML_SAX_XEXTENDEDDOCUMENTHANDLER_HPP_
#include <com/sun/star/xml/sax/XExtendedDocumentHandler.hpp>
#endif

namespace uno = com::sun::star::uno;
namespace frame = com::sun::star::frame;
namespace util = com::sun::star::util;
namespace sax = com::sun::star::xml::sax;
namespace container = com::sun::star::container;
namespace awt = com::sun::star::awt;
namespace lang = com::sun::star::lang;
namespace sdbc = com::sun::star::sdbc;
namespace beans = ::com::sun::star::beans;
namespace io = ::com::sun::star::io;
namespace sheet = ::com::sun::star::sheet;
namespace table = ::com::sun::star::table;
namespace xml = ::com::sun::star::xml;
namespace sdb = ::com::sun::star::sdb;
namespace sdbcx = ::com::sun::star::sdbcx;


typedef uno::Reference<sax::XDocumentHandler> pFuncImportDialogModel (
    uno::Reference<container::XNameContainer> const & xDialogModel,
    uno::Reference<uno::XComponentContext> const & xContext,
    uno::Reference<frame::XModel> const & xDocument );


typedef ::cppu::WeakImplHelper1 <awt::XActionListener> ClickHandlerBase;
typedef ::cppu::WeakImplHelper1 <awt::XTextListener> EditHandlerBase;

class ClickHandler : public ClickHandlerBase
{
private:
  Reference <awt::XDialog> mxDialog;
public:
  ClickHandler( Reference <awt::XDialog> &aDialog ) : mxDialog(aDialog) {};
protected:
  // XActionListener
  virtual void SAL_CALL actionPerformed( const awt::ActionEvent& rEvent ) throw (RuntimeException);
  // XEventListener
  virtual void SAL_CALL disposing( const lang::EventObject& Source ) throw (RuntimeException);
};

void SAL_CALL ClickHandler::actionPerformed( const awt::ActionEvent& rEvent ) throw (RuntimeException)
{
  if (rEvent.ActionCommand.compareToAscii("OK")==0) {
    mxDialog->endExecute();
  }
  else
    mxDialog->endExecute();
}

void SAL_CALL ClickHandler::disposing( const lang::EventObject& /*Source*/ ) throw (RuntimeException)
{
  // not interested in
}


class UpdateHandler : public ClickHandlerBase
{
private:
    Reference <sdbc::XConnection> mxConnection;
    Reference <awt::XDialog> mxDialog;
    sal_Int32 mnGestion;
    sal_Int32 mnPeriodo;
    OUString mstrAsignatura;
public:
    UpdateHandler( const Reference<sdbc::XConnection>& aConnection, 
                   const Reference <awt::XDialog>& aDialog,
                   sal_Int32 aGestion,
                   sal_Int32 aPeriodo,
                   OUString astrAsignatura ) : mxConnection(aConnection), 
                                                      mxDialog(aDialog), 
                                                      mnGestion(aGestion), 
                                                      mnPeriodo(aPeriodo), 
                                                      mstrAsignatura(astrAsignatura)
    {}
protected:
    // XActionListener
    virtual void SAL_CALL actionPerformed( const awt::ActionEvent& rEvent ) throw (RuntimeException);
    // XEventListener
    virtual void SAL_CALL disposing( const lang::EventObject& Source ) throw (RuntimeException);
};

void SAL_CALL UpdateHandler::actionPerformed( const awt::ActionEvent& rEvent ) throw (RuntimeException)
{
    Reference <awt::XControlContainer> xControlContainer ( mxDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XControl> xControl ( rEvent.Source , uno::UNO_QUERY );
    Reference <awt::XTextComponent> xText;
    Reference <awt::XFixedText> xFixedText;
    Reference <beans::XPropertySet> xPropFixedText;  
    Reference <awt::XButton> xButton ( xControl , uno::UNO_QUERY );
    Reference <beans::XPropertySet> xPropButton ( xControl->getModel(), uno::UNO_QUERY );
    OUString strLabel;

    xPropButton->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Label"))) >>= strLabel;

    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT TITULO FROM EVALUACION WHERE GESTION=")) +
        OUString::number(mnGestion) +
        OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
        mstrAsignatura +
        OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO =")) +
        OUString::number(mnPeriodo);


    Reference<sdbc::XStatement> xStatement ( mxConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference<sdbc::XStatement> xUpdate ( mxConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference<sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference<sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    OUString strTag;
    OUString strInsert;
    OUString strUpdate;
    OUString strEstudiante;
    OUString strEvaluacion;
    OString strDebug;
    

    if ( strLabel.compareToAscii("Insert") == 0 ) {
        while (xResultSet->next()){
            strInsert = OUString(RTL_CONSTASCII_USTRINGPARAM("INSERT INTO NOTA VALUES ( NULL,")) +
                OUString::number(mnGestion) + OUString(RTL_CONSTASCII_USTRINGPARAM(", '")) +
                mstrAsignatura + OUString(RTL_CONSTASCII_USTRINGPARAM("', ")) +
                OUString::number(mnPeriodo) + OUString(RTL_CONSTASCII_USTRINGPARAM(", ")) ;

            xControl.set (xControlContainer->getControl( xRow->getString(1) + OUString(RTL_CONSTASCII_USTRINGPARAM("_label"))), 
                          uno::UNO_QUERY );
            xFixedText.set ( xControl, uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel() , uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                strInsert += strTag + OUString(RTL_CONSTASCII_USTRINGPARAM(", ")) ;
            }

            xControl.set (xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))), 
                          uno::UNO_QUERY );
            xFixedText.set ( xControl, uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel() , uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                strInsert += strTag + OUString(RTL_CONSTASCII_USTRINGPARAM(", ")) ;
            }

            // 7 Fecha
            strInsert += OUString(RTL_CONSTASCII_USTRINGPARAM("NULL, ")) ;

            xControl.set (xControlContainer->getControl( xRow->getString(1)), 
                          uno::UNO_QUERY );
            xText.set ( xControl, uno::UNO_QUERY );
            if ( xControl.is() && xText.is() ) {
                strTag = xText->getText();
                strInsert += strTag + OUString(RTL_CONSTASCII_USTRINGPARAM(" )")) ;                
            }

            strDebug = OUStringToOString(strInsert, RTL_TEXTENCODING_UTF8);
            xUpdate->executeUpdate( strInsert );
        }
        xButton->setLabel(OUString(RTL_CONSTASCII_USTRINGPARAM("Done")));
        xPropButton->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), uno::makeAny(sal_False));
    }
    else if ( strLabel.compareToAscii("Update") == 0 ) {
        while (xResultSet->next()){

            strUpdate = OUString(RTL_CONSTASCII_USTRINGPARAM("UPDATE NOTA SET FECHA=NULL, NOTA="));

            xControl.set (xControlContainer->getControl(xRow->getString(1) + OUString(RTL_CONSTASCII_USTRINGPARAM("_label"))), 
                          uno::UNO_QUERY );
            xFixedText.set ( xControl, uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel(), uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strEvaluacion;
            }

            xControl.set (xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))), 
                          uno::UNO_QUERY );
            xFixedText.set ( xControl, uno::UNO_QUERY );
            if ( xControl.is() && xFixedText.is() ) {
                xPropFixedText.set ( xControl->getModel() , uno::UNO_QUERY );              
                xPropFixedText->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strEstudiante;
            }

            xControl.set (xControlContainer->getControl(xRow->getString(1)), 
                          uno::UNO_QUERY );
            xText.set ( xControl, uno::UNO_QUERY );
            if ( xControl.is() && xText.is() ) {
                strUpdate += xText->getText() + OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) ;                
            }

            strUpdate += OUString(RTL_CONSTASCII_USTRINGPARAM("  WHERE GESTION=")) +
                OUString::number(mnGestion) + OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +                
                mstrAsignatura + OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO=")) +
                OUString::number(mnPeriodo) + OUString(RTL_CONSTASCII_USTRINGPARAM(" AND EVALUACION=")) + 
                strEvaluacion + OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ESTUDIANTE=")) + 
                strEstudiante;

            strDebug = OUStringToOString(strUpdate, RTL_TEXTENCODING_UTF8);
            xUpdate->executeUpdate( strUpdate );
            
        }
        xButton->setLabel(OUString(RTL_CONSTASCII_USTRINGPARAM("Done")));
        xPropButton->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), uno::makeAny(sal_False));
    }
}

void SAL_CALL UpdateHandler::disposing( const lang::EventObject& /*Source*/ ) throw (RuntimeException)
{
  // not interested in
}



class EditHandler : public EditHandlerBase
{
private:
    Reference <sdbc::XConnection> mxConnection;
    Reference <awt::XDialog> mxDialog;
    sal_Int32 mnGestion;
    sal_Int32 mnPeriodo;
    OUString mstrAsignatura;

public:
    EditHandler( const Reference<sdbc::XConnection>& aConnection, 
                 const Reference <awt::XDialog>& aDialog,
                 sal_Int32 aGestion,
                 sal_Int32 aPeriodo,
                 OUString astrAsignatura ) : mxConnection(aConnection), mxDialog(aDialog), mnGestion(aGestion), mnPeriodo(aPeriodo), mstrAsignatura(astrAsignatura) {};
protected:
  // XTextListener           -> modify setzen
  virtual void SAL_CALL textChanged(const awt::TextEvent& rEvent) throw( uno::RuntimeException );
  // XEventListener
  virtual void SAL_CALL disposing( const lang::EventObject& Source ) throw (RuntimeException);

};
 

void SAL_CALL EditHandler::textChanged(const awt::TextEvent& rEvent) throw( uno::RuntimeException )
{
    util::Color clrGreenColor = 0x0000FF00;
    util::Color clrRedColor = 0x00FF0000;
    //util::Color clrBlueColor = 0x000000FF;
    sal_Int32 nIDEstudiante = -1;
    bool bNotas = false;

    Reference <awt::XControlContainer> xControlContainer ( mxDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XTextComponent> xText ( rEvent.Source , uno::UNO_QUERY_THROW );
    Reference <awt::XButton> xButton;
    Reference <awt::XTextComponent> xEditText;
    Reference <awt::XFixedText> xFixedText (xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))), uno::UNO_QUERY_THROW );

    Reference <awt::XControl> xControl ( xFixedText, uno::UNO_QUERY_THROW );
    Reference <beans::XPropertySet> xPropTextModel ( xControl->getModel() , uno::UNO_QUERY_THROW );  

    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.NOMBRE, ESTUDIANTE.ESTUDIANTE_ID FROM INSCRIPCION INNER JOIN ESTUDIANTE ON INSCRIPCION.GESTION=")) +
          OUString::number(mnGestion) +
          OUString(RTL_CONSTASCII_USTRINGPARAM(" AND INSCRIPCION.ASIGNATURA='")) +
          mstrAsignatura +
           OUString(RTL_CONSTASCII_USTRINGPARAM("' AND INSCRIPCION.ESTUDIANTE = ESTUDIANTE.ESTUDIANTE_ID AND ESTUDIANTE.NOMBRE LIKE '%")) + 
          xText->getText().toAsciiUpperCase() +
           OUString(RTL_CONSTASCII_USTRINGPARAM("%'"));

    Reference<sdbc::XStatement> xStatement ( mxConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference<sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference<sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    OUString strFound = OUString(RTL_CONSTASCII_USTRINGPARAM("ESTUDIANTE NO ENCONTRADO!"));
    sal_Int32 nCounter = 0;

    while ( xResultSet->next() ) {
        if ( xResultSet->isFirst())
            strFound = xRow->getString(1);
        else
            strFound += OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) + xRow->getString(1);

        nIDEstudiante = xRow->getInt(2);
        nCounter++;
    }


    if ( nCounter == 1 ) {
        strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT TITULO FROM EVALUACION WHERE GESTION=")) +
            OUString::number(mnGestion) +
            OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
            mstrAsignatura +
            OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO =")) +
            OUString::number(mnPeriodo);

        //xStatement.set ( mxConnection->createStatement(), uno::UNO_QUERY_THROW );
        xResultSet.set ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
        xRow.set ( xResultSet, uno::UNO_QUERY_THROW );  


        while ( xResultSet->next() ) {
            xEditText.set (xControlContainer->getControl( xRow->getString(1) ), uno::UNO_QUERY );
            if ( xEditText.is() )
                xEditText->setText( OUString(RTL_CONSTASCII_USTRINGPARAM("0")));
        }


        strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT EVALUACION.TITULO, NOTA.* FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION=EVALUACION.ID AND NOTA.GESTION=")) +
            OUString::number(mnGestion) +
            OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA='")) +
            mstrAsignatura +
            OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO =")) + 
            OUString::number(mnPeriodo) +
            OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ESTUDIANTE =")) + 
            OUString::number(nIDEstudiante);

        //xStatement.set ( mxConnection->createStatement(), uno::UNO_QUERY_THROW );
        xResultSet.set ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
        xRow.set ( xResultSet, uno::UNO_QUERY_THROW );  

        while ( xResultSet->next() ) {
            bNotas = true;
            xEditText.set (xControlContainer->getControl( xRow->getString(1) ), uno::UNO_QUERY );
            if ( xEditText.is() )
                xEditText->setText( OUString::number(xRow->getInt(9)));
        }
    } else {
        strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT TITULO FROM EVALUACION WHERE GESTION=")) +
            OUString::number(mnGestion) +
            OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
            mstrAsignatura +
            OUString(RTL_CONSTASCII_USTRINGPARAM("' AND PERIODO =")) +
            OUString::number(mnPeriodo);

        //xStatement.set ( mxConnection->createStatement(), uno::UNO_QUERY_THROW );
        xResultSet.set ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
        xRow.set ( xResultSet, uno::UNO_QUERY_THROW );  


        while ( xResultSet->next() ) {
            xEditText.set (xControlContainer->getControl( xRow->getString(1) ), uno::UNO_QUERY );
            if ( xEditText.is() )
                xEditText->setText( OUString(RTL_CONSTASCII_USTRINGPARAM("0")));
        }
    }

    OUString strTag;
    Reference <beans::XPropertySet> xPropControl;

    if (nCounter == 1) {
        xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TextColor")), uno::makeAny(clrGreenColor) );
        Sequence<Reference<awt::XControl>> controls = xControlContainer->getControls();

        for( sal_Int32 nIterator = 0; nIterator < controls.getLength(); nIterator++) {
            xPropControl.set (controls[nIterator]->getModel(), uno::UNO_QUERY );
            if (xPropControl.is()) {
                xPropControl->getPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                if ( strTag.getLength() != 0 ) {
                    xPropControl->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), uno::makeAny( sal_True ) );
                }
            }
        }
    }
    else  {
        xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TextColor")), uno::makeAny(clrRedColor) );
        Sequence<Reference<awt::XControl>> controls = xControlContainer->getControls();

        for( sal_Int32 nIterator = 0; nIterator < controls.getLength(); nIterator++) {
            xPropControl.set (controls[nIterator]->getModel(), uno::UNO_QUERY );
            if (xPropControl.is()) {
                xPropControl->getPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                if ( strTag.getLength() != 0 )
                    xPropControl->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), uno::makeAny( sal_False ) );
            }
        }
    }

    xButton.set (xControlContainer->getControl( OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate"))), uno::UNO_QUERY);
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), uno::makeAny(OUString::number(nIDEstudiante)) );
    xFixedText->setText(strFound);

    if ( bNotas )
        xButton->setLabel (OUString(RTL_CONSTASCII_USTRINGPARAM("Update")));
    else
        xButton->setLabel (OUString(RTL_CONSTASCII_USTRINGPARAM("Insert")));

}

void SAL_CALL EditHandler::disposing( const lang::EventObject& /*Source*/ ) throw (RuntimeException)
{
  // not interested in
}


class NotaHandler : public EditHandlerBase
{
private:
    Reference <awt::XDialog> mxDialog;

public:
    NotaHandler( const Reference <awt::XDialog>& aDialog ) : mxDialog(aDialog) {};
protected:
    // XTextListener           -> modify setzen
    virtual void SAL_CALL textChanged(const awt::TextEvent& rEvent) throw( uno::RuntimeException );
    // XEventListener
    virtual void SAL_CALL disposing( const lang::EventObject& Source ) throw (RuntimeException);
    static void updateNotaFinal( const Reference <awt::XDialog>& );

};

void SAL_CALL NotaHandler::textChanged(const awt::TextEvent& rEvent) throw( uno::RuntimeException )
{
    float fPonderacion = 0.0;
    float fNota = 0.0;
    OUString strTag;
    OUString strNota;

    Reference <awt::XControlContainer> xControlContainer ( mxDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XTextComponent> xTextComponent ( rEvent.Source , uno::UNO_QUERY_THROW );
    Reference <awt::XControl> xControlText ( xTextComponent , uno::UNO_QUERY_THROW );
    Reference <beans::XPropertySet> xPropTextModel ( xControlText->getModel() , uno::UNO_QUERY_THROW );  

    xPropTextModel->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
    Reference <awt::XFixedText> xFixedText (xControlContainer->getControl(strTag), uno::UNO_QUERY );
    
    if ( xFixedText.is() ) {
        Reference <awt::XControl> xControlFixedText ( xFixedText, uno::UNO_QUERY );
        Reference <beans::XPropertySet> xPropFixedTextModel ( xControlFixedText->getModel(), uno::UNO_QUERY_THROW );  

        xPropFixedTextModel->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
        fPonderacion = strTag.toFloat();
        strNota = xTextComponent->getText();
        fNota = strNota.toFloat();
        fNota = fNota * fPonderacion;
    }

    xFixedText->setText( OUString(RTL_CONSTASCII_USTRINGPARAM(" = ")) + OUString::number( fNota ) );

    NotaHandler::updateNotaFinal( mxDialog );

}


void NotaHandler::updateNotaFinal(const Reference <awt::XDialog>& aDialog )
{
    Reference <awt::XControlContainer> xControlContainer ( aDialog, uno::UNO_QUERY_THROW );
    Sequence<Reference<awt::XControl>> controls = xControlContainer->getControls();
    Reference<awt::XControl> xControl;
    Reference<awt::XTextComponent> xEditControl;
    Reference <beans::XPropertySet> xPropEditControl;  
    Reference <awt::XFixedText> xFixedText;
    Reference <beans::XPropertySet> xPropFixedControl;  
    float fPonderacion;
    float fNota;
    float fNotaFinal = 0.0;
    OUString strTag;

    for( sal_Int32 nIterator = 0; nIterator < controls.getLength(); nIterator++) {
        fNota = 0.0;
        fPonderacion = 0.0;
        xControl.set ( controls[nIterator], uno::UNO_QUERY );
        xEditControl.set ( xControl, uno::UNO_QUERY );
        if (xEditControl.is()) {
            xPropEditControl.set ( xControl->getModel(), uno::UNO_QUERY );
            fNota = xEditControl->getText().toFloat();            
            xPropEditControl->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
            xControl.set (xControlContainer->getControl (strTag), uno::UNO_QUERY );
            xFixedText.set ( xControl, uno::UNO_QUERY );
            if (xControl.is() && xFixedText.is()) {
                xPropFixedControl.set (xControl->getModel(), uno::UNO_QUERY );
                xPropFixedControl->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= strTag;
                fPonderacion = strTag.toFloat();
                fNota = fNota * fPonderacion;
            }
            else fNota = 0.0;
        }
        fNotaFinal += fNota;
    }

    xFixedText.set( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("nota_final"))), uno::UNO_QUERY );
    xFixedText->setText(OUString::number( fNotaFinal));
}


void SAL_CALL NotaHandler::disposing( const lang::EventObject& /*Source*/ ) throw (RuntimeException)
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
EnsecProtocolHandler::initialize( const Sequence<Any>& aArguments ) 
  throw ( Exception, RuntimeException )
{
    Reference <frame::XFrame> xFrame;
  if ( aArguments.getLength() ) {
    aArguments[0] >>= xFrame;
    mxFrame = xFrame;
  }

  // Create the toolkit to have access to it later
  mxToolkit = Reference<awt::XToolkit>( awt::Toolkit::create(mxContext), UNO_QUERY_THROW );
}


// XDispatchProvider
// -------------------------------------------------------------------------------------------
// Searches for an XDispatch for the specified URL within the specified target frame. 
//
Reference< frame::XDispatch > SAL_CALL 
EnsecProtocolHandler::queryDispatch( const util::URL& aURL, 
				      const OUString& sTargetFrameName, 
				      sal_Int32 nSearchFlags )
  throw( RuntimeException )
{
    Reference <frame::XDispatch> xRet;

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
Sequence <Reference<frame::XDispatch>> SAL_CALL 
EnsecProtocolHandler::queryDispatches( const Sequence <frame::DispatchDescriptor>& seqDescripts )
  throw( RuntimeException )
{
  sal_Int32 nCount = seqDescripts.getLength();
  Sequence <Reference<frame::XDispatch>> lDispatcher( nCount );
 
  for( sal_Int32 i=0; i<nCount; ++i )
    lDispatcher[i] = queryDispatch( seqDescripts[i].FeatureURL, seqDescripts[i].FrameName, seqDescripts[i].SearchFlags );
 
  return lDispatcher;
}


// XServiceInfo
// ---------------------------------------------------------------
// Provides the implementation name of the service implementation.
//
OUString SAL_CALL 
EnsecProtocolHandler::getImplementationName(  )
  throw (RuntimeException)
{
  return EnsecProtocolHandler_getImplementationName();
}

// -------------------------------------------------------------------------
// Tests whether the specified service is supported, i.e. implemented by the implementation.
//
sal_Bool SAL_CALL 
EnsecProtocolHandler::supportsService( const OUString& rServiceName )
  throw (RuntimeException)
{
  return EnsecProtocolHandler_supportsService( rServiceName );
}

// -------------------------------------------------------------------------------------------------
// Provides the supported service names of the implementation, including also indirect service names.
//
Sequence< OUString > SAL_CALL 
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
EnsecProtocolHandler::dispatch( const util::URL& aURL, 
				 const Sequence <beans::PropertyValue>& lArgs ) 
  throw (RuntimeException)
{
    lArgs.getLength();

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

Reference<container::XNameContainer> lcl_createControlModel(
  const Reference<XComponentContext>& xContext) {
    Reference<lang::XMultiComponentFactory> xSMgr_( xContext->getServiceManager(), UNO_QUERY_THROW );
    Reference<container::XNameContainer> xControlModel( xSMgr_->createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", xContext ), UNO_QUERY_THROW );
    return xControlModel;
}

Reference<container::XNameContainer> lcl_createDialogModel(
        const Reference< XComponentContext>& xContext,
        const Reference<io::XInputStream>& xInput,
        const Reference<frame::XModel>& xModel
        ) throw ( Exception ) {
    
    Reference<container::XNameContainer> xDialogModel( lcl_createControlModel(xContext) );

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
    Reference<sdbc::XDataSource> xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference<sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }

    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    // get periodo data
    sal_Int32 nPeriodo = getPeriodo(xConnection);
    if (nPeriodo == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("El periodo no ha sido seleccionado!")));
        return;
    }


    // get Plantilla
    OUString strPlantilla = getPlantilla(xConnection);
    if ( strPlantilla.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La Plantilla no ha sido seleccionada!")));
        return;
    }


    // get Document Sheet 
    Reference<sheet::XSpreadsheetDocument> xDocSheet = getDocumentSheet();
    if (!xDocSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Documento Sheet Actual!")));
        return;
    }

    // get Current Sheet 
    Reference<sheet::XSpreadsheet> xSheet = getCurrentSheet();

    if (!xSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Sheet Actual!")));
        return;
    }

    sheetNotas(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, strPlantilla );

    showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                    OUString(RTL_CONSTASCII_USTRINGPARAM("Sheet Notas generado!")));

}

void
EnsecProtocolHandler::sheetNotas ( const Reference<sdbc::XConnection>& xConnection,
                                   const Reference<sheet::XSpreadsheet>& xSheet,
                                   sal_Int32 nGestion,
                                   const OUString& strAsignatura,
                                   sal_Int32 nPeriodo,
                                   const OUString& strPlantilla)
{

    sal_Int32 nRow = sheetNotasHeader(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, strPlantilla);
    sheetNotasRow(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, nRow);
}



void
EnsecProtocolHandler::sheetNotasEstudiante ( const Reference<sdbc::XConnection>& xConnection, 
                                             const Reference<sheet::XSpreadsheet >& xSheet,
                                             sal_Int32 nGestion,                                         
                                             const OUString& strAsignatura,                                         
                                             sal_Int32 nPeriodo,
                                             sal_Int32 nEstudiante,
                                             sal_Int32 nCol,
                                             sal_Int32 nRow)
{
    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT EVALUACION.TITULO, NOTA.NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.GESTION = EVALUACION.GESTION AND NOTA.ASIGNATURA = EVALUACION.ASIGNATURA AND NOTA.PERIODO = EVALUACION.PERIODO AND NOTA.EVALUACION = EVALUACION.ID  WHERE NOTA.GESTION = ? AND NOTA.ASIGNATURA = ? AND NOTA.PERIODO = ? AND NOTA.ESTUDIANTE = ?"));

    Reference<sdbc::XPreparedStatement> xPreparedStatement ( xConnection->prepareStatement(strSQL), uno::UNO_QUERY );
    Reference<sdbc::XParameters> xParameters ( xPreparedStatement, uno::UNO_QUERY );

    xParameters->setInt(1, nGestion);
    xParameters->setString(2, strAsignatura);
    xParameters->setInt(3, nPeriodo);
    xParameters->setInt(4, nEstudiante);

    Reference<sdbc::XResultSet> xResult = xPreparedStatement->executeQuery();
    Reference<sdbc::XRow> xRow ( xResult, uno::UNO_QUERY );  
    Reference<sdbc::XColumnLocate> xColumn ( xResult, uno::UNO_QUERY );
    Reference<table::XCell> xCell;
    Reference<beans::XPropertySet> xCellProp; 

    while (xResult->next()) {
        nCol = findColumn(xSheet, xRow->getString(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("TITULO")))));
        if ( nCol != -1 ) {
            xCell.set (xSheet->getCellByPosition(nCol, nRow), uno::UNO_QUERY);
            xCellProp.set (xCell, uno::UNO_QUERY);
            xCellProp->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                        uno::makeAny(sal_Int32(2)) );
            xCellProp->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                        uno::makeAny(sal_Int32(2)) );
            xCell->setValue(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("NOTA")))));
        }
    }
}

sal_Int32 EnsecProtocolHandler::findColumn ( const Reference<sheet::XSpreadsheet>& xSheet,
                                             const OUString& strColumn )
{
    sal_Int32 nCol = 0;
    sal_Int32 nRet = -1;
    Reference<sheet::XCellAddressable> xCellAddress;
    Reference<table::XCell> xCellFound;
    Reference<table::XCell> xCell;
    Reference<beans::XPropertySet> xCellProp; 
    Reference<container::XNameContainer> xAttributes;
    xml::AttributeData rAttribute;
    table::CellAddress rCellAddress;

    xCell.set (xSheet->getCellByPosition(nCol, 0), uno::UNO_QUERY);
    xCellProp.set (xCell, uno::UNO_QUERY);

    while (xCell->getType() != table::CellContentType_EMPTY) {
        xCellProp->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("UserDefinedAttributes"))) >>= xAttributes;        
        if (xAttributes.is() && xAttributes->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")))) {
            xAttributes->getByName(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag"))) >>= rAttribute;
            if ( rAttribute.Value.compareTo(strColumn)==0) {
                xCellFound = xCell;
                break;
            }
        }
        
        xCell.set (xSheet->getCellByPosition(++nCol, 0), uno::UNO_QUERY);
        xCellProp.set (xCell, uno::UNO_QUERY);
    }

    if ( xCellFound.is () ) {
        xCellAddress.set (xCellFound, uno::UNO_QUERY);
        rCellAddress = xCellAddress->getCellAddress();
        nRet = rCellAddress.Column;
    }

    return nRet;
}


void
EnsecProtocolHandler::sheetNotasRow( const Reference<sdbc::XConnection>& xConnection, 
                                     const Reference<sheet::XSpreadsheet >& xSheet,
                                     sal_Int32 nGestion,                                         
                                     const OUString& strAsignatura,                                         
                                     sal_Int32 nPeriodo,
                                     sal_Int32 nRow)
{
    sal_Int32 nCol = 0;
    sal_Int32 nIterator = 1;
    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.* FROM INSCRIPCION INNER JOIN ESTUDIANTE ON INSCRIPCION.ESTUDIANTE = ESTUDIANTE.ESTUDIANTE_ID WHERE INSCRIPCION.GESTION = ? AND INSCRIPCION.ASIGNATURA = ? ORDER BY ESTUDIANTE.NOMBRE"));

    Reference<sdbc::XPreparedStatement> xPreparedStatement ( xConnection->prepareStatement(strSQL), uno::UNO_QUERY );
    Reference<sdbc::XParameters> xParameters ( xPreparedStatement, uno::UNO_QUERY );

    xParameters->setInt(1, nGestion);
    xParameters->setString(2, strAsignatura);

    Reference<sdbc::XResultSet> xResult = xPreparedStatement->executeQuery();
    Reference<sdbc::XRow> xRow ( xResult, uno::UNO_QUERY );  
    Reference<sdbc::XColumnLocate> xColumn ( xResult, uno::UNO_QUERY );

    Reference<table::XCell> xCell;
    Reference<beans::XPropertySet> xCellProp; 

    while (xResult->next()) {
        xCell.set (xSheet->getCellByPosition(nCol++, nRow), uno::UNO_QUERY);
        xCellProp.set (xCell, uno::UNO_QUERY);
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                     uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                     uno::makeAny(sal_Int32(2)) );
        xCell->setFormula(OUString::number(nIterator++));

        xCell.set (xSheet->getCellByPosition(nCol++, nRow), uno::UNO_QUERY);
        xCellProp.set (xCell, uno::UNO_QUERY);
        xCell->setFormula(xRow->getString(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("NOMBRE")))));

        sheetNotasEstudiante(xConnection, xSheet, nGestion, strAsignatura, nPeriodo, 
                             xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("ESTUDIANTE_ID")))),
                             nCol, nRow);
        nRow++;
        nCol = 0;
    }
}



sal_Int32
EnsecProtocolHandler::sheetNotasHeader ( const Reference<sdbc::XConnection>& xConnection, 
                                         const Reference<sheet::XSpreadsheet>& xSheet,
                                         sal_Int32 nGestion,                                         
                                         const OUString& strAsignatura,                                         
                                         sal_Int32 nPeriodo,
                                         const OUString& strPlantilla)
{

    if (nGestion && strAsignatura.isEmpty() && nPeriodo )
    {
    }

    sal_Int32 nRow = 0;
    sal_Int32 nCol = 0;
    util::Color clrBlackColor = 0x00000000;
    table::BorderLine aBorderLine;
    table::TableBorder aTableBorder;

    aBorderLine.Color = clrBlackColor;
    aBorderLine.OuterLineWidth = 1;
    aBorderLine.InnerLineWidth = 0;
    aBorderLine.LineDistance = 0;

    aTableBorder.IsTopLineValid = aTableBorder.IsBottomLineValid = aTableBorder.IsLeftLineValid = aTableBorder.IsRightLineValid = aTableBorder.IsDistanceValid = sal_True;
    aTableBorder.TopLine = aTableBorder.BottomLine = aTableBorder.LeftLine = aTableBorder.RightLine = aBorderLine;
    aTableBorder.Distance = 0;

    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM COLUMNA WHERE PLANTILLA ='")) + strPlantilla + OUString(RTL_CONSTASCII_USTRINGPARAM("' ORDER BY CODIGO"));

    Reference<sdbc::XPreparedStatement> xPreparedStatement ( xConnection->prepareStatement(strSQL), uno::UNO_QUERY );
    Reference<sdbc::XParameters> xParameters ( xPreparedStatement, uno::UNO_QUERY );

    //xParameters->setInt(1, nGestion);
    //xParameters->setString(2, strAsignatura);
    //xParameters->setInt(3, nPeriodo);

    Reference<sdbc::XResultSet> xResult = xPreparedStatement->executeQuery();
    Reference<sdbc::XRow> xRow ( xResult, uno::UNO_QUERY );  
    Reference<sdbc::XColumnLocate> xColumn ( xResult, uno::UNO_QUERY );
    Reference<table::XColumnRowRange> xColsRows( xSheet, uno::UNO_QUERY );
    Reference<table::XTableColumns> xTableColumns ( xColsRows->getColumns(), uno::UNO_QUERY );    
    Reference<table::XTableRows> xTableRows ( xColsRows->getRows(), uno::UNO_QUERY );    
    Reference<beans::XPropertySet> xColumnProp;
    Reference<beans::XPropertySet> xRowProp; 

    Reference<table::XCell> xCell;
    Reference<beans::XPropertySet> xCellProp; 
    Reference<container::XNameContainer> xAttributes;
    //xml::AttributeData rAttribute;

    xRowProp.set (xTableRows->getByIndex(nRow), uno::UNO_QUERY );
    xRowProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny( sal_Int32(5000)));

    while (xResult->next()) {
        xml::AttributeData rAttribute;
        xColumnProp.set (xTableColumns->getByIndex(nCol), uno::UNO_QUERY );
        xColumnProp->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), 
                                      uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("WIDTH"))))) );

        xCell.set (xSheet->getCellByPosition(nCol, nRow), uno::UNO_QUERY );
        xCellProp.set (xCell, uno::UNO_QUERY);
        xCellProp->getPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("UserDefinedAttributes"))) >>= xAttributes;
        rAttribute.Type = OUString(RTL_CONSTASCII_USTRINGPARAM("CDATA"));
        rAttribute.Value = xRow->getString(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("TIPO"))));
        xAttributes->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), uno::makeAny(rAttribute));
        xCellProp->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("UserDefinedAttributes")), uno::makeAny(xAttributes) );
        xCellProp->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("RotateAngle")), 
                                    uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("ANGULO"))))) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("IsTextWrapped")), 
                                     uno::makeAny( sal_True ) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                     uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("HORIZONTAL"))))) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                     uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("VERTICAL"))))) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                     uno::makeAny(aTableBorder) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("ParaTopMargin")), 
                                     uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("MARGEN_TOP"))))));
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("ParaBottomMargin")), 
                                     uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("MARGEN_BOTTOM"))))));
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharHeight")), 
                                     uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("FONT_HEIGHT"))))));
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     uno::makeAny(xRow->getInt(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("FONT_WEIGHT"))))));

        xCell->setFormula(xRow->getString(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("TITULO")))));

        xCell.set (xSheet->getCellByPosition(nCol, nRow + 1), uno::UNO_QUERY );
        xCellProp.set (xCell, uno::UNO_QUERY);
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                     uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                     uno::makeAny(sal_Int32(2)) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                     uno::makeAny(aTableBorder) );
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharHeight")), 
                                     uno::makeAny(sal_Int32(8)));
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     uno::makeAny(awt::FontWeight::BOLD));

        xCell->setFormula(xRow->getString(xColumn->findColumn(OUString(RTL_CONSTASCII_USTRINGPARAM("SUBTITULO")))));

        nCol++;
    }

    return nRow + 3;
}


void 
EnsecProtocolHandler::reporteNotas()
{
    // get Data Source ensec
    Reference<sdbc::XDataSource> xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference<sdb::XOfficeDatabaseDocument> xOfficeDbD  = getDatabaseDocument();

    if (!xOfficeDbD.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("El documento base de datos ensec no esta registrado!")));
        return;
    }

    Reference<sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );

    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }

    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    // get periodo data
    sal_Int32 nPeriodo = getPeriodo(xConnection);
    if (nPeriodo == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("El periodo no ha sido seleccionado!")));
        return;
    }

    Reference <sdb::XQueryDefinitionsSupplier> xQueryDS ( xDataSource, uno::UNO_QUERY_THROW);
    Reference<beans::XPropertySet> xPropQueryDefinition ( xQueryDS->getQueryDefinitions()->getByName(OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA"))), uno::UNO_QUERY_THROW );
        
    
    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.NOMBRE, CALC.NOTA FROM ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) AS NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          OUString::number(nGestion) + 
          OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +
          strAsignatura + 
          OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = ")) + 
          OUString::number(nPeriodo) +
          OUString(RTL_CONSTASCII_USTRINGPARAM("GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) CALC INNER JOIN ESTUDIANTE ON CALC.ESTUDIANTE = ESTUDIANTE.ESTUDIANTE_ID ORDER BY CALC.NOTA DESC"));

    xPropQueryDefinition->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Command")), uno::makeAny( strSQL ) );

    Reference<sdb::XReportDocumentsSupplier> xReportDS ( xOfficeDbD, uno::UNO_QUERY_THROW );
    Reference<container::XNameAccess> xContainer ( xReportDS->getReportDocuments(), uno::UNO_QUERY_THROW );
    Reference<frame::XComponentLoader> xLoader ( xReportDS->getReportDocuments(), uno::UNO_QUERY_THROW );
    
    if ( xContainer->hasElements() && xContainer->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA"))) ) {
      Sequence <beans::PropertyValue> aArgs(1);      
      aArgs[0].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("ActiveConnection"));
      aArgs[0].Value <<= xConnection;
      /*aArgs[1].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("CommandType"));
      aArgs[1].Value <<= uno::makeAny(2);
      aArgs[2].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("Command"));
      aArgs[2].Value <<= strSQL;*/


      xLoader->loadComponentFromURL( OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA")), 
				     OUString(RTL_CONSTASCII_USTRINGPARAM("_blank")),
				     0,
				     aArgs );
    }
}

void 
EnsecProtocolHandler::reporteAnualNotas()
{
    // get Data Source ensec
    Reference<sdbc::XDataSource> xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference<sdb::XOfficeDatabaseDocument> xOfficeDbD  = getDatabaseDocument();

    if (!xOfficeDbD.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("El documento base de datos ensec no esta registrado!")));
        return;
    }

    Reference<sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }

    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    Reference <sdb::XQueryDefinitionsSupplier> xQueryDS ( xDataSource, uno::UNO_QUERY_THROW);
    Reference<beans::XPropertySet> xPropQueryDefinition ( xQueryDS->getQueryDefinitions()->getByName(OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA_FINAL"))), uno::UNO_QUERY_THROW );
        
    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ESTUDIANTE.NOMBRE, COALESCE ( PRIMERO.NOTA, 0 ) PRIMERO, COALESCE ( SEGUNDO.NOTA, 0 ) SEGUNDO, COALESCE ( TERCERO.NOTA, 0 ) TERCERO, ( ( COALESCE ( PRIMERO.NOTA, 0 ) + COALESCE ( SEGUNDO.NOTA, 0 ) + COALESCE ( TERCERO.NOTA, 0 ) ) / 3 ) \"NOTA FINAL\" FROM ESTUDIANTE LEFT OUTER JOIN ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          OUString::number(nGestion) + 
          OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +
          strAsignatura + 
          OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = 1 GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) PRIMERO ON ESTUDIANTE.ESTUDIANTE_ID = PRIMERO.ESTUDIANTE LEFT OUTER JOIN ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          OUString::number(nGestion) +           
          OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +
          strAsignatura +
          OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = 2 GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) SEGUNDO ON PRIMERO.ESTUDIANTE = SEGUNDO.ESTUDIANTE LEFT OUTER JOIN ( SELECT NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE, SUM( NOTA.NOTA * ( EVALUACION.PONDERACION / EVALUACION.NOTA ) ) NOTA FROM NOTA INNER JOIN EVALUACION ON NOTA.EVALUACION = EVALUACION.ID AND NOTA.GESTION = ")) +
          OUString::number(nGestion) +
          OUString(RTL_CONSTASCII_USTRINGPARAM(" AND NOTA.ASIGNATURA = '")) +          
          strAsignatura +
          OUString(RTL_CONSTASCII_USTRINGPARAM("' AND NOTA.PERIODO = 3 GROUP BY NOTA.GESTION, NOTA.ASIGNATURA, NOTA.PERIODO, NOTA.ESTUDIANTE ) TERCERO ON PRIMERO.ESTUDIANTE = TERCERO.ESTUDIANTE WHERE PRIMERO.GESTION = ")) + 
          OUString::number(nGestion);


    xPropQueryDefinition->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Command")), uno::makeAny( strSQL ) );

    Reference<sdb::XReportDocumentsSupplier> xReportDS ( xOfficeDbD, uno::UNO_QUERY_THROW );
    Reference<container::XNameAccess> xContainer ( xReportDS->getReportDocuments(), uno::UNO_QUERY_THROW );
    Reference<frame::XComponentLoader> xLoader ( xReportDS->getReportDocuments(), uno::UNO_QUERY_THROW );
    
    if ( xContainer->hasElements() && xContainer->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA"))) ) {
      Sequence <beans::PropertyValue> aArgs(1);      
      aArgs[0].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("ActiveConnection"));
      aArgs[0].Value <<= xConnection;
      /*aArgs[1].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("CommandType"));
      aArgs[1].Value <<= uno::makeAny(2);
      aArgs[2].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("Command"));
      aArgs[2].Value <<= strSQL;*/

    
      xLoader->loadComponentFromURL( OUString(RTL_CONSTASCII_USTRINGPARAM("CALCULO_NOTA_FINAL")), 
				     OUString(RTL_CONSTASCII_USTRINGPARAM("_blank")),
				     0,
				     aArgs );
                     }
}



void
EnsecProtocolHandler::ingresarNotas()
{
    // get Data Source ensec
    Reference<sdbc::XDataSource> xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference<sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }


    // get gestion data
    sal_Int32 nGestion = getGestion(xConnection);
    if (nGestion == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La gestion no ha sido seleccionado!")));
        return;
    }

    // get Asignatura
    OUString strAsignatura = getAsignatura(xConnection, nGestion);
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    // get periodo data
    sal_Int32 nPeriodo = getPeriodo(xConnection);
    if (nPeriodo == -1) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("El periodo no ha sido seleccionado!")));
        return;
    }

    formularioNotas( xConnection, nGestion, nPeriodo, strAsignatura );

    //showMessageBox (OUString(RTL_CONSTASCII_USTRINGPARAM("Notas Ingresado! ")) + strAsignatura); // OUString::valueOf(nGestion));
}

void 
EnsecProtocolHandler::generarCronogramaTrabajo()
{
    // get Data Source ensec
    Reference<sdbc::XDataSource> xDataSource = getDataSource();

    if (!xDataSource.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La base de datos ensec no esta registrado!")));
        return;
    }

    Reference<sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );


    // connect database
    if (!xConnection.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede connectar a al base de datos ensec!")));
        return;
    }


    // get Document Sheet 
    Reference<sheet::XSpreadsheetDocument> xDocSheet = getDocumentSheet();
    if (!xDocSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Documento Sheet Actual!")));
        return;
    }

    // get Current Sheet 
    Reference<sheet::XSpreadsheet> xSheet = getCurrentSheet();

    if (!xSheet.is()) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se pudo abrir el Sheet Actual!")));
        return;
    }

    // get Asignatura
    OUString strAsignatura = getAsignatura( xConnection, 2014 );
    if ( strAsignatura.getLength() == 0) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La Asignatura no ha sido seleccionada!")));
        return;
    }

    sal_Int32 nCalendarRow = generarEncabezado( xDataSource, xSheet, strAsignatura );

    generarCalendario( xDataSource, xDocSheet, xSheet, strAsignatura, nCalendarRow );


    showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                    OUString(RTL_CONSTASCII_USTRINGPARAM("Cronograma de Trabajo generado!")));
}


void
EnsecProtocolHandler::generarCalendario( const Reference<sdbc::XDataSource>& xDataSource,
                                         const Reference<sheet::XSpreadsheetDocument>& xDocSheet,
                                         const Reference<sheet::XSpreadsheet>& xSheet,
                                         const OUString & strAsignatura,
                                         sal_Int32 nCalendarRow ) 
{
    // 12 meses.
    for ( sal_Int32 nMonth = 1; nMonth <= 12; nMonth++ ) {
        nCalendarRow = generarMes( xDataSource, xDocSheet, xSheet, strAsignatura, nCalendarRow, nMonth );
    }
}

sal_Int32 
EnsecProtocolHandler::generarMes( const Reference<sdbc::XDataSource>& xDataSource,
                                  const Reference<sheet::XSpreadsheetDocument>& xDocSheet,                                  
                                  const Reference<sheet::XSpreadsheet>& xSheet,
                                  const OUString & strAsignatura,
                                  sal_Int32 nCalendarRow,
                                  sal_Int32 nMonth )
{
    char strBuffer[700];
    OString strParam = OUStringToOString(strAsignatura, RTL_TEXTENCODING_UTF8);
    sprintf(strBuffer, "SELECT STARTDATE, DESCRIPTION, CATEGORY FROM CALENDARIO WHERE MONTH(STARTDATE)=%d UNION SELECT STARTDATE, DESCRIPTION, CATEGORY FROM CRONOGRAMA WHERE MONTH(STARTDATE)=%d AND ASIGNATURA='%s'  ORDER BY STARTDATE, DESCRIPTION", nMonth, nMonth, strParam.getStr());

    Reference<sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );
    //OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM CALENDARIO"));
    Reference<sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference<sdbc::XResultSet> xResultSet ( xStatement->executeQuery(OUString::createFromAscii(strBuffer)), 
                                                                 uno::UNO_QUERY_THROW );
    Reference<sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

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
EnsecProtocolHandler::mergeMonth( const Reference<sheet::XSpreadsheetDocument>& xDocSheet, 
                                  const Reference<sheet::XSpreadsheet>& xSheet, 
                                  sal_Int32 nMonth,
                                  sal_Int32 nStartRow,
                                  sal_Int32 nEndRow
                                  )
{
    xDocSheet.is();
    util::Color clrBlackColor = 0x00000000;
    table::BorderLine aBorderLine;
    table::TableBorder aTableBorder;

    aBorderLine.Color = clrBlackColor;
    aBorderLine.OuterLineWidth = 1;
    aBorderLine.InnerLineWidth = 0;
    aBorderLine.LineDistance = 0;

    aTableBorder.IsTopLineValid = aTableBorder.IsBottomLineValid = aTableBorder.IsLeftLineValid = aTableBorder.IsRightLineValid = aTableBorder.IsDistanceValid = sal_True;
    aTableBorder.TopLine = aTableBorder.BottomLine = aTableBorder.LeftLine = aTableBorder.RightLine = aBorderLine;
    aTableBorder.Distance = 0;

    if ( nStartRow == -1 || nEndRow == -1 )
        return;

    Reference <table::XCellRange> xRange ( xSheet->getCellRangeByPosition( 0, nStartRow, 0, nEndRow ), uno::UNO_QUERY);
    Reference <util::XMergeable> xMerge ( xRange, uno::UNO_QUERY);
    Reference<table::XCell> xCell (xRange->getCellByPosition(0,0), uno::UNO_QUERY );
    Reference<beans::XPropertySet> xCellProp ( xCell, uno::UNO_QUERY );

    OUString strMes;

    switch( nMonth ) {
    case 1 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("ENERO"));
        break;
    case 2 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("FEBRERO"));
        break;
    case 3 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("MARZO"));
        break;
    case 4 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("ABRIL"));
        break;
    case 5 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("MAYO"));
        break;
    case 6 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("JUNIO"));
        break;
    case 7 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("JULIO"));
        break;
    case 8 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("AGOSTO"));
        break;
    case 9 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("SEPTIEMBRE"));
        break;
    case 10 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("OCTUBRE"));
        break;
    case 11 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("NOVIEMBRE"));
        break;
    case 12 :
        strMes = OUString(RTL_CONSTASCII_USTRINGPARAM("DICIEMBRE"));
        break;
    }

    xCell->setFormula( strMes );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                 uno::makeAny(sal_Int16(2)) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 uno::makeAny(aTableBorder) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("RotateAngle")), 
                                 uno::makeAny(sal_Int16(9000)) );

    xMerge->merge( sal_True );


}


void
EnsecProtocolHandler::generarHeaderRow ( const Reference<sheet::XSpreadsheet>& xSheet,
                                         sal_Int32 nRow )
{
    util::Color clrLightGreyColor = 0x00C0C0C0;
    util::Color clrBlackColor = 0x00000000;
    table::BorderLine aBorderLine;
    table::TableBorder aTableBorder;

    aBorderLine.Color = clrBlackColor;
    aBorderLine.OuterLineWidth = 1;
    aBorderLine.InnerLineWidth = 0;
    aBorderLine.LineDistance = 0;

    aTableBorder.IsTopLineValid = aTableBorder.IsBottomLineValid = aTableBorder.IsLeftLineValid = aTableBorder.IsRightLineValid = aTableBorder.IsDistanceValid = sal_True;
    aTableBorder.TopLine = aTableBorder.BottomLine = aTableBorder.LeftLine = aTableBorder.RightLine = aBorderLine;
    aTableBorder.Distance = 0;
    
    Reference<table::XCell> xCell (xSheet->getCellByPosition(0, nRow), uno::UNO_QUERY );
    Reference<beans::XPropertySet> xCellProp ( xCell, uno::UNO_QUERY );

    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("MES")) );
    xCellProp->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CellBackColor")), 
                                 uno::makeAny(clrLightGreyColor) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 uno::makeAny(aTableBorder) );

    xCell.set (xSheet->getCellByPosition(1, nRow), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("FECHA")) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CellBackColor")), 
                                 uno::makeAny(clrLightGreyColor) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 uno::makeAny(aTableBorder) );


    xCell.set (xSheet->getCellByPosition(2, nRow), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("ACTIVIDAD")) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_CENTER) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CellBackColor")), 
                                 uno::makeAny(clrLightGreyColor) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 uno::makeAny(aTableBorder) );
}

void
EnsecProtocolHandler::generarRow( const Reference<sheet::XSpreadsheetDocument>& xDocSheet, 
                                  const Reference<sheet::XSpreadsheet>& xSheet, 
                                  const Reference<sdbc::XRow>& xRow,
                                  sal_Int32 nRow )
{
    util::Color clrBlackColor = 0x00000000;
    table::BorderLine aBorderLine;
    table::TableBorder aTableBorder;
    table::TableBorderDistances aTableBorderDistances;
    util::Date aDate;
    lang::Locale olocale;
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

    OUString strCategory = xRow->getString(3);

    Reference <util::XNumberFormatsSupplier> xFormatsSupplier ( xDocSheet, uno::UNO_QUERY_THROW);
    Reference <util::XNumberFormats> xNumberFormats ( xFormatsSupplier->getNumberFormats(), uno::UNO_QUERY_THROW); 
    //Reference < util::XNumberFormatTypes > xFormatTypes ( xFormatsSupplier->getNumberFormats(), uno::UNO_QUERY_THROW); 

    Reference<table::XCell> xCell (xSheet->getCellByPosition(1, nRow), uno::UNO_QUERY );
    Reference<beans::XPropertySet> xCellProp ( xCell, uno::UNO_QUERY );

    //int nKeyFormat = xFormatTypes->getStandardFormat(util::NumberFormat::DATE, olocale);
    
    //Reference< beans::XPropertySet > xFormatProp =  xNumberFormats->getByKey( nKeyFormat );
    //xFormatProp->getPropertyValue( OUString( RTL_CONSTASCII_USTRINGPARAM( "Locale" ) ) ) >>= olocale;

    nNewIndex = xNumberFormats->queryKey( OUString(RTL_CONSTASCII_USTRINGPARAM("NN D")), olocale, false );
    if ( nNewIndex == -1 ) // format not defined
        nNewIndex = xNumberFormats->addNew( OUString(RTL_CONSTASCII_USTRINGPARAM("NN D")), olocale );


    char strBuffer[20];
    aDate = xRow->getDate(1);
    sprintf(strBuffer, "%d/%d/%d", aDate.Year, aDate.Month, aDate.Day);

    xCell->setFormula( OUString::createFromAscii(strBuffer) );
    if ( strCategory.getLength() != 0 && strCategory.compareToAscii("Normal") != 0 ) 
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     uno::makeAny(awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                 uno::makeAny(sal_Int16(2)) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_LEFT) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("ParaIndent")), 
                                 uno::makeAny(sal_Int16(400)) );
    xCellProp->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("NumberFormat")), uno::makeAny(nNewIndex));
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 uno::makeAny(aTableBorder) );
    /*xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorderDistances")), 
      uno::makeAny(aTableBorderDistances) );*/


    xCell.set (xSheet->getCellByPosition(2, nRow), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( xRow->getString(2) );
    if ( strCategory.getLength() != 0 && strCategory.compareToAscii("Normal") != 0 ) 
        xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                     uno::makeAny(awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("IsTextWrapped")), 
                                 uno::makeAny( sal_True ) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_LEFT) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("ParaIndent")), 
                                 uno::makeAny(sal_Int16(300)) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("VertJustify")), 
                                 uno::makeAny(sal_Int16(2)) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TableBorder")), 
                                 uno::makeAny(aTableBorder) );


}


sal_Int32
EnsecProtocolHandler::generarEncabezado( const Reference<sdbc::XDataSource>& xDataSource, 
                                         const Reference<sheet::XSpreadsheet>& xSheet, 
                                         OUString& strAsignatura )
{

    Reference<sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );
    Reference<sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT ASIGNATURA.CODIGO, CARRERA.CARRERA, ASIGNATURA.NOMBRE, DOCENTE.NOMBRE, ASIGNATURA.CURSO, ASIGNATURA.GESTION  FROM ASIGNATURA INNER JOIN CARRERA ON ASIGNATURA.NOMBRE='")) + 
        strAsignatura + 
        OUString(RTL_CONSTASCII_USTRINGPARAM("' AND ASIGNATURA.GESTION=2014 AND CARRERA.CODIGO = ASIGNATURA.CARRERA INNER JOIN DOCENTE ON ASIGNATURA.DOCENTE = DOCENTE.CODIGO"));

    OString strDebug = OUStringToOString(strSQL, RTL_TEXTENCODING_UTF8);
    Reference<sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference<sdbc::XRow> xAsignatura ( xResultSet, uno::UNO_QUERY_THROW );  

    if ( !xResultSet->next() )
        return 8;

    strAsignatura = xAsignatura->getString(1);

    Reference<table::XColumnRowRange> xColsRows( xSheet, uno::UNO_QUERY_THROW );
    Reference<table::XTableColumns> xTableColumns ( xColsRows->getColumns(), uno::UNO_QUERY_THROW );    
    Reference<beans::XPropertySet> xColumnProp (xTableColumns->getByIndex( 0 ), uno::UNO_QUERY_THROW );
	xColumnProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(2470)) );

    xColumnProp.set (xTableColumns->getByIndex( 1 ), uno::UNO_QUERY_THROW );
	xColumnProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(2470)) );
    
    xColumnProp.set (xTableColumns->getByIndex( 2 ), uno::UNO_QUERY_THROW );
	xColumnProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(13840)) );

    Reference <table::XCellRange> xHeader ( xSheet->getCellRangeByPosition( 0, 0, 2, 0 ), uno::UNO_QUERY);
    Reference <util::XMergeable> xMerge ( xHeader, uno::UNO_QUERY);
    Reference<table::XCell> xCell (xHeader->getCellByPosition(0,0), uno::UNO_QUERY );
    Reference<beans::XPropertySet> xCellProp ( xCell, uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(3) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_CENTER) );
    xMerge->merge( sal_True );

    xHeader.set ( xSheet->getCellRangeByPosition( 0, 1, 2, 1 ), uno::UNO_QUERY );
    xMerge.set ( xHeader, uno::UNO_QUERY );
    xCell.set ( xHeader->getCellByPosition(0,0), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("CRONOGRAMA DE TRABAJO")) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_CENTER) );
    xMerge->merge( sal_True );

    xHeader.set ( xSheet->getCellRangeByPosition( 0, 2, 2, 2 ), uno::UNO_QUERY );
    xMerge.set ( xHeader, uno::UNO_QUERY );
    xCell.set ( xHeader->getCellByPosition(0,0), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );
    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("GESTION ")) + xAsignatura->getString(6) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
                                 uno::makeAny(table::CellHoriJustify_CENTER) );
    xMerge->merge( sal_True );

    xCell.set ( xSheet->getCellByPosition( 0, 4 ), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("CARRERA :")) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );


    xCell.set ( xSheet->getCellByPosition( 1, 4 ), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(2) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );


    xCell.set ( xSheet->getCellByPosition( 0, 5 ), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("DOCENTE :")) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );

    xCell.set ( xSheet->getCellByPosition( 1, 5 ), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(4) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );


    xCell.set ( xSheet->getCellByPosition( 0, 6 ), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("CURSO :")) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );

    xCell.set ( xSheet->getCellByPosition( 1, 6 ), uno::UNO_QUERY );
    xCellProp.set ( xCell, uno::UNO_QUERY );

    xCell->setFormula( xAsignatura->getString(5) );
    xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
                                 uno::makeAny(awt::FontWeight::BOLD) );

    return 8;
}

void 
EnsecProtocolHandler::formularioNotas(const Reference<sdbc::XConnection>& xConnection, 
                                      sal_Int32 nGestion,
                                      sal_Int32 nPeriodo,
                                      OUString&  strAsignatura)
{
    OUString strTitle;
    Reference<lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
    Reference <uno::XInterface> xDialogModel (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), uno::UNO_QUERY_THROW );
    Reference <beans::XPropertySet> xPropDlg ( xDialogModel, uno::UNO_QUERY_THROW );
    Reference <container::XNameContainer> xNameContainer ( xDialogModel, uno::UNO_QUERY_THROW );

    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(250)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(250)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Formulario Notas"))) );

    Reference<lang::XMultiServiceFactory> xSMDialog ( xDialogModel, uno::UNO_QUERY_THROW);

    // Label 
    Reference<uno::XInterface> xTextModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropTextModel ( xTextModel , uno::UNO_QUERY_THROW );  

    strTitle = OUString::number(nGestion) + OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) +
        strAsignatura + OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) +
        OUString::number(nPeriodo);

    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(10)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(70)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(strTitle) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("txtAsignatura"))) );

    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("txtAsignatura")), uno::makeAny(xTextModel));

    // Edit Ok
    Reference<uno::XInterface> xEditModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlEditModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropEditModel ( xEditModel , uno::UNO_QUERY_THROW );  

    xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(25)) );
    xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(70)) );
    xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("editEstudiante"))) );

    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("editEstudiante")), uno::makeAny(xEditModel));

    // Label 
    xTextModel.set (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), uno::UNO_QUERY_THROW);
    xPropTextModel.set ( xTextModel , uno::UNO_QUERY_THROW );  

    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(45)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(150)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(strAsignatura) );
    xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante"))) );

    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("txtEstudiante")), uno::makeAny(xTextModel));


    // GroupBox OK
    Reference<uno::XInterface> xGroupBoxModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlGroupBoxModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropGroupBoxModel (xGroupBoxModel, uno::UNO_QUERY_THROW);  
								     
    xPropGroupBoxModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropGroupBoxModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(55)) );
    xPropGroupBoxModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(200)) );
    xPropGroupBoxModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(100)) );
    xPropGroupBoxModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), 
                                           uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("containerNotas"))) );
    xPropGroupBoxModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Notas")))  );

    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("grpNotas")), uno::makeAny(xGroupBoxModel));


    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM EVALUACION WHERE GESTION=")) + 
          OUString::number(nGestion) +
          OUString(RTL_CONSTASCII_USTRINGPARAM(" AND PERIODO=")) +
          OUString::number(nPeriodo) +
          OUString(RTL_CONSTASCII_USTRINGPARAM(" AND ASIGNATURA='")) +
          strAsignatura +
          OUString(RTL_CONSTASCII_USTRINGPARAM("'"));
    
    Reference<sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference<sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference<sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    sal_Int32 nYStart = 65;
    sal_Int32 nYPos = nYStart;
    sal_Int32 nXPos = 37;
    sal_Int32 nHeight = 10;
    sal_Int32 nWidth = 100;
    sal_Int32 nYOffset = 10;
    sal_Int32 nXOffset = 5;
    sal_Int32 nEditWidth = 20;
    sal_Int32 nCalcWidth = 50;
    
    OUString strLabel;
    
    while ( xResultSet->next() ) {
        xTextModel.set (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), uno::UNO_QUERY_THROW);
        xPropTextModel.set (xTextModel , uno::UNO_QUERY_THROW );  

        xEditModel.set (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlEditModel"))), uno::UNO_QUERY_THROW);
        xPropEditModel.set ( xEditModel , uno::UNO_QUERY_THROW );  

        strLabel = xRow->getString(4) + OUString(RTL_CONSTASCII_USTRINGPARAM(" ")) + OUString::number(xRow->getFloat(6)) + OUString(RTL_CONSTASCII_USTRINGPARAM("-")) + OUString::number(xRow->getFloat(7));            

        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(nXPos)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(nYPos)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(nWidth)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(nHeight)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(strLabel));
        xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), uno::makeAny(OUString::number(xRow->getInt(1))));

        xNameContainer->insertByName(xRow->getString(4) + OUString(RTL_CONSTASCII_USTRINGPARAM("_label")), uno::makeAny(xTextModel));

        xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny( nXPos + nWidth + nXOffset ) );
        xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny( nYPos ) );
        xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny( nEditWidth ) );
        xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny( nHeight ) );
        xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), uno::makeAny( sal_False ) );
        xPropEditModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), uno::makeAny(xRow->getString(4) + OUString(RTL_CONSTASCII_USTRINGPARAM("_nota")) ));

        xNameContainer->insertByName(xRow->getString(4), uno::makeAny(xEditModel));

        xTextModel.set (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), uno::UNO_QUERY_THROW);
        xPropTextModel.set (xTextModel , uno::UNO_QUERY_THROW );  

        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16( nXPos + nWidth + nEditWidth + nXOffset)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(nYPos)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(nCalcWidth)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(nHeight)));
        xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM(" = "))));
        xPropTextModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), uno::makeAny(OUString::number( xRow->getFloat(7) / xRow->getFloat(6) )));


        xNameContainer->insertByName(xRow->getString(4) + OUString(RTL_CONSTASCII_USTRINGPARAM("_nota")) , uno::makeAny(xTextModel));

        nYPos += nYOffset;

	}

    xTextModel.set (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), 
                    uno::UNO_QUERY_THROW);
    xPropTextModel.set (xTextModel , uno::UNO_QUERY_THROW );  

    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(nXPos)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(nYPos)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(nWidth)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(nHeight)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), 
                                     uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Nota Final"))));

    xNameContainer->insertByName( OUString(RTL_CONSTASCII_USTRINGPARAM("nota_final_texto")),
                                  uno::makeAny(xTextModel));

    xTextModel.set (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), 
                    uno::UNO_QUERY_THROW);
    xPropTextModel.set (xTextModel , uno::UNO_QUERY_THROW );  

    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(nXPos + nWidth + nXOffset)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(nYPos)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(nEditWidth)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(nHeight)));
    xPropTextModel->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), 
                                     uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("0"))));

    xNameContainer->insertByName( OUString(RTL_CONSTASCII_USTRINGPARAM("nota_final")),
                                  uno::makeAny(xTextModel));

    nYPos += nYOffset;

    xPropGroupBoxModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(nYPos + nHeight - nYStart) );    
    
    nYPos += nYOffset;

    // Button Update OK
    Reference<uno::XInterface> xButtonModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropButtonModel ( xButtonModel , uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(nYPos)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate"))) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("None"))) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Tag")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Update"))));
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Enabled")), uno::makeAny( sal_False ) );

    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate")), uno::makeAny(xButtonModel));


    // Button OK
    xButtonModel.set (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
    xPropButtonModel.set ( xButtonModel , uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35 + 50 + 5)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(nYPos)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), uno::makeAny(xButtonModel));


    Reference <uno::XInterface> xDialog (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XControl> xControl ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XControlModel> xControlModel ( xDialogModel, uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference <awt::XToolkit> xToolkit (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XWindow> xWindow ( xControl, uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference <awt::XControlContainer> xControlContainer ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XButton> xButton ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), uno::UNO_QUERY_THROW );
    Reference <awt::XButton> xUpdate ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("btnUpdate"))), uno::UNO_QUERY_THROW );
    Reference <awt::XTextComponent> xText ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("editEstudiante"))), uno::UNO_QUERY_THROW );
    Reference <awt::XTextComponent> xTextNota;

    Reference <awt::XDialog> xDlg ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XActionListener> xEnsureDelete( new ClickHandler( xDlg ) );
    Reference <awt::XTextListener> xNotaHandler( new NotaHandler( xDlg ) );
    Reference <awt::XTextListener> xTextHandler( new EditHandler( xConnection, xDlg, nGestion, nPeriodo, strAsignatura ) );
    Reference <awt::XActionListener> xUpdateHandler( new UpdateHandler( xConnection, xDlg, nGestion, nPeriodo, strAsignatura ) );

    xButton->addActionListener( xEnsureDelete );
    xUpdate->addActionListener( xUpdateHandler );
    xText->addTextListener( xTextHandler );

    xResultSet.set ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    xRow.set ( xResultSet, uno::UNO_QUERY_THROW );  

    while ( xResultSet->next() ) {
        xTextNota.set ( xControlContainer->getControl(xRow->getString(4)), uno::UNO_QUERY );        
        if ( xTextNota.is() )
            xTextNota->addTextListener( xNotaHandler );
    }

    xDlg->execute();

    Reference <lang::XComponent> xDispose ( xDialog, uno::UNO_QUERY_THROW );
    xDispose->dispose();
}



sal_Int32  
EnsecProtocolHandler::getGestion(const Reference<sdbc::XConnection>& xConnection)
{
    sal_Int32 nResult = 0;
    Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
    Reference <uno::XInterface> xDialogModel (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), uno::UNO_QUERY_THROW );
    Reference <beans::XPropertySet> xPropDlg ( xDialogModel, uno::UNO_QUERY_THROW );

    //xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(100)) );
    //xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(100)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Gestion"))) );

    Reference <lang::XMultiServiceFactory> xSMDialog ( xDialogModel, uno::UNO_QUERY_THROW);
    Reference <uno::XInterface> xComboModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropComboModel ( xComboModel , uno::UNO_QUERY_THROW );  
    Reference <awt::XItemList> xItemList ( xComboModel , uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), uno::makeAny(sal_True) );
    //xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("ReadOnly")), uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), uno::makeAny(sal_Int8(10)) );


    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM GESTION"));
    Reference <sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference <sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference <sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));			
        xItemList->setItemData( nIndex, uno::makeAny(xRow->getInt(1)));
        nIndex = xItemList->getItemCount();
	}

    Reference <uno::XInterface> xButtonModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropButtonModel ( xButtonModel , uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference <container::XNameContainer> xNameContainer ( xDialogModel, uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), uno::makeAny(xComboModel));
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), uno::makeAny(xButtonModel));

    Reference <uno::XInterface> xDialog (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XControl> xControl ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XControlModel> xControlModel ( xDialogModel, uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference <awt::XToolkit> xToolkit (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XWindow> xWindow ( xControl, uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference <awt::XControlContainer> xControlContainer ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XButton> xButton ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), uno::UNO_QUERY_THROW );
    Reference <awt::XListBox> xListBox ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference <awt::XDialog> xDlg ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XActionListener> xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //   sal_Int32 clipformat;
    //    aSysDataType >>= clipformat;
 
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= nResult;

    Reference <lang::XComponent> xDispose ( xDialog, uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return nResult;

}


sal_Int32  
EnsecProtocolHandler::getPeriodo(const Reference<sdbc::XConnection>& xConnection)
{
    sal_Int32 nResult = 0;
    Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
    Reference <uno::XInterface> xDialogModel (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), uno::UNO_QUERY_THROW );
    Reference <beans::XPropertySet> xPropDlg ( xDialogModel, uno::UNO_QUERY_THROW );

    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Periodo"))) );

    Reference <lang::XMultiServiceFactory> xSMDialog ( xDialogModel, uno::UNO_QUERY_THROW);
    Reference <uno::XInterface> xComboModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropComboModel ( xComboModel , uno::UNO_QUERY_THROW );  
    Reference <awt::XItemList> xItemList ( xComboModel , uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), uno::makeAny(sal_Int8(10)) );


    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM PERIODO"));
    Reference <sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference <sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference <sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));			
        xItemList->setItemData( nIndex, uno::makeAny(xRow->getInt(1)));
        nIndex = xItemList->getItemCount();
	}

    Reference <uno::XInterface> xButtonModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropButtonModel ( xButtonModel , uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference <container::XNameContainer> xNameContainer ( xDialogModel, uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), uno::makeAny(xComboModel));
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), uno::makeAny(xButtonModel));

    Reference <uno::XInterface> xDialog (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XControl> xControl ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XControlModel> xControlModel ( xDialogModel, uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference <awt::XToolkit> xToolkit (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XWindow> xWindow ( xControl, uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference <awt::XControlContainer> xControlContainer ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XButton> xButton ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), uno::UNO_QUERY_THROW );
    Reference <awt::XListBox> xListBox ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference <awt::XDialog> xDlg ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XActionListener> xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //   sal_Int32 clipformat;
    //    aSysDataType >>= clipformat;
 
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= nResult;

    Reference <lang::XComponent> xDispose ( xDialog, uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return nResult;
}


OUString 
EnsecProtocolHandler::getAsignatura(const Reference<sdbc::XConnection>& xConnection, sal_Int32 nGestion)
{
    OUString strResult;
    Reference<lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
    Reference <uno::XInterface> xDialogModel (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), uno::UNO_QUERY_THROW );
    Reference <beans::XPropertySet> xPropDlg ( xDialogModel, uno::UNO_QUERY_THROW );

    //xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(100)) );
    //xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(100)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Asignatura"))) );

    Reference <lang::XMultiServiceFactory> xSMDialog ( xDialogModel, uno::UNO_QUERY_THROW);
    Reference <uno::XInterface> xComboModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropComboModel ( xComboModel , uno::UNO_QUERY_THROW );  
    Reference <awt::XItemList> xItemList ( xComboModel , uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), uno::makeAny(sal_True) );
    //xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("ReadOnly")), uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), uno::makeAny(sal_Int8(10)) );

    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM ASIGNATURA WHERE GESTION=")) + OUString::number((sal_Int32) nGestion);
    Reference <sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference <sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference <sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));
        xItemList->setItemData( nIndex, uno::makeAny(xRow->getString(1)));			
        nIndex = xItemList->getItemCount();
	}

    Reference <uno::XInterface> xButtonModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropButtonModel ( xButtonModel , uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference <container::XNameContainer> xNameContainer ( xDialogModel, uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), uno::makeAny(xComboModel));
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), uno::makeAny(xButtonModel));

    Reference <uno::XInterface> xDialog (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XControl> xControl ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XControlModel> xControlModel ( xDialogModel, uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference <awt::XToolkit> xToolkit (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XWindow> xWindow ( xControl, uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference <awt::XControlContainer> xControlContainer ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XButton> xButton ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), uno::UNO_QUERY_THROW );
    Reference <awt::XListBox> xListBox ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference <awt::XDialog> xDlg ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XActionListener> xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //strResult = xListBox->getSelectedItem();
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= strResult;

    Reference <lang::XComponent> xDispose ( xDialog, uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return strResult;
}

OUString 
EnsecProtocolHandler::getPlantilla(const Reference <sdbc::XConnection>& xConnection)
{
    OUString strResult;
    Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
    Reference <uno::XInterface> xDialogModel (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), uno::UNO_QUERY_THROW );
    Reference <beans::XPropertySet> xPropDlg ( xDialogModel, uno::UNO_QUERY_THROW );

    //xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(100)) );
    //xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(100)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(150)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(60)) );
    xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Plantilla"))) );

    Reference <lang::XMultiServiceFactory> xSMDialog ( xDialogModel, uno::UNO_QUERY_THROW);
    Reference <uno::XInterface> xComboModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlListBoxModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropComboModel ( xComboModel , uno::UNO_QUERY_THROW );  
    Reference <awt::XItemList> xItemList ( xComboModel , uno::UNO_QUERY_THROW );  
    
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(35)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(10)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(80)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Dropdown")), uno::makeAny(sal_True) );
    //xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("ReadOnly")), uno::makeAny(sal_True) );
    xPropComboModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("LineCount")), uno::makeAny(sal_Int8(10)) );

    OUString strSQL = OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT PLANTILLA FROM COLUMNA GROUP BY PLANTILLA"));
    Reference <sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    Reference <sdbc::XResultSet> xResultSet ( xStatement->executeQuery(strSQL), uno::UNO_QUERY_THROW );
    Reference<sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    sal_Int32 nIndex = xItemList->getItemCount();
    while ( xResultSet->next() ) {
        xItemList->insertItemText( nIndex, xRow->getString(2));
        xItemList->setItemData( nIndex, uno::makeAny(xRow->getString(1)));			
        nIndex = xItemList->getItemCount();
	}

    Reference <uno::XInterface> xButtonModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
    Reference <beans::XPropertySet> xPropButtonModel ( xButtonModel , uno::UNO_QUERY_THROW );  
								     
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(30)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
    xPropButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("OK"))) );

    Reference <container::XNameContainer> xNameContainer ( xDialogModel, uno::UNO_QUERY_THROW );
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura")), uno::makeAny(xComboModel));
    xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK")), uno::makeAny(xButtonModel));

    Reference <uno::XInterface> xDialog (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XControl> xControl ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XControlModel> xControlModel ( xDialogModel, uno::UNO_QUERY_THROW );
    xControl->setModel(xControlModel);

    Reference <awt::XToolkit> xToolkit (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), uno::UNO_QUERY_THROW );

    Reference <awt::XWindow> xWindow ( xControl, uno::UNO_QUERY_THROW );
    xWindow->setVisible( sal_True );
    xControl->createPeer( xToolkit, 0 );

    Reference <awt::XControlContainer> xControlContainer ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XButton> xButton ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("btnOK"))), uno::UNO_QUERY_THROW );
    Reference <awt::XListBox> xListBox ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("cmbAsignatura"))), uno::UNO_QUERY_THROW );

    //xButton->setActionCommand(OUString(RTL_CONSTASCII_USTRINGPARAM("OK")))

    Reference <awt::XDialog> xDlg ( xDialog, uno::UNO_QUERY_THROW );
    Reference <awt::XActionListener> xEnsureDelete( new ClickHandler( xDlg ) );
    xButton->addActionListener( xEnsureDelete );

    xDlg->execute();

    //strResult = xListBox->getSelectedItem();
    xItemList->getItemData(xListBox->getSelectedItemPos()) >>= strResult;

    Reference <lang::XComponent> xDispose ( xDialog, uno::UNO_QUERY_THROW );
    xDispose->dispose();

    return strResult;
}


Reference <sheet::XSpreadsheetDocument>
EnsecProtocolHandler::getDocumentSheet()
{
    Reference <sheet::XSpreadsheetDocument> xDocSheet;
    Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
    Reference <frame::XDesktop> xDesktop ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), uno::UNO_QUERY_THROW);    
    Reference <frame::XFrame> xFrame ( xDesktop->getCurrentFrame(), uno::UNO_QUERY_THROW);
    Reference <frame::XController> xController ( xFrame->getController(), uno::UNO_QUERY_THROW);
    Reference <frame::XModel> xModel ( xController->getModel(), uno::UNO_QUERY_THROW);

    xDocSheet.set ( xModel, uno::UNO_QUERY_THROW );

    return xDocSheet;
}




Reference <sheet::XSpreadsheet>
EnsecProtocolHandler::getCurrentSheet()
{
    Reference <sheet::XSpreadsheet> xSheet;
    Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
    Reference <frame::XDesktop> xDesktop ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), uno::UNO_QUERY_THROW);    
    Reference <frame::XFrame> xFrame ( xDesktop->getCurrentFrame(), uno::UNO_QUERY_THROW);
    Reference <sheet::XSpreadsheetView > xSheetView ( xFrame->getController(), uno::UNO_QUERY);    


    if ( !xSheetView.is() )
        return xSheet;

    xSheet.set( xSheetView->getActiveSheet(), uno::UNO_QUERY );

    return xSheet;
}

void
EnsecProtocolHandler::createForm()
{
  Reference <lang::XMultiComponentFactory> xServiceManager (mxContext->getServiceManager(), uno::UNO_QUERY_THROW);
  Reference <frame::XDesktop> xDesktop ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), uno::UNO_QUERY_THROW);
  Reference <frame::XFrame> xFrame ( xDesktop->getCurrentFrame(), uno::UNO_QUERY_THROW);
  Reference <frame::XController> xController ( xFrame->getController(), uno::UNO_QUERY_THROW);
  Reference <frame::XModel> xModel ( xController->getModel(), uno::UNO_QUERY_THROW);
  Reference <sheet::XSpreadsheetDocument> xSpreadsheetDocument ( xModel, uno::UNO_QUERY_THROW);
  Reference <sheet::XSpreadsheets> xSpreadsheets ( xSpreadsheetDocument->getSheets(), uno::UNO_QUERY_THROW);
  Reference <container::XIndexAccess> xIndexAccess ( xSpreadsheets, uno::UNO_QUERY_THROW);
  Reference <sheet::XSpreadsheet> xSheet ( xIndexAccess->getByIndex(0), uno::UNO_QUERY_THROW);
  Reference <container::XNamed> xNamed ( xSheet, uno::UNO_QUERY_THROW);
  //OUString strName = xNamed->getName();
  //Reference< table::XCell > xSheetCell (xSheet->getCellByPosition(0,0), uno::UNO_QUERY_THROW );
  //util::DateTime oDate, oDate0;

  //oDate.Year=2013;
  //oDate.Month=2;
  //oDate.Day=1;

  //xSheetCell->setValue( 0.0 );
  //xSheetCell->setFormula(OUString(RTL_CONSTASCII_USTRINGPARAM("2014/05/16")));

  //lang::Locale olocale;

  //Reference < util::XNumberFormatsSupplier > xFormatsSupplier ( xSpreadsheetDocument, uno::UNO_QUERY_THROW);
  //Reference < util::XNumberFormatTypes > xFormatTypes ( xFormatsSupplier->getNumberFormats(), uno::UNO_QUERY_THROW);
  //int nFormat = xFormatTypes->getStandardFormat(util::NumberFormat::DATE, olocale);
  //Reference < beans::XPropertySet > xPropSet ( xSheetCell, uno::UNO_QUERY_THROW);
  //xPropSet->setPropertyValue(OUString(RTL_CONSTASCII_USTRINGPARAM("NumberFormat")), uno::makeAny(nFormat) );


  Reference <uno::XInterface> xIDialog (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialogModel")), mxContext), uno::UNO_QUERY_THROW );
  Reference <beans::XPropertySet> xPropDlg ( xIDialog, uno::UNO_QUERY_THROW );

  xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(100)) );
  xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(100)) );
  xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(150)) );
  xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(100)) );
  xPropDlg->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Title")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Dialogo"))) );

  Reference <lang::XMultiServiceFactory> xSMDialog ( xIDialog, uno::UNO_QUERY_THROW);
  Reference <uno::XInterface> objButtonModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
  Reference <beans::XPropertySet> xButtonModel ( objButtonModel , uno::UNO_QUERY_THROW );  
								     
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(20)) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(70)) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Button1"))) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Click Me"))) );

  Reference <uno::XInterface> objLabelModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlFixedTextModel"))), uno::UNO_QUERY_THROW);
  Reference <beans::XPropertySet> xLabelModel ( objLabelModel, uno::UNO_QUERY_THROW );  

  xLabelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(40)) );
  xLabelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(30)) );
  xLabelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(100)) );
  xLabelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
  xLabelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Label1"))) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(1)) );
  xButtonModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Snake"))) );

  Reference <uno::XInterface> objCancelModel (xSMDialog->createInstance(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlButtonModel"))), uno::UNO_QUERY_THROW);
  Reference <beans::XPropertySet> xCancelModel ( objCancelModel, uno::UNO_QUERY_THROW );  
								     
  xCancelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionX")), uno::makeAny(sal_Int16(80)) );
  xCancelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("PositionY")), uno::makeAny(sal_Int16(70)) );
  xCancelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(50)) );
  xCancelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Height")), uno::makeAny(sal_Int16(14)) );
  xCancelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Name")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel"))) );
  xCancelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("TabIndex")), uno::makeAny(sal_Int8(0)) );
  xCancelModel->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Label")), uno::makeAny(OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel"))) );

  Reference <container::XNameContainer> xNameContainer ( xIDialog, uno::UNO_QUERY_THROW );
  xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("Button1")), uno::makeAny(objButtonModel));
  xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("Label1")), uno::makeAny(objLabelModel));
  xNameContainer->insertByName(OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel")), uno::makeAny(objCancelModel));

  Reference <uno::XInterface> xDialog (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.UnoControlDialog")), mxContext), uno::UNO_QUERY_THROW );
  
  Reference <awt::XControl> xControl ( xDialog, uno::UNO_QUERY_THROW );
  Reference <awt::XControlModel> xControlModel ( xIDialog, uno::UNO_QUERY_THROW );
  xControl->setModel(xControlModel);

  Reference <awt::XToolkit> xToolkit (xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.awt.Toolkit")), mxContext), uno::UNO_QUERY_THROW );

  Reference <awt::XWindow> xWindow ( xControl, uno::UNO_QUERY_THROW );
  xWindow->setVisible( sal_True );
  xControl->createPeer( xToolkit, 0 );

  Reference <awt::XControlContainer> xControlContainer ( xDialog, uno::UNO_QUERY_THROW );

  Reference <awt::XButton> xButton ( xControlContainer->getControl(OUString(RTL_CONSTASCII_USTRINGPARAM("Cancel"))), uno::UNO_QUERY_THROW );
  xButton->setActionCommand(OUString(RTL_CONSTASCII_USTRINGPARAM("OK")));

  Reference <awt::XDialog> xDlg ( xDialog, uno::UNO_QUERY_THROW );
  Reference<awt::XActionListener> xEnsureDelete( new ClickHandler( xDlg ) );
  xButton->addActionListener( xEnsureDelete );

  xDlg->execute();

  Reference <lang::XComponent > xDispose ( xDialog, uno::UNO_QUERY_THROW );
  xDispose->dispose();
}


void
EnsecProtocolHandler::showMessageBox( const Reference <awt::XToolkit>& rToolkit, const Reference <frame::XFrame>& rFrame, const OUString& aTitle, const OUString& aMsgText )
{
    if ( rFrame.is() && rToolkit.is() )
    {
        // describe window properties.
        awt::WindowDescriptor                aDescriptor;
        aDescriptor.Type              = awt::WindowClass_MODALTOP;
        aDescriptor.WindowServiceName = "infobox";
        aDescriptor.ParentIndex       = -1;
        aDescriptor.Parent            = Reference <awt::XWindowPeer>( rFrame->getContainerWindow(), UNO_QUERY );
        aDescriptor.Bounds            = awt::Rectangle(0,0,300,200);
        aDescriptor.WindowAttributes  = awt::WindowAttribute::BORDER |
            awt::WindowAttribute::MOVEABLE |
            awt::WindowAttribute::CLOSEABLE;

        Reference <awt::XWindowPeer> xPeer = rToolkit->createWindow( aDescriptor );
        if ( xPeer.is() )
        {
            Reference <awt::XMessageBox> xMsgBox( xPeer, UNO_QUERY );
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
Reference <sdbc::XDataSource>
EnsecProtocolHandler::getDataSource( ) throw ( uno::RuntimeException )
{
  Reference <sdbc::XDataSource> xDataSource;
  Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
  Reference <container::XNameAccess> xNameAccess ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), uno::UNO_QUERY_THROW );

  if ( xNameAccess->hasElements() && 
       xNameAccess->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("ensec")))) 
      xDataSource.set ( xNameAccess->getByName( OUString(RTL_CONSTASCII_USTRINGPARAM("ensec"))), uno::UNO_QUERY_THROW );

  return xDataSource;
}

Reference <sdb::XOfficeDatabaseDocument>
EnsecProtocolHandler::getDatabaseDocument( ) throw ( uno::RuntimeException )
{
  Reference <sdb::XOfficeDatabaseDocument> xDatabaseDocument;
  Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
  Reference <container::XNameAccess > xNameAccess ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), uno::UNO_QUERY_THROW );
  Reference <sdb::XDocumentDataSource> xDocumentDataSource;

  if ( xNameAccess->hasElements() && 
       xNameAccess->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("ensec"))))  {

      xDocumentDataSource.set ( xNameAccess->getByName( OUString(RTL_CONSTASCII_USTRINGPARAM("ensec"))), uno::UNO_QUERY_THROW );
      xDatabaseDocument.set ( xDocumentDataSource->getDatabaseDocument(), uno::UNO_QUERY_THROW );
  }
  
  return xDatabaseDocument;
}



// ---------------------------------------------------------------------------------
// helper to check flisol table definition
//
void
EnsecProtocolHandler::checkTable( Reference <sdbc::XDataSource> &aDataSource)
{
    aDataSource.is();
  bool bTableName = false;
  Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
  Reference <container::XNameAccess> xDatabaseContext ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), uno::UNO_QUERY_THROW );
  Reference <sdbc::XDataSource> xDataSource ( xDatabaseContext->getByName( OUString(RTL_CONSTASCII_USTRINGPARAM("Bibliography"))), uno::UNO_QUERY_THROW );
  Reference <sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );

  Reference <sdbcx::XTablesSupplier> xTSupplier( xConnection, UNO_QUERY);

  if ( xTSupplier.is() ){
    Reference <container::XNameAccess> xNameAccess ( xTSupplier->getTables(), uno::UNO_QUERY_THROW );    
    if (xNameAccess->hasElements() &&  xNameAccess->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("registro"))) ) {
      bTableName = true;
    }
  }

  if ( !bTableName ) {
    OUString sTable = OUString(RTL_CONSTASCII_USTRINGPARAM("CREATE TABLE registro ( ID INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY, NOMBRE VARCHAR, APELLIDO VARCHAR, PROFESION VARCHAR, UNIVERSIDAD VARCHAR, PUBLICIDAD VARCHAR )"));
    Reference <sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW );
    xStatement->executeQuery( sTable );
  }
}

// -----------------------------------------------------------------------------------------------
// helper to check flisol column definitions
//
void 
EnsecProtocolHandler::checkColumnHeader( Reference <sheet::XSpreadsheetDocument> & xDocument )
{
  Reference <sheet::XSpreadsheets> xSheets ( xDocument->getSheets(), uno::UNO_QUERY_THROW );
  Reference <container::XIndexAccess> xIndex ( xSheets, uno::UNO_QUERY_THROW );
  Reference <sheet::XSpreadsheet> xSheet ( xIndex->getByIndex(0), uno::UNO_QUERY_THROW );

  if ( xSheet.is() ) {
    Reference <sheet::XCellRangesQuery> xQuery ( xSheet, uno::UNO_QUERY_THROW );

    const sal_Int16 nContentFlags = sheet::CellFlags::STRING | 
      sheet::CellFlags::VALUE | 
      sheet::CellFlags::DATETIME | 
      sheet::CellFlags::FORMULA | 
      sheet::CellFlags::ANNOTATION;

    Reference <sheet::XSheetCellRanges> xSheetCellRanges ( xQuery->queryContentCells( nContentFlags ), uno::UNO_QUERY_THROW );    
    if ( xSheetCellRanges->getCount() > 0 ) {
      //Reference< sheet::XCellRange > xCellRange ( xSheetCellRanges->getByIndex(0), uno::UNO_QUERY_THROW );
      

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
  Reference <lang::XMultiComponentFactory> xServiceManager ( mxContext->getServiceManager(), uno::UNO_QUERY_THROW );
  Reference <uno::XInterface> xIDatabaseContext ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), uno::UNO_QUERY_THROW );
  Reference <container::XNameAccess> xNameAccess ( xIDatabaseContext, uno::UNO_QUERY_THROW );

  if (xNameAccess->hasElements() && xNameAccess->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014"))) ) {
    Reference <sdb::XDocumentDataSource> xDocumentDataSource (xNameAccess->getByName( OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014")) ), uno::UNO_QUERY_THROW);
    Reference <sdb::XOfficeDatabaseDocument> xOfficeDbD ( xDocumentDataSource->getDatabaseDocument(), uno::UNO_QUERY_THROW);
    Reference <sdb::XFormDocumentsSupplier> xFormDS ( xOfficeDbD, uno::UNO_QUERY_THROW );
    Reference <container::XNameAccess> xContainer ( xFormDS->getFormDocuments(), uno::UNO_QUERY_THROW );
    Reference <frame::XComponentLoader> xLoader ( xFormDS->getFormDocuments(), uno::UNO_QUERY_THROW );
    
    if ( xContainer->hasElements() && xContainer->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO"))) ) {
      //Reference< beans::XFastPropertySet > xFastProperty ( xContainer->getByName(OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO"))), uno::UNO_QUERY_THROW );
      //Reference< lang::XTypeProvider > xTypeProvider ( xContainer->getByName(OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO"))), uno::UNO_QUERY_THROW );      
      Reference <sdbc::XDataSource> xDataSource ( xOfficeDbD->getDataSource(), uno::UNO_QUERY_THROW );
      Reference <sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );

      Sequence <beans::PropertyValue> aArgs(2);      
      aArgs[0].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("ActiveConnection"));
      aArgs[0].Value <<= xConnection;
      aArgs[1].Name = OUString(RTL_CONSTASCII_USTRINGPARAM("OpenMode"));
      aArgs[1].Value <<= OUString(RTL_CONSTASCII_USTRINGPARAM("open"));;

      //form::NavigationBarMode mode = form::NavigationBarMode_CURRENT;

      //xFastProperty->setFastPropertyValue( 13, uno::makeAny( form::NavigationBarMode_CURRENT ) );
      //xFastProperty->setFastPropertyValue( 15, uno::makeAny( sal_True ) );
      //xFastProperty->setFastPropertyValue( 16, uno::makeAny( sal_True ) );
      //xFastProperty->setFastPropertyValue( 17, uno::makeAny( sal_True ) );
      
      
      /*uno::Sequence < uno::Type > seqTypes = xTypeProvider->getTypes();
      
            OString ouTypeName;
      for ( int nIterator = 0; nIterator < seqTypes.getLength(); nIterator++ ) {
	  ouTypeName = OUStringToOString( seqTypes[nIterator].getTypeName(), RTL_TEXTENCODING_ASCII_US );
	  }*/

      xLoader->loadComponentFromURL( OUString(RTL_CONSTASCII_USTRINGPARAM("REGISTRO")), 
				     OUString(RTL_CONSTASCII_USTRINGPARAM("_blank")),
				     0,
				     aArgs );

    }
    else showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                         OUString(RTL_CONSTASCII_USTRINGPARAM("La tabla REGISTRO no encontrado")));
  }
  else showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                       OUString(RTL_CONSTASCII_USTRINGPARAM("La Base de Datos Flisol2014 no encontrado")));
}

// -----------------------------------------------------------------------------
// open flisol sheet 
//
void 
EnsecProtocolHandler::openSheet()
{
  Reference <lang::XMultiComponentFactory> xServiceManager (mxContext->getServiceManager(), uno::UNO_QUERY_THROW);
  Reference <uno::XInterface> iDatabaseContext ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.sdb.DatabaseContext")), mxContext), uno::UNO_QUERY_THROW );
  Reference <container::XNameAccess> xNameAccess ( iDatabaseContext, uno::UNO_QUERY_THROW );

  if (xNameAccess->hasElements() && xNameAccess->hasByName(OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014"))) ) {

    Reference <sdbc::XDataSource>  xDataSource ( xNameAccess->getByName( OUString(RTL_CONSTASCII_USTRINGPARAM("flisol2014")) ), uno::UNO_QUERY_THROW);
    Reference <sdbc::XConnection> xConnection ( xDataSource->getConnection( OUString(RTL_CONSTASCII_USTRINGPARAM("")), OUString(RTL_CONSTASCII_USTRINGPARAM(""))), uno::UNO_QUERY_THROW );
    Reference <sdbc::XStatement> xStatement ( xConnection->createStatement(), uno::UNO_QUERY_THROW);
    Reference <sdbc::XResultSet> xResultSet ( xStatement->executeQuery( OUString(RTL_CONSTASCII_USTRINGPARAM("SELECT * FROM REGISTRO"))), uno::UNO_QUERY_THROW );

    Reference <frame::XDesktop> xDesktop ( xServiceManager->createInstanceWithContext(OUString(RTL_CONSTASCII_USTRINGPARAM("com.sun.star.frame.Desktop")), mxContext), uno::UNO_QUERY_THROW);    
    Reference <frame::XFrame> xFrame ( xDesktop->getCurrentFrame(), uno::UNO_QUERY_THROW);
    Reference <frame::XController> xController ( xFrame->getController(), uno::UNO_QUERY_THROW);    
    Reference <frame::XModel> xModel ( xController->getModel(), uno::UNO_QUERY_THROW);
    Reference <sheet::XSpreadsheetDocument> xSpreadsheetDocument ( xModel, uno::UNO_QUERY );

    if ( !xSpreadsheetDocument.is() ) {
        showMessageBox (mxToolkit, mxFrame, 
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("No se puede abrir un documento de Calc")));
        return;
    }

    Reference <sheet::XSpreadsheets> xSpreadsheets ( xSpreadsheetDocument->getSheets(), uno::UNO_QUERY_THROW);
    Reference <container::XIndexAccess> xIndexAccess ( xSpreadsheets, uno::UNO_QUERY_THROW);
    Reference <sheet::XSpreadsheet> xSheet ( xIndexAccess->getByIndex(0), uno::UNO_QUERY_THROW);
    Reference <table::XColumnRowRange> xColsRows( xSheet, uno::UNO_QUERY_THROW );
    Reference <sdbcx::XColumnsSupplier> xColumns ( xResultSet, uno::UNO_QUERY_THROW );
    Reference <container::XNameAccess> xColumnNames ( xColumns->getColumns(), uno::UNO_QUERY_THROW );    
    Reference <table::XTableColumns> xTableColumns ( xColsRows->getColumns(), uno::UNO_QUERY_THROW );    
    Sequence <rtl::OUString> arrayNames = xColumnNames->getElementNames();

    sal_Int32 nRowColumn = 0;
    sal_Int32 nRow = 1;

    for ( sal_Int32 nIterator = 0; nIterator < arrayNames.getLength(); nIterator++ ) {
	//Reference< sdb::XColumn > xColumn ( xColumnIndex->getByIndex( nIterator ), uno::UNO_QUERY_THROW );
      Reference <table::XCell> xCell ( xSheet->getCellByPosition( nIterator, nRowColumn ), uno::UNO_QUERY_THROW );
      //Reference< table::XTableColumn > xTableColumn ( xTableColumns->getByIndex( nIterator ), uno::UNO_QUERY_THROW );
      Reference <beans::XPropertySet> xColumnProp ( xTableColumns->getByIndex( nIterator ), uno::UNO_QUERY_THROW );
      Reference <beans::XPropertySet> xCellProp ( xCell, uno::UNO_QUERY_THROW );
      if ( arrayNames[nIterator].compareToAscii("ID") == 0) {
	xCell->setFormula( OUString(RTL_CONSTASCII_USTRINGPARAM("No.")) );
	xColumnProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(2000)) );
      }
      else {
	xCell->setFormula( arrayNames[nIterator] );	
	xColumnProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("Width")), uno::makeAny(sal_Int16(6000)) );
      }
      xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("CharWeight")), 
				   uno::makeAny(awt::FontWeight::BOLD) );
      xCellProp->setPropertyValue( OUString(RTL_CONSTASCII_USTRINGPARAM("HoriJustify")), 
				   uno::makeAny(table::CellHoriJustify_CENTER) );
    }

    Reference <sdbc::XRow> xRow ( xResultSet, uno::UNO_QUERY_THROW );  

    while ( xResultSet->next() ) {
      for ( sal_Int32 nIterator = 1; nIterator <= arrayNames.getLength(); nIterator++) {
	Reference <table::XCell> xCell ( xSheet->getCellByPosition( nIterator - 1, nRow ), uno::UNO_QUERY_THROW );	      
	Reference <beans::XPropertySet> xCellProp ( xCell, uno::UNO_QUERY_THROW );	      
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
                        OUString(RTL_CONSTASCII_USTRINGPARAM("Error!")),
                        OUString(RTL_CONSTASCII_USTRINGPARAM("La Base de Datos Flisol2014 no encontrado")));
}

// ------------------------------------------------------------------------------------------------
//
//
void SAL_CALL 
EnsecProtocolHandler::addStatusListener( const Reference< frame::XStatusListener >& xControl, 
                                         const util::URL& aURL )
  throw (RuntimeException)
{
    xControl.is();
    aURL.Complete.isEmpty();
}

// -------------------------------------------------------------------------------------------------
//
//
void SAL_CALL 
EnsecProtocolHandler::removeStatusListener( const Reference< frame::XStatusListener >& xControl, 
                                            const util::URL& aURL ) 
  throw (RuntimeException)
{
    xControl.is();
    aURL.Complete.isEmpty();
}
