/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*************************************************************************
 *
 * DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
 *
 * Copyright 2014, Bolivia@SHELL:$_ Henry Castro <hcvcastro@gmail.com>.
 *
 * OpenOffice.org - a multi-platform office productivity suite
 *
 * This file is part of ensec extension component.
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
#include <cppuhelper/bootstrap.hxx>
#include <cppuhelper/component.hxx>

#ifndef _COM_SUN_STAR_DEPLOYMENT_PACKAGEINFORMATIONPROVIDER_HPP_
#include <com/sun/star/deployment/PackageInformationProvider.hpp>
#endif

#ifndef _COM_SUN_STAR_UTIL_XURLTransformer_HPP_
#include <com/sun/star/util/XURLTransformer.hpp>
#endif

#ifndef _COM_SUN_STAR_UTIL_URLTransformer_HPP_
#include <com/sun/star/util/URLTransformer.hpp>
#endif

#ifndef _COM_SUN_STAR_IO_XINPUTSTREAMPROVIDER_HPP_
#include <com/sun/star/io/XInputStreamProvider.hpp>
#endif

#ifndef _COM_SUN_STAR_UCB_XSIMPLEFILEACCESS_HPP_
#include <com/sun/star/ucb/SimpleFileAccess.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XDIALOG_HPP_
#include "com/sun/star/awt/XDialog.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_XCONTROL_HPP_
#include "com/sun/star/awt/XControl.hpp"
#endif

#ifndef _COM_SUN_STAR_AWT_UNOCONTROLDIALOG_HPP_
#include <com/sun/star/awt/UnoControlDialog.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_POSSIZE_HPP_
#include <com/sun/star/awt/PosSize.hpp>
#endif

#ifndef _COM_SUN_STAR_UI_DIALOGS_XEXECUTABLEDIALOG_HPP_
#include <com/sun/star/ui/dialogs/XExecutableDialog.hpp>
#endif

#ifndef _COM_SUN_STAR_UI_DIALOGS_XFILEPICKER3_HPP_
#include <com/sun/star/ui/dialogs/XFilePicker3.hpp>
#endif

#ifndef _COM_SUN_STAR_UI_DIALOGS_TEMPLATEDESCRIPTION_HPP_
#include <com/sun/star/ui/dialogs/TemplateDescription.hpp>
#endif

#ifndef _COM_SUN_STAR_UI_DIALOGS_EXECUTABLEDIALOGRESULTS_HPP_
#include <com/sun/star/ui/dialogs/ExecutableDialogResults.hpp>
#endif

#ifndef _COM_SUN_STAR_AWT_XMESSAGEBOX_HPP_
#include <com/sun/star/awt/XMessageBox.hpp>
#endif

#include "ensec-protocol-handler.h"
#include "xmlscript/xmldlg_imexp.hxx"

namespace uno = com::sun::star::uno;
namespace awt = com::sun::star::awt;
namespace deployment = com::sun::star::deployment;
namespace util = com::sun::star::util;
namespace io = com::sun::star::io;
namespace container = com::sun::star::container;
namespace ucb = com::sun::star::ucb;
namespace frame = com::sun::star::frame;
namespace dialogs = com::sun::star::ui::dialogs;
namespace lang = com::sun::star::lang;

void EnsecProtocolHandler::UpdateCalendar()
{
    try
    {
        OUString sFile = PickCSVFile();

        Reference<ucb::XSimpleFileAccess3> xFile( ucb::SimpleFileAccess::create(mxContext) );
        Reference<io::XInputStream> xInput;
        if( xFile->exists( sFile ) )
        {
            xInput = xFile->openFileRead( sFile );
            

        }



        /*Reference <util::XURLTransformer> xTransformer (util::URLTransformer::create( mxContext ) );
        const Reference<deployment::XPackageInformationProvider> xPackageInfo = deployment::PackageInformationProvider::get(mxContext);

        OUString sRoot = xPackageInfo->getPackageLocation(OUString("bolivia@shell:Ensec"));

        util::URL aURL;
        //aURL.Complete = sRoot + OUString("");
        sRoot = sRoot + OUString("/dlgFormat.xdl");

        xTransformer->parseStrict( aURL );

        Reference<ucb::XSimpleFileAccess3> xSFI(ucb::SimpleFileAccess::create(mxContext) );
        // load a dialog model from the stream describing it
        //Reference<io::XInputStreamProvider> xISP( aURL, UNO_QUERY_THROW );
        //Reference<io::XInputStream> xInput( xISP->createInputStream(), UNO_QUERY_THROW );

        const Reference<frame::XModel> xModel;
        Reference<container::XNameContainer> xDialogModel( mxContext->getServiceManager()->createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", mxContext), uno::UNO_QUERY_THROW );
        Reference<awt::XControlModel> xControlModel ( xDialogModel, uno::UNO_QUERY_THROW );

        Reference<awt::XUnoControlDialog> xDialog(mxContext->getServiceManager()->createInstanceWithContext
("com.sun.star.awt.UnoControlDialog", mxContext), uno::UNO_QUERY_THROW );
        //Reference<awt::XControl> xControl ( xDialog, uno::UNO_QUERY_THROW );

        //Reference <awt::XToolkit> xToolkit (mxContext->getServiceManager()->createInstanceWithContext
("com.sun.star.awt.Toolkit", mxContext), uno::UNO_QUERY_THROW );

        Reference<awt::XWindow> xWindow( mxFrame->getContainerWindow(), uno::UNO_SET_THROW );
        awt::Rectangle aPosSize = xWindow->getPosSize();

        Reference<io::XInputStream> xInput;
        if( xSFI->exists( sRoot ) )
        {
            xInput = xSFI->openFileRead( sRoot );
            ::xmlscript::importDialogModel( xInput, xDialogModel, mxContext, xModel );
            xDialog->setModel(xControlModel);
            xDialog->setVisible(true);
            xDialog->createPeer(mxToolkit, mxToolkit->getDesktopWindow());  

            awt::Rectangle aRect = xDialog->getPosSize();

            Reference<awt::XWindow> xControlWindow( xDialog->getPeer(), uno::UNO_QUERY_THROW );

            xControlWindow->setPosSize( static_cast<sal_Int32>((aPosSize.Width - aRect.Width) / 2.0), 
                                        static_cast<sal_Int32>((aPosSize.Height - aRect.Height) / 2.0), 0, 0, awt::PosSize::POS );

            xDialog->execute();
        }*/
    }
    catch ( const Exception& aError )
    {
        ShowMessageBox ( awt::MessageBoxType_ERRORBOX,
                         awt::MessageBoxButtons::BUTTONS_OK,
                         "Excpetion",
                         aError.Message );
    }
}

OUString EnsecProtocolHandler::PickCSVFile()
{
    OUString sFile;
    Reference<dialogs::XFilePicker3> xFilePicker( CreateInstance(FILE_PICKER), uno::UNO_QUERY_THROW );
    Reference<dialogs::XFilterManager> xFilter( xFilePicker, uno::UNO_QUERY_THROW );
    Reference<lang::XInitialization> xInit( xFilePicker, uno::UNO_QUERY_THROW );

    Sequence < Any > aInitArguments( 1 );
    aInitArguments[0] <<= dialogs::TemplateDescription::FILEOPEN_SIMPLE;
    xInit->initialize( aInitArguments );

    xFilter->appendFilter( "(Icedove) Comma Separated Values", "*.csv" );
    xFilePicker->setTitle( "Open Icedov exported calendar to comma separated value (CSV)" );        

    if ( xFilePicker->execute() == dialogs::ExecutableDialogResults::OK )
    {
        Sequence<OUString> aFile = xFilePicker->getFiles();
        sFile = aFile[0];
        /*Reference<awt::XMessageBox> xMsgBox( CreateMessageBox(  awt::MessageBoxType_INFOBOX, 
                                                                awt::MessageBoxButtons::BUTTONS_OK,
                                                                "Information",
                                                                aFile[0] ), 
                                              css::uno::UNO_SET_THROW );
        xMsgBox->execute();*/
    }

    Reference<lang::XComponent> xFini( xFilePicker, uno::UNO_QUERY_THROW );
    xFini->dispose();

    return sFile;
}




