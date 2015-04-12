# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
#

$(eval $(call gb_Extension_Extension,ensec-extension,component/source))

$(eval $(call gb_Extension_use_default_description,ensec-extension,component/source/description-en-es.txt))

#$(eval $(call gb_Extension_use_default_license,ensec-extension))

$(eval $(call gb_Extension_add_library,ensec-extension,ensec-component))

$(eval $(call gb_Extension_add_file,ensec-extension,ensec-component.rdb,$(call gb_Rdb_get_target,ensec-rdb)))

$(eval $(call gb_Extension_add_file,ensec-extension,ensec-component-addons.xcu,$(call gb_XcuFile_for_extension,component/xcu/ensec-component-addons.xcu)))

$(eval $(call gb_Extension_add_file,ensec-extension,ensec-component-protocol-handler.xcu,$(call gb_XcuFile_for_extension,component/xcu/ensec-component-protocol-handler.xcu)))

$(eval $(call gb_Extension_add_file,ensec-extension,inputstringdialog.ui,$(SRCDIR)/component/source/inputstringdialog.ui))

$(eval $(call gb_Extension_add_file,ensec-extension,dlgFormat.xdl,$(SRCDIR)/component/source/dlgFormat.xdl))

$(eval $(call gb_Extension_add_file,ensec-extension,bolivia.shell.png,$(SRCDIR)/component/img/bolivia.shell.png))

# vim: set noet sw=4 ts=4:
