# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

$(eval $(call gb_Module_Module,external))

ifeq ($(OS)$(COM),WNTGCC)
$(eval $(call gb_Module_add_moduledir,external,mingw-externals))
endif

ifeq ($(COM),MSC)
$(eval $(call gb_Module_add_moduledir,external,msc-externals))
endif

$(eval $(call gb_Module_add_moduledirs,external,\
	$(call gb_Helper_optional,BOOST,boost) \
	$(call gb_Helper_optional,ICU,icu) \
	$(call gb_Helper_optional,LIBLANGTAG,liblangtag) \
))

# vim: set noet sw=4 ts=4:
