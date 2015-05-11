# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
include $(BUILDDIR)/config_ensec_extension.mk

$(eval $(call gb_Library_Library,ensec-component))

$(eval $(call gb_Library_use_externals,ensec-component,\
	boost_headers \
))

$(eval $(call gb_Library_use_custom_headers,ensec-component,sdkapi/sdkapi))

$(eval $(call gb_Library_set_include,ensec-component,\
    -I$(SRCDIR)/ensec/source/inc \
	-I$(dir $(call gb_CustomTarget_get_target,sdkapi/sdkapi))   \
	-I$(OO_SDK_INCLUDE) \
    $$(INCLUDE) \
))

$(eval $(call gb_Library_add_ldflags,ensec-component,\
	-L$(OO_SDK_HOME)/lib \
))

$(eval $(call gb_Library_add_libs,ensec-component,\
	-luno_cppu \
	-luno_cppuhelpergcc3 \
	-luno_sal \
))

$(eval $(call gb_Library_add_exception_objects,ensec-component,\
    ensec/source/ensec-component-export \
    ensec/source/ensec-protocol-handler \
    ensec/source/calendar \
))

$(eval $(call gb_Library_set_componentfile,ensec-component,ensec/util/ensec-component))

# vim: set noet sw=4 ts=4:
