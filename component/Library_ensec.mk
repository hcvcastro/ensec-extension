# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

$(eval $(call gb_Library_Library,ensec-component))

#$(eval $(call gb_Library_set_warnings_not_errors,ensec-component))

#$(eval $(call gb_Library_add_cxxflags,ensec-component,-DRTL_DISABLE_FAST_STRING))

#$(eval $(call gb_Library_use_externals,ensec-component,\
	boost_headers \
	mysqlcppconn \
))


$(eval $(call gb_Library_use_custom_headers,ensec-component,sdkapi/sdkapi))

$(eval $(call gb_Library_set_include,ensec-component,\
    -I$(SRCDIR)/component/source/inc \
	-I$(dir $(call gb_CustomTarget_get_target,sdkapi/sdkapi))   \
	-I$(OO_SDK_INCLUDE) \
    $$(INCLUDE) \
))

#$(eval $(call gb_Library_add_ldflags,ensec-component,\
    $(SNAKE_LIBS) \
))


#$(eval $(call gb_Library_use_libraries,ensec-component,\
	cppu \
	sal \
	salhelper \
	cppuhelper \
))

#$(eval $(call gb_Library_add_defs,ensec-component,\
	-DCPPDBC_EXPORTS \
	-DCPPCONN_LIB_BUILD \
	-DMARIADBC_VERSION_MAJOR=$(MARIADBC_MAJOR) \
	-DMARIADBC_VERSION_MINOR=$(MARIADBC_MINOR) \
	-DMARIADBC_VERSION_MICRO=$(MARIADBC_MICRO) \
	$(if $(filter NO,$(SYSTEM_MYSQL_CPPCONN)),\
	-DCPPCONN_LIB=\"$(call gb_Library_get_runtime_filename,mysqlcppconn)\") \
	$(if $(filter YES,$(BUNDLE_MARIADB)),\
	-DBUNDLE_MARIADB=\"$(LIBMARIADB)\") \
))


$(eval $(call gb_Library_add_exception_objects,ensec-component,\
    component/source/ensec-component-export \
    component/source/ensec-protocol-handler \
))


$(eval $(call gb_Library_set_componentfile,ensec-component,component/util/ensec-component))

# vim: set noet sw=4 ts=4:
