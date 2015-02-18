# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

$(eval $(call gb_Package_Package,ensec-package,$(dir $(call gb_Extension_get_target,ensec-extension))))

$(eval $(call gb_Package_use_custom_target,ensec-package,sdkapi/sdkappi))

$(eval $(call gb_Package_set_outdir,ensec-package,$(INSTDIR)))

$(warning $(notdir $(call gb_Extension_get_target,ensec-extension)))
$(eval $(call gb_Package_add_files,ensec-package,component,\
	$(notdir $(call gb_Extension_get_target,ensec-extension)) \
))

# vim: set ts=4 sw=4 noet:
