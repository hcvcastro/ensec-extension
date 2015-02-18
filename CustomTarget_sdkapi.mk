# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

$(eval $(call gb_CustomTarget_CustomTarget,sdkapi/sdkapi))

$(call gb_CustomTarget_get_target,sdkapi/sdkapi) : | \
	$(dir $(call gb_UnpackedTarball_get_target,%)).dir
	$(call gb_Output_announce,Generating SDK headers,$(true),SDK API HEADER,1)
	$(CPPUMAKER) -Gc -O$(dir $(call gb_CustomTarget_get_target))sdkapi $(OO_SDK_URE_TYPES) $(OO_SDK_OFFICE_TYPES)

# vim: set noet sw=4 ts=4:
