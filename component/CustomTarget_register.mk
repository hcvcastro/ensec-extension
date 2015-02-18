# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

$(eval $(call gb_CustomTarget_CustomTarget,ensec-register/ensec-register))

$(call gb_CustomTarget_get_target,ensec-register/ensec-register) : \
	$(call gb_Extension_get_target,ensec-extension)
	$(call gb_Output_announce,Installing extension $(notdir $(call gb_Extension_get_target,ensec-extension)),$(true),INSTALL,1)
	$(OO_UNOPKG) add -f $(call gb_Extension_get_target,ensec-extension)

# vim: set noet sw=4 ts=4:
