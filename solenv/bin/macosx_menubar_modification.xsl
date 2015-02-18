<?xml version="1.0" encoding="utf-8"?>
<!--
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This file incorporates work covered by the following license notice:
 *
 *   Licensed to the Apache Software Foundation (ASF) under one or more
 *   contributor license agreements. See the NOTICE file distributed
 *   with this work for additional information regarding copyright
 *   ownership. The ASF licenses this file to you under the Apache
 *   License, Version 2.0 (the "License"); you may not use this file
 *   except in compliance with the License. You may obtain a copy of
 *   the License at http://www.apache.org/licenses/LICENSE-2.0 .
-->
<xsl:stylesheet version='1.0'
  xmlns:menu="http://openoffice.org/2001/menu"
  xmlns:xsl='http://www.w3.org/1999/XSL/Transform' >

  <!-- identity template, does reproduce every IN node on the output -->
  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- filtering template : removes the concerned nodes -->
  <!-- removes the separator just before the expected item -->
  <xsl:template match="menu:menuseparator[following-sibling::menu:menuitem[1]/@menu:id='.uno:Quit']"/>
  <!-- suppression of the Quit item -->
  <xsl:template match="menu:menuitem[@menu:id='.uno:Quit']"/>

  <xsl:template match="menu:menuseparator[following-sibling::menu:menuitem[1]/@menu:id='.uno:About']"/>
  <!-- suppression of the About item -->
  <xsl:template match="menu:menuitem[@menu:id='.uno:About']"/>

  <!-- suppression of the OptionsTreeDialog item -->
  <xsl:template match="menu:menuitem[@menu:id='.uno:OptionsTreeDialog']"/>

</xsl:stylesheet>
