<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0"
    xmlns="http://www.tei-c.org/ns/1.0"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:xhtml="http://www.w3.org/1999/xhtml"
    xpath-default-namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:tei-spreadsheet="https://github.com/oucs/tei-spreadsheet">

  <xsl:template match="c[@s]" mode="apply-style">
    <xsl:param name="styles"/>
    <xsl:param name="value"/>
    <xsl:variable name="s" select="@s"/>
    <xsl:variable name="cell-xf" select="$styles/cellXfs/xf[position()=$s]"/>
    <xsl:variable name="number-format" select="$styles/numFmts/numFmt[position()=$cell-xf/@applyNumberFormat]"/>
    <xsl:choose>
      <xsl:when test="number($value) = number($value) and matches($number-format/@formatCode, '[ymd]')">
        <!-- It looks like a date. -->
        <xsl:variable name="days" select="floor($value) - 2"/>
        <xsl:variable name="seconds" select="floor(($value - floor($value)) * 24 * 3600)"/>
        <xsl:variable name="duration" select="xs:dayTimeDuration(concat('P', $days, 'DT', $seconds, 'S'))"/>
        <xsl:variable name="date" select="xs:date('1900-01-01') + $duration"/>
        <xsl:value-of select="$date"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:copy-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="sml:c" mode="apply-style">
    <xsl:param name="value"/>
    <xsl:copy-of select="$value"/>
  </xsl:template>
</xsl:stylesheet>
