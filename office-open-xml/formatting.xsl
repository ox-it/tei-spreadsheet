<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0"
    xmlns="http://www.tei-c.org/ns/1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:xhtml="http://www.w3.org/1999/xhtml"
    xpath-default-namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:tei-spreadsheet="https://github.com/oucs/tei-spreadsheet">

  <xsl:template match="si[not(r)]">
    <xsl:value-of select="t/text()"/>
  </xsl:template>

  <xsl:template match="si[r]">
    <xsl:variable name="xhtml">
      <xsl:apply-templates/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="count($xhtml/node()) = 1 and $xhtml/text()">
        <xsl:copy-of select="$xhtml"/>
      </xsl:when>
      <xsl:otherwise>
        <xhtml:div>
          <xsl:copy-of select="$xhtml"/>
        </xhtml:div>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="r">
    <xsl:variable name="with-br">
      <xsl:for-each select="tokenize(t/text(), '\r\n')">
        <xsl:if test="position() gt 1">
          <xhtml:br/>
        </xsl:if>
        <xsl:copy-of select="."/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="with-bold">
      <xsl:choose>
        <xsl:when test="rPr/b">
          <xhtml:b>
            <xsl:copy-of select="$with-br"/>
          </xhtml:b>
        </xsl:when>
        <xsl:otherwise>
          <xsl:copy-of select="$with-br"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="with-vertalign">
      <xsl:choose>
        <xsl:when test="rPr/vertAlign[@val='superscript']">
          <xsl:choose>
            <xsl:when test="$with-bold='2'">²</xsl:when>
            <xsl:when test="$with-bold='3'">³</xsl:when>
            <xsl:otherwise>
              <xhtml:sup>
                <xsl:copy-of select="$with-bold"/>
              </xhtml:sup>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="rPr/vertAlign[@val='subscript']">
          <xhtml:sub>
            <xsl:copy-of select="$with-bold"/>
          </xhtml:sub>
        </xsl:when>
        <xsl:otherwise>
          <xsl:copy-of select="$with-bold"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:copy-of select="$with-vertalign"/>
  </xsl:template>
</xsl:stylesheet>
