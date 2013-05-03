<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0"
    xmlns="http://www.tei-c.org/ns/1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:tei-spreadsheet="https://github.com/oucs/tei-spreadsheet">
  <xsl:output method="xml" indent="yes"/>

  <xsl:param name="url"/>

  <!-- This is where we start finding out where things are in the document. -->
  <xsl:variable name="base-rels"
    select="document(concat($url, '/_rels/.rels'))/rels:Relationships"/>

  <!-- These are XML documents we expect to be able to find links to from
       within the rels document -->
  <xsl:variable name="office-document"
    select="document(concat($url, $base-rels/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument']/@Target))"/>
  <xsl:variable name="core-properties"
    select="document(concat($url, $base-rels/rels:Relationship[@Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties']/@Target))/cp:coreProperties"/>

  <xsl:function name="tei-spreadsheet:rels">
    <xsl:param name="node"/>
    <xsl:variable name="document-uri">
      <xsl:value-of select="document-uri($node/ancestor::document-node())"/>
    </xsl:variable>
    <xsl:variable name="rels-filename">
      <xsl:value-of select="replace($document-uri, '^(.*)/([^/]*)$', '$1/_rels/$2.rels')"/>
    </xsl:variable>
    <tei-spreadsheet:rels>
      <xsl:for-each select="document($rels-filename)/rels:Relationships/rels:Relationship">
        <tei-spreadsheet:rel
          id="{@Id}"
          type="{@Type}"
          target="{concat(replace($document-uri, '^(.*)/[^/]*$', '$1/'), @Target)}"/>
      </xsl:for-each>
    </tei-spreadsheet:rels>
  </xsl:function>

  <xsl:function name="tei-spreadsheet:parse-bstr">
    <!-- Section 22.4.2.4 of the Office Open XML Standard¹  defines a bstr type
         for reresenting Unicode characters that cannot be represented in XML
         1.0. Hence, a carriage return can be represented as "_x000d_". This
         function replaces such things with normal characters or decimal
         entities for all but the null character (_x0000_).

         ¹ http://www.ecma-international.org/publications/standards/Ecma-376.htm -->
    <xsl:param name="text"/>
    <xsl:variable name="pattern">^(.*?)(_x[\da-z]{4}_)(.*)$</xsl:variable>
    <xsl:choose>
      <xsl:when test="matches($text, $pattern, 'si')">
          <xsl:variable name="before" select="replace($text, $pattern, '$1', 'si')"/>
          <xsl:variable name="code" select="replace($text, $pattern, '$2', 'si')"/>
          <xsl:variable name="after" select="replace($text, $pattern, '$3', 'si')"/>

          <xsl:value-of select="$before"/>
          <xsl:value-of select="codepoints-to-string((tei-spreadsheet:hex-to-decimal(substring($code, 3, 4))))"/>
          <xsl:value-of select="tei-spreadsheet:parse-bstr($after)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$text"/>
        </xsl:otherwise>
      </xsl:choose>
  </xsl:function>


  <xsl:function name="tei-spreadsheet:hex-to-decimal">
    <xsl:param name="text"/>
    <xsl:variable name="codepoints" select="string-to-codepoints(upper-case($text))"/>
    <xsl:value-of select="tei-spreadsheet:hex-codepoints-to-decimal($codepoints)"/>
  </xsl:function>

  <xsl:function name="tei-spreadsheet:hex-codepoints-to-decimal">
    <xsl:param name="codepoints"/>
    <xsl:choose>
      <xsl:when test="empty($codepoints)">0</xsl:when>
      <xsl:otherwise>
        <xsl:variable name="high" select="subsequence($codepoints, 1, count($codepoints)-1)"/>
        <xsl:variable name="low" select="xs:integer($codepoints[count($codepoints)])"/>
        <xsl:variable name="digit">
          <xsl:choose>
            <xsl:when test="48 le $low and $low lt 58">
              <xsl:value-of select="$low - 48"/>
            </xsl:when>
            <xsl:when test="65 le $low and $low lt 71">
              <xsl:value-of select="$low - 55"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:message terminate="yes">Unexpected hex digit: <xsl:value-of select="$low"/></xsl:message>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="tei-spreadsheet:hex-codepoints-to-decimal($high) * 16 + $digit"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template name="main">
    <!-- Let's make sure this looks like a spreadsheet -->
    <xsl:if test="not($office-document/sml:workbook)">
      <xsl:message terminate="yes">This does not look like an Office OpenXML workbook.
The root element of this office document is a <xsl:value-of select="$office-document/*[1]/name()"/>, not a workbook.</xsl:message>
    </xsl:if>

    <TEI>
      <teiHeader>
        <fileDesc>
          <xsl:if test="$core-properties/dc:title or $core-properties/dc:creator">
            <titleStmt>
              <xsl:if test="$core-properties/dc:title">
                <title>
                  <xsl:value-of select="$core-properties/dc:title"/>
                </title>
              </xsl:if>
              <xsl:if test="$core-properties/dc:creator">
                <author>
                  <xsl:value-of select="$core-properties/dc:creator"/>
                </author>
              </xsl:if>
            </titleStmt>
          </xsl:if>
        </fileDesc>
      </teiHeader>
      <text>
        <body>
          <xsl:apply-templates select="$office-document/sml:workbook"/>
        </body>
      </text>
    </TEI>
  </xsl:template>

  <xsl:template match="sml:workbook">
    <xsl:apply-templates select="sml:sheets/sml:sheet"/>
  </xsl:template>

  <xsl:key name="strings" match="sml:si" use="count(preceding-sibling::*)"/>

  <xsl:template match="sml:sheet">
    <xsl:variable name="rels" select="tei-spreadsheet:rels(.)"/>
    <xsl:variable name="shared-strings" select="document($rels/*[@type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings']/@target)/sml:sst"/>


    <xsl:variable name="sheet-document" select="document($rels/*[@id=current()/@r:id]/@target)"/>
    <table>
      <head>
        <xsl:value-of select="@name"/>
      </head>
      <xsl:for-each select="$sheet-document/sml:worksheet/sml:sheetData/sml:row">
        <row n="{position()}">
          <xsl:for-each select="sml:c">
            <xsl:if test="preceding-sibling::sml:c">
              <xsl:call-template name="insert-omitted-cells">
                <xsl:with-param name="before" select="preceding-sibling::sml:c[1]"/>
                <xsl:with-param name="after" select="."/>
              </xsl:call-template>
            </xsl:if>
            <cell>
              <xsl:choose>
                <xsl:when test="@t='s'">
                  <xsl:value-of select="tei-spreadsheet:parse-bstr(key('strings', number(sml:v/text()), $shared-strings)/sml:t)"/>
                </xsl:when>
                <xsl:when test="@t='inlineStr'">
                  <xsl:value-of select="tei-spreadsheet:parse-bstr(sml:is/sml:t/text())"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="sml:v"/>
                </xsl:otherwise>
              </xsl:choose>
            </cell>
          </xsl:for-each>
        </row>
      </xsl:for-each>
    </table>
  </xsl:template>

  <xsl:template name="insert-omitted-cells">
    <xsl:param name="before"/> <!-- e.g. 'Y6' -->
    <xsl:param name="after"/><!-- e.g. 'AB6' -->

    <xsl:variable name="before-column-number" select="tei-spreadsheet:column-number($before/@r)"/>
    <xsl:variable name="after-column-number" select="tei-spreadsheet:column-number($after/@r)"/>

<!--
    <x ba="{$before}" bc="{$before-column-number}"  aa="{$after}" ac="{$after-column-number}"/>
-->

    <xsl:call-template name="empty-cells">
      <xsl:with-param name="count" select="$after-column-number - $before-column-number - 1"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:function name="tei-spreadsheet:column-number">
    <xsl:param name="cell-name"/>
    <xsl:value-of select="tei-spreadsheet:flatten-column-codepoints(reverse(string-to-codepoints(replace($cell-name, '\d+', ''))))"/>
  </xsl:function>

  <xsl:function name="tei-spreadsheet:flatten-column-codepoints">
    <xsl:param name="column-codepoints"/>
    <xsl:choose>
      <xsl:when test="not(count($column-codepoints))">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="26*tei-spreadsheet:flatten-column-codepoints(subsequence($column-codepoints, 2)) + $column-codepoints[1]-64"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template name="empty-cells">
    <xsl:param name="count"/>
    <xsl:if test="$count &gt; 0">
      <cell/>
      <xsl:call-template name="empty-cells">
        <xsl:with-param name="count" select="$count - 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

