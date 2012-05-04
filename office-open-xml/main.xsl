<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0"
    xmlns="http://www.tei-c.org/ns/1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:sml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
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
    </TEI>
  </xsl:template>
</xsl:stylesheet>

