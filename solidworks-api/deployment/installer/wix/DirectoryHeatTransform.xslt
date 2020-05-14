<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:wix="http://schemas.microsoft.com/wix/2006/wi"
    xmlns="http://schemas.microsoft.com/wix/2006/wi"
    version="1.0"
    exclude-result-prefixes="xsl wix">

    <xsl:output method="xml" indent="yes" omit-xml-declaration="yes" />

    <xsl:strip-space elements="*" />

    <xsl:key
        name="AddInToRemove"
        match="wix:Component[ wix:File/@Source = '$(var.SourceOutDir)\CodeStack.Sw.MyToolbar.dll']"
        use="@Id" />

    <xsl:key
        name="TlbToRemove"
        match="wix:Component[ substring( wix:File/@Source, string-length( wix:File/@Source ) - 3 ) = '.tlb' ]"
        use="@Id" />

    <xsl:key
        name="XmlToRemove"
        match="wix:Component[ substring( wix:File/@Source, string-length( wix:File/@Source ) - 3 ) = '.xml' ]"
        use="@Id" />

    <xsl:key
        name="PdbToRemove"
        match="wix:Component[ substring( wix:File/@Source, string-length( wix:File/@Source ) - 3 ) = '.pdb' ]"
        use="@Id" />

    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()" />
        </xsl:copy>
    </xsl:template>

    <xsl:template match="*[ self::wix:Component or self::wix:ComponentRef ][ key( 'AddInToRemove', @Id ) ]" />
    <xsl:template match="*[ self::wix:Component or self::wix:ComponentRef ][ key( 'TlbToRemove', @Id ) ]" />
    <xsl:template match="*[ self::wix:Component or self::wix:ComponentRef ][ key( 'XmlToRemove', @Id ) ]" />
    <xsl:template match="*[ self::wix:Component or self::wix:ComponentRef ][ key( 'PdbToRemove', @Id ) ]" />
</xsl:stylesheet>