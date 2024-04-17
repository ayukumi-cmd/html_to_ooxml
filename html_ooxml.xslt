<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

<xsl:template match="/">
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
            <xsl:apply-templates/>
        </w:body>
    </w:document>
</xsl:template>

<xsl:template match="p">
    <w:p>
        <xsl:apply-templates/>
    </w:p>
</xsl:template>

<xsl:template match="strong">
    <w:r>
        <w:rPr>
            <w:b/>
        </w:rPr>
        <w:t>
            <xsl:value-of select="."/>
        </w:t>
    </w:r>
</xsl:template>

<xsl:template match="em">
    <w:r>
        <w:rPr>
            <w:i/>
        </w:rPr>
        <w:t>
            <xsl:value-of select="."/>
        </w:t>
    </w:r>
</xsl:template>

<xsl:template match="span">
    <w:r>
        <w:rPr>
            <xsl:if test="@style='color: rgb(224, 62, 45);'">
                <w:color w:val="E03E2D"/>
            </xsl:if>
            <xsl:if test="@style='color: rgb(53, 152, 219);'">
                <w:color w:val="3598DB"/>
            </xsl:if>
            <xsl:if test="@style='color: blue; font-size: 40px;'">
                <w:color w:val="0000FF"/>
                <w:sz w:val="40"/>
            </xsl:if>
            <xsl:if test="@style='color: black; background: aqua;'">
                <w:color w:val="000000"/>
                <w:shd w:fill="00FFFF"/>
            </xsl:if>
            <xsl:if test="@style='color: white; background: purple;'">
                <w:color w:val="FFFFFF"/>
                <w:shd w:fill="800080"/>
            </xsl:if>
        </w:rPr>
        <w:t>
            <xsl:value-of select="."/>
        </w:t>
    </w:r>
</xsl:template>

<xsl:template match="text()">
    <w:t>
        <xsl:value-of select="."/>
    </w:t>
</xsl:template>

</xsl:stylesheet>
