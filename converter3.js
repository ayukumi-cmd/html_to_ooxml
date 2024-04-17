// Require necessary modules for file system operations and path manipulation
const fs = require('fs');
const path = require('path');

// Function to convert HTML content to OOXML format
function convertToOOXML(htmlContent) {
    // Get HTML content from the specified element
    const html_Content = document.getElementById('htmlContent').innerHTML;
    // Convert HTML content to OOXML format
    const ooxml = convertHtmlToOOXML(html_Content);
    // Save the OOXML content to a file
    saveOOXMLFile(ooxml, 'output3.xml');
}

// Function to convert HTML content to OOXML format
function convertHtmlToOOXML(htmlContent) {
    // Split the HTML content into paragraphs
    const paragraphs = htmlContent.split(/\n*<\/?p[^>]*>/g).filter(p => p.trim().length > 0);
    // Initialize OOXML content
    let ooxml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:wordDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">';

    // Initialize IDs and other variables
    let paraId = 1;
    let textId = 1;
    let rsidR = "00B76DC2";
    let rsidRPr = "00B8076A";
    let rsidRDefault = "00000000";

    // Loop through each paragraph
    for (const paragraph of paragraphs) {
        // Generate IDs for the paragraph and text
        const paraIdStr = paraId.toString(16).toUpperCase().padStart(8, "0");
        const textIdStr = textId.toString(16).toUpperCase().padStart(8, "0");

        // Add paragraph to OOXML content
        ooxml += `<w:p w14:paraId="${paraIdStr}" w14:textId="${textIdStr}" w:rsidR="${rsidR}" w:rsidRDefault="${rsidRDefault}">`;
        paraId++;
        textId++;

        // Parse and add styles to OOXML content
        const styles = parseStyles(paragraph);
        if (styles) {
            ooxml += `<w:pPr>${styles}</w:pPr>`;
        }

        // Split paragraph into runs and add them to OOXML content
        const runs = paragraph.split(/<(?:\/?\w+)[^>]*>/g);
        let runId = 1;
        for (const run of runs) {
            if (run.trim().length === 0) continue;

            // Parse and add run styles to OOXML content
            const runStyles = parseRunStyles(run);
            ooxml += `<w:r${runStyles && runStyles.includes('<w:b/>') ? ` w:rsidRPr="${rsidRPr}"` : ''}>${runStyles ? `<w:rPr>${runStyles}</w:rPr>` : ''}`;

            // Add text to OOXML content
            const text = run.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ');
            if (text.trim().length > 0) {
                ooxml += `<w:t xml:space="preserve">${encodeXml(text)}</w:t>`;
            }

            // Close run tag
            ooxml += '</w:r>';
            runId++;
        }

        // Close paragraph tag
        ooxml += '</w:p>';
    }

    // Close OOXML document tag
    ooxml += '</w:wordDocument>';
    return ooxml;
}

// Function to save OOXML content to a file
function saveOOXMLFile(content, html_ooxml) {
    fs.writeFileSync(html_ooxml, content, 'utf8');
    console.log(`File '${html_ooxml}' saved successfully.`);
}

// Function to parse styles from paragraphs
function parseStyles(paragraph) {
    // Match and extract styles from paragraph
    const match = paragraph.match(/<p\s+style="([^"]+)"/i);
    if (!match) return null;

    // Split styles and process each one
    const styles = match[1].split(';').filter(s => s.trim().length > 0);
    let ooxml = '';

    for (const style of styles) {
        const [key, value] = style.split(':').map(s => s.trim());
        // Process each style and add to OOXML content
        switch (key) {
            case 'margin-left':
                ooxml += `<w:ind w:left="${parseLength(value)}"/>`;
                break;
            case 'margin-right':
                ooxml += `<w:ind w:right="${parseLength(value)}"/>`;
                break;
            case 'margin-top':
                ooxml += `<w:spacing w:before="${parseLength(value)}"/>`;
                break;
            case 'margin-bottom':
                ooxml += `<w:spacing w:after="${parseLength(value)}"/>`;
                break;
            case 'font-size':
                ooxml += `<w:sz w:val="${parseLength(value, 'pt')}"/>`;
                break;
            case 'font-family':
                ooxml += `<w:rFonts w:ascii="${value}" w:hAnsi="${value}"/>`;
                break;
            // Add more cases for other styles as needed
        }
    }

    return ooxml;
}

// Function to parse styles from runs
function parseRunStyles(run) {
    // Match and extract styles from run
    const styleMatch = run.match(/<\w+\s+style="([^"]+)"/i);
    const tagMatch = run.match(/<(\/?)\w+/g);

    let ooxml = '';
    let isOpeningTag = true;
    let currentTag = '';

    if (styleMatch) {
        // Split styles and process each one
        const styles = styleMatch[1].split(';').filter(s => s.trim().length > 0);
        for (const style of styles) {
            const [key, value] = style.split(':').map(s => s.trim());
            // Process each style and add to OOXML content
            switch (key) {
                case 'color':
                    ooxml += `<w:color w:val="${parseColor(value)}"/>`;
                    break;
                case 'background-color':
                    ooxml += `<w:shd w:fill="${parseColor(value)}"/>`;
                    break;
                case 'font-weight':
                    ooxml += value.toLowerCase() === 'bold' ? '<w:b/>' : '';
                    break;
                case 'font-style':
                    ooxml += value.toLowerCase() === 'italic' ? '<w:i/>' : '';
                    break;
                case 'text-decoration':
                    ooxml += value.toLowerCase() === 'underline' ? '<w:u w:val="single"/>' : '';
                    break;
                // Add more cases for other styles as needed
            }
        }
    }

    if (tagMatch) {
        // Process each tag and add to OOXML content
        for (const tag of tagMatch) {
            if (tag.startsWith('</')) {
                isOpeningTag = false;
                currentTag = tag.slice(2, -1);
            } else {
                isOpeningTag = true;
                currentTag = tag.slice(1);
            }

            switch (currentTag.toLowerCase()) {
                case 'strong':
                    ooxml += isOpeningTag ? '<w:b/>' : '</w:b>';
                    break;case 'u':
                    ooxml += isOpeningTag ? '<w:u w:val="single"/>' : '</w:u>';
                    break;
                // Add more cases for other tags as needed
            }
        }
    }

    return ooxml;
}

// Function to parse length values and convert them to the appropriate unit
function parseLength(value, unit = 'dxa') {
    // Match and extract numeric value and unit from the input
    const match = value.match(/^(\d+(\.\d+)?)(cm|pt|in|px)?$/);
    if (!match) return value;

    const [, numValue, , inputUnit] = match;
    const numericValue = parseFloat(numValue);

    // Convert the numeric value to the specified unit
    switch (inputUnit || 'px') {
        case 'cm':
            return Math.round(numericValue * 567);
        case 'pt':
            return Math.round(numericValue * 20);
        case 'in':
            return Math.round(numericValue * 1440);
        case 'px':
            return Math.round(numericValue * 20 / 3);
        default:
            return value;
    }
}

// Function to parse color values and convert them to hex format
function parseColor(value) {
    if (value.startsWith('rgb(')) {
        // Match and extract RGB values
        const [, r, g, b] = value.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
        // Convert RGB values to hex format
        return `${toHex(r)}${toHex(g)}${toHex(b)}`;
    }
    return value;
}

// Function to convert a single digit to hex format
function toHex(value) {
    const hex = parseInt(value).toString(16);
    return hex.length === 1 ? `0${hex}` : hex;
}

// Function to encode XML special characters
function encodeXml(text) {
    return text.replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// Path to the HTML file
const htmlFilePath = path.join(__dirname, 'index.html');
// Read the HTML file content
const htmlContent = fs.readFileSync(htmlFilePath, 'utf8');

// Extract the HTML content between the <body> tags
const bodyStartIndex = htmlContent.indexOf('<body>') + 7;
const bodyEndIndex = htmlContent.indexOf('</body>');
const htmlBodyContent = htmlContent.substring(bodyStartIndex, bodyEndIndex);

// Convert the extracted HTML content to OOXML format
convertToOOXML(htmlBodyContent);