/* eslint-disable no-console */
exports.readRunProperties = readRunProperties;
exports.readParagraphProperties = readParagraphProperties;
exports.parseToNumber = parseToNumber;
exports.StyleReader = StyleReader;
exports.readTableProperties = readTableProperties;
exports.readTableCellBorders = readTableCellBorders;

// function isCharacter(element) {
//     return element.type === "character";
// }

// function isParagraph(element) {
//     return element.type === "paragraph";
// }
// readonly default?: IDefaultStylesOptions;
// readonly initialStyles?: BaseXmlComponent;
// readonly paragraphStyles?: readonly IParagraphStyleOptions[];
// readonly characterStyles?: readonly ICharacterStyleOptions[];
// readonly importedStyles?: readonly (XmlComponent | StyleForParagraph | StyleForCharacter | ImportedXmlComponent)[];

function parseToNumber(value) {
    return isNaN(Number(value)) ? null : Number(value);
}


function readUnderline(element) {
    if (element) {
        var value = element.attributes["w:val"];
        return value !== undefined && value !== "false" && value !== "0" && value !== "none";
    } else {
        return false;
    }
}

function readSize(element) {
    if (element) {
        var value = element.attributes["w:val"];
        var hasSize = value !== undefined && value !== "false" && value !== "0" && value !== "none";
        return hasSize ? parseToNumber(value) : null;
    } else {
        return null;
    }
}

function readBooleanElement(element) {
    if (element) {
        var value = element.attributes["w:val"];
        return value !== "false" && value !== "0";
    } else {
        return false;
    }
}

// 新增获取颜色方法
function readColor(element) {
    var value = element.attributes['w:fill'] || element.attributes['w:val'];
    if (!value || value === 'none') {
        return null;
    }
    return /^([0-9a-fA-F]{6}|[0-9a-fA-F]{3})$/.test(value) ? '#' + value : value;
}

// 新增获取字体属性
function readFont(element) {
    var font = {
        ascii: element.attributes['w:ascii'],
        cs: element.attributes['w:cs'],
        eastAsia: element.attributes['w:eastAsia'],
        hAnsi: element.attributes['w:hAnsi'],
        hint: element.attributes['w:hint']
    };

    return filterObject(font);
}

function filterObject(object) {
    for (var key in object) {
        if (Object.hasOwnProperty.call(object, key)) {
            var ele = object[key];
            if (!ele) {
                delete object[key];
            }
        }
    }
    if (Object.values(object).length) {
        return object;
    } else {
        return null;
    }
}

// 新增获取行间距和首行缩进方法
function readSpacing(element, type) {
    var spacing = {
        after: parseToNumber(element.attributes['w:after']),
        before: parseToNumber(element.attributes['w:before']),
        line: parseToNumber(element.attributes['w:line']),
        lineRule: element.attributes['w:lineRule'],
        beforeAutoSpacing: parseToNumber(element.attributes['w:beforeAutoSpacing']),
        afterAutoSpacing: parseToNumber(element.attributes['w:afterAutoSpacing'])
    };

    return filterObject(spacing);
}

function readCharacterSpacing(element) {
    return parseToNumber(element.attributes['w:val']);
}

function readShading(element) {
    var shading = {
        fill: element.attributes["w:fill"],
        color: element.attributes["w:color"],
        type: element.attributes["w:val"]
    };
    return filterObject(shading);
}

function createLanguageComponent(element) {
    var language = {
        value: element.attributes['w:val'],
        eastAsia: element.attributes['w:eastAsia'],
        bidirectional: element.attributes['w:bidi']
    };
    return filterObject(language);
}

function readTableCellBorderOptions(element) {
    if (element) {
        var border = {
            style: element.attributes["w:val"],
            /** Border color, in hex (eg 'FF00AA') */
            color: element.attributes["w:color"],
            /** Size of the border in 1/8 pt */
            size: parseToNumber(element.attributes["w:sz"]),
            /** Spacing offset. Values are specified in pt */
            space: parseToNumber(element.attributes["w:space"])
        };
        return filterObject(border);
    } else {
        return null;
    }
}

function readRunProperties(element) {
    var runProp = {
        'vertical-alignment': readSize(element.first("w:vertAlign")),
        bold: readBooleanElement(element.first("w:b")),
        'bold-complex-script': readBooleanElement(element.first("w:bCs")),
        italics: readBooleanElement(element.first("w:i")),
        'italics-complex-script': readBooleanElement(element.first("w:iCs")),
        underline: readUnderline(element.first("w:u")),
        effect: readSize(element.first("w:effect")),
        'emphasis-mark': '',
        color: readColor(element.firstOrEmpty('w:color')), // 添加文字颜色解析
        kern: readSize(element.first("w:kern")),
        position: '',
        size: readSize(element.first("w:sz")),
        'size-complex-script': readSize(element.first("w:szCs")),
        'small-caps': readBooleanElement(element.first("w:smallCaps")),
        'all-caps': readBooleanElement(element.first("w:caps")),
        strike: readBooleanElement(element.first("w:strike")),
        'double-strike': readBooleanElement(element.first("w:dstrike")),
        font: readFont(element.firstOrEmpty('w:rFonts')),
        highlight: readSize(element.first("w:highlightCs")),
        'highlight-complex-script': readSize(element.first("w:highlight")),
        'character-spacing': readCharacterSpacing(element.firstOrEmpty('w:spacing')), // 添加间距和缩进解析
        shading: readShading(element.firstOrEmpty("w:shd")),
        emboss: readBooleanElement(element.first("w:emboss")),
        imprint: readBooleanElement(element.first("w:imprint")),
        language: createLanguageComponent(element.firstOrEmpty("w:lang")),
        border: readTableCellBorderOptions(element.firstOrEmpty("w:bdr")),
        'snap-to-grid': readBooleanElement(element.first("w:snapToGrid")),
        vanish: readBooleanElement(element.first("w:vanish")),
        'spec-vanish': readBooleanElement(element.first("w:specVanish")),
        scale: readSize(element.first("w:w")),
        math: readBooleanElement(element.first("w:oMath")),
        bgColor: readColor(element.firstOrEmpty('w:highlight')) // 文本的背景色
    };
    return filterObject(runProp);
}

function readParagraphIndent(element) {
    var indent = {
        start: parseToNumber(element.attributes["w:start"]),
        end: parseToNumber(element.attributes["w:end"]),
        left: parseToNumber(element.attributes["w:left"]),
        right: parseToNumber(element.attributes["w:right"]),
        firstLine: parseToNumber(element.attributes["w:firstLine"]),
        hanging: parseToNumber(element.attributes["w:hanging"]),
        firstLineChars: parseToNumber(element.attributes["w:firstLineChars"])
    };
    return filterObject(indent);
}

function readTableCellBorders(element) {
    var borders = {
        top: readTableCellBorderOptions(element.first("w:top")),
        start: readTableCellBorderOptions(element.first("w:start")),
        left: readTableCellBorderOptions(element.first("w:left")),
        bottom: readTableCellBorderOptions(element.first("w:bottom")),
        end: readTableCellBorderOptions(element.first("w:end")),
        right: readTableCellBorderOptions(element.first("w:right"))
    };

    return filterObject(borders);
}

function readTableFloat(element, elements) {
    var float = {
        horizontalAnchor: element.attributes["w:horzAnchor"],
        verticalAnchor: element.attributes["w:vertAnchor"],
        absoluteHorizontalPosition: parseToNumber(element.attributes["w:tblpX"]),
        relativeHorizontalPosition: parseToNumber(element.attributes["w:tblpXSpec"]),
        absoluteVerticalPosition: parseToNumber(element.attributes["w:tblpY"]),
        relativeVerticalPosition: parseToNumber(element.attributes["w:tblpYSpec"]),
        topFromText: parseToNumber(element.attributes["w:bottomFromText"]),
        bottomFromText: parseToNumber(element.attributes["w:topFromText"]),
        rightFromText: parseToNumber(element.attributes["w:leftFromText"]),
        leftFromText: parseToNumber(element.attributes["w:rightFromText"]),
        overlap: elements.firstOrEmpty("w:overlap").attributes["w:val"]
    };
    return filterObject(float);
}

function readTableWidthElement(element) {
    var ele = {
        type: element.attributes["w:type"],
        size: parseToNumber(element.attributes["w:w"])
    };
    var result = filterObject(ele);
    return result && result['size'] ? result : null;
}

function readTableCellMargins(element) {
    var top = readTableCellMarginOptions(element.first("w:top"));
    var left = readTableCellMarginOptions(element.first("w:left"));
    var bottom = readTableCellMarginOptions(element.first("w:bottom"));
    var right = readTableCellMarginOptions(element.first("w:right"));
    var marginUnitType = (top || left || bottom || right);
    var margins = {
        marginUnitType: marginUnitType ? marginUnitType.type : null,
        top: top ? top.size : null,
        left: left ? left.size : null,
        bottom: bottom ? bottom.size : null,
        right: right ? right.size : null
    };

    return filterObject(margins);
}

function readTableCellMarginOptions(element) {
    if (element) {
        return {
            size: parseToNumber(element.attributes["w:w"]),
            type: element.attributes["w:type"]
        };
    } else {
        return null;
    }
}


function readParagraphProperties(element) {
    var paragraphProp = {
        heading: element.firstOrEmpty('w:pStyle').attributes['w:val'], // w:pStyle
        'keep-next': readBooleanElement(element.first("w:keepNext")),
        'keep-lines': readBooleanElement(element.first("w:keepLines")),
        'page-break-before': element.firstOrEmpty('w:pageBreakBefore').attributes['w:val'],
        'widow-control': readBooleanElement(element.first("w:widowControl")),
        alignment: element.firstOrEmpty("w:jc").attributes["w:val"],
        indent: readParagraphIndent(element.firstOrEmpty("w:ind")),
        bgColor: readColor(element.firstOrEmpty('w:shd')), // 添加背景色的解析
        spacing: readSpacing(element.firstOrEmpty('w:spacing'), 'w:p'), // 添加间距和缩进解析
        border: readTableCellBorders(element.firstOrEmpty('w:pBdr')),
        bidirectional: readBooleanElement(element.first("w:bidi")),
        shading: readShading(element.firstOrEmpty("w:shd")),
        'suppress-line-numbers': readBooleanElement(element.first("w:suppressLineNumbers")),
        'word-wrap': readBooleanElement(element.first("w:suppressLineNumbers")),
        scale: readSize(element.first("w:w"))
        // TODO: 未知属性
        // tabStops: '',
        // frame: '',
    };
    return filterObject(paragraphProp);
}

function readTableProperties(element, elements) {
    return {
        float: readTableFloat(element.firstOrEmpty("w:tblpPr"), element),
        width: readTableWidthElement(element.firstOrEmpty("w:tblW")),
        indent: readTableWidthElement(element.firstOrEmpty("w:tblInd")),
        layout: element.firstOrEmpty("w:tblLayout").attributes["w:type"],
        cellMargin: readTableCellMargins(element.firstOrEmpty("w:tblCellMar")),
        shading: readShading(element.firstOrEmpty("w:shd")),
        alignment: element.firstOrEmpty("w:jc").attributes['w:val'],
        visuallyRightToLeft: element.first("w:bidiVisual") ? true : false,
        borders: readTableCellBorders(element.firstOrEmpty("w:tblBorders"))
    };
}

function StyleReader(styleElement, type) {
    var base = {
        name: styleElement.firstOrEmpty("w:name").attributes['w:val'],
        basedOn: styleElement.firstOrEmpty("w:basedOn").attributes['w:val'],
        next: styleElement.firstOrEmpty("w:next").attributes['w:val'],
        link: styleElement.firstOrEmpty("w:link").attributes['w:val'],
        uiPriority: styleElement.firstOrEmpty("w:uiPriority").attributes['w:val'],
        semiHidden: styleElement.firstOrEmpty("w:semiHidden").attributes['w:val'],
        unhideWhenUsed: styleElement.firstOrEmpty("w:unhideWhenUsed").attributes['w:val'],
        quickFormat: styleElement.firstOrEmpty("w:quickFormat").attributes['w:val']
    };
    var xmlStyleReaders = {
        "paragraph": function(element) {
            var paragraphProperties = readParagraphProperties(element.firstOrEmpty("w:pPr"));
            var runProperties = readRunProperties(element.firstOrEmpty("w:rPr"));
            if (runProperties) {
                paragraphProperties = paragraphProperties || {};
                paragraphProperties['run-properties'] = runProperties;
            }
            var result = Object.assign({}, base, paragraphProperties);
            return filterObject(result);
        },
        "character": function(element) {
            var runProperties = readRunProperties(element.firstOrEmpty("w:rPr"));
            var result = Object.assign({}, base, runProperties);
            return filterObject(result);
        },
        "table": function(element) {
            var tableProperties = readTableProperties(element.firstOrEmpty("w:tblPr"));
            var result = Object.assign({}, base, tableProperties);
            return filterObject(result);
        },
        "numbering": function(element) {
            return {};
        }
        
    };
    var handler = xmlStyleReaders[type];
    if (handler) {
        return handler(styleElement);
    } else {
        return {};
    }
}
