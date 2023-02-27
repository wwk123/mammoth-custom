/* eslint-disable no-console */
exports.createBodyReader = createBodyReader;
exports._readNumberingProperties = readNumberingProperties;

var dingbatToUnicode = require("dingbat-to-unicode");
var _ = require("underscore");

var documents = require("../documents");
var Result = require("../results").Result;
var warning = require("../results").warning;
var uris = require("./uris");

function createBodyReader(options) {
    return {
        readXmlElement: function(element) {
            return new BodyReader(options).readXmlElement(element);
        },
        readXmlElements: function(elements) {
            return new BodyReader(options).readXmlElements(elements);
        }
    };
}

function parseToNumber(value) {
    return isNaN(Number(value)) ? null : Number(value);
}

function BodyReader(options) {
    var complexFieldStack = [];
    var currentInstrText = [];
    var relationships = options.relationships;
    var contentTypes = options.contentTypes;
    var docxFile = options.docxFile;
    var files = options.files;
    var numbering = options.numbering;
    var styles = options.styles;

    function readXmlElements(elements) {
        var results = elements.map(readXmlElement);
        return combineResults(results);
    }

    function readXmlElement(element) {
        if (element.type === "element") {
            var handler = xmlElementReaders[element.name];
            if (handler) {
                return handler(element);
            } else if (!Object.prototype.hasOwnProperty.call(ignoreElements, element.name)) {
                var message = warning("An unrecognised element was ignored: " + element.name);
                return emptyResultWithMessages([message]);
            }
        }
        return emptyResult();
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

    function readRunProperties(element) {
        return readRunStyle(element).map(function(style) {
            var fontSizeString = element.firstOrEmpty("w:sz").attributes["w:val"];
            // w:sz gives the font size in half points, so halve the value to get the size in points
            var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : null;

            return {
                type: "runProperties",
                styleId: style.styleId,
                styleName: style.name,
                'vertical-alignment': readSize(element.first("w:vertAlign")),
                fontSize: fontSize,
                bold: readBooleanElement(element.first("w:b")),
                'bold-complex-script': readBooleanElement(element.first("w:bCs")),
                italics: readBooleanElement(element.first("w:i")),
                'italics-complex-script': readBooleanElement(element.first("w:iCs")),
                underline: readUnderline(element.first("w:u")),
                effect: readSize(element.first("w:effect")),
                emphasisMark: '',
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
        });
    }

    function readBodyProperties(element) {
        var customDocDesc = {
            page: reaBodyPageProperties(element),
            grid: readBodyGrid(element.firstOrEmpty('w:docGrid')),
            headerWrapperGroup: readBodyWrapperGroup(element.firstOrEmpty('w:headerReference')),
            footerWrapperGroup: readBodyWrapperGroup(element.firstOrEmpty('w:footerReference ')),
            lineNumbers: readBodyLineNumbers(element.firstOrEmpty('w:lineNumbers')),
            // w:titlePg
            titlePage: element.firstOrEmpty('w:titlePg').attributes['w:val'],
            // w:vAlign attr w:val
            verticalAlign: element.firstOrEmpty('w:vAlign').attributes['w:val'],
            column: readBodyColumn(element.firstOrEmpty('w:cols')),
            // w:type w:val
            type: element.firstOrEmpty('w:type').attributes['w:val']
        };
        console.log(customDocDesc, 'customDocDesc');
        return readCustomDocDescEl(filterObject(customDocDesc));
    }

    function reaBodyPageProperties(element) {
        var page = {
            size: reaBodyPageSizeProperties(element.firstOrEmpty('w:pgSz')),
            margin: reaBodyPageMarginProperties(element.firstOrEmpty('w:pgMar')),
            pageNumbers: reaBodyPageNumbersProperties(element.firstOrEmpty('w:pgNumType')),
            borders: readTableCellBorders(element.firstOrEmpty('w:w:pgBorders')), // w:pgBorders
            textDirection: element.firstOrEmpty('w:w:pgBorders').attributes['w:val'] // <w:textDirection w:val="lrTb"/>
        };
        console.log(filterObject(page), 'filterObject(page)');
        return filterObject(page);
    }

    function reaBodyPageSizeProperties(element) {
        // <w:pgSz w:w="11906" w:h="16838" w:orient="w:orient" />
        var pageSize = {
            width: parseToNumber(element.attributes['w:w']),
            height: parseToNumber(element.attributes['w:h']),
            orientation: element.attributes['w:orient']
        };
        return filterObject(pageSize);
    }

    function reaBodyPageMarginProperties(element) {
        // <w:pgMar w:top="1383" w:right="1800" w:bottom="1270" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
        var margins = {
            top: parseToNumber(element.attributes['w:top']),
            right: parseToNumber(element.attributes['w:right']),
            bottom: parseToNumber(element.attributes['w:bottom']),
            left: parseToNumber(element.attributes['w:left']),
            header: parseToNumber(element.attributes['w:header']),
            footer: parseToNumber(element.attributes['w:footer']),
            gutter: parseToNumber(element.attributes['w:gutter'])
        };
        return filterObject(margins);
    }

    function reaBodyPageNumbersProperties(element) {
        // <w:pgNumType w:start="1" w:fmt="w:chapSep" w:chapSep="hyphen" />
        var numbers = {
            start: parseToNumber(element.attributes['w:start']),
            formatType: element.attributes['w:fmt'],
            separator: element.attributes['w:chapSep']
        };
        return filterObject(numbers);
    }

    function readBodyGrid(element) {
        // <w:docGrid w:type="lines" w:linePitch="312" w:w:charSpace="111"/>
        var grid = {
            type: element.attributes['w:type'],
            linePitch: parseToNumber(element.attributes['w:linePitch']),
            charSpace: parseToNumber(element.attributes['w:charSpace'])
        };
        return filterObject(grid);
    }

    function readBodyWrapperGroup(element) {
        // <w:headerReference w:type="default" r:id="rId38"/>
        // <w:footerReference w:type="default" r:id="rId39"/>
        var group = {
            type: element.attributes['w:type'],
            id: element.attributes['r:id']
        };
        return filterObject(group);
    }

    function readBodyLineNumbers(element) {
        // "w:lnNumType": { _attr: { "w:countBy": 2, "w:distance": 4, "w:restart": "continuous", "w:start": 2 } },
        var lineNumbers = {
            countBy: parseToNumber(element.attributes['w:countBy']),
            distance: parseToNumber(element.attributes['w:distance']),
            restart: element.attributes['w:restart'],
            start: parseToNumber(element.attributes['w:start'])
        };
        return filterObject(lineNumbers);
    }

    function readBodyColumn(element) {
        // <w:cols w:space="720" w:num w:sep w:equalWidth />
        var column = {
            // readonly space?: number | PositiveUniversalMeasure;
            // readonly count?: number;
            // readonly separate?: boolean;
            // readonly equalWidth?: boolean;
            // readonly children?: readonly Column[];
            space: parseToNumber(element.attributes['w:space']),
            count: parseToNumber(element.attributes['w:count']),
            separate: element.attributes['w:separate'] ? true : false,
            equalWidth: element.attributes['w:equalWidth'] ? true : false
        };
        return filterObject(column);
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

    function createLanguageComponent(element) {
        var language = {
            value: element.attributes['w:val'],
            eastAsia: element.attributes['w:eastAsia'],
            bidirectional: element.attributes['w:bidi']
        };
        return filterObject(language);
    }

    function readBooleanElement(element) {
        if (element) {
            var value = element.attributes["w:val"];
            return value !== "false" && value !== "0";
        } else {
            return false;
        }
    }

    function readParagraphStyle(element) {
        return readStyle(element, "w:pStyle", "Paragraph", styles.findParagraphStyleById);
    }

    function readRunStyle(element) {
        return readStyle(element, "w:rStyle", "Run", styles.findCharacterStyleById);
    }

    function readTableStyle(element) {
        return readStyle(element, "w:tblStyle", "Table", styles.findTableStyleById);
    }

    function readStyle(element, styleTagName, styleType, findStyleById) {
        var messages = [];
        var styleElement = element.first(styleTagName);
        var styleId = null;
        var name = null;
        if (styleElement) {
            styleId = styleElement.attributes["w:val"];
            if (styleId) {
                var style = findStyleById(styleId);
                if (style) {
                    name = style.name;
                } else {
                    messages.push(undefinedStyleWarning(styleType, styleId));
                }
            }
        }
        return elementResultWithMessages({styleId: styleId, name: name}, messages);
    }

    var unknownComplexField = {type: "unknown"};

    function readFldChar(element) {
        var type = element.attributes["w:fldCharType"];
        if (type === "begin") {
            complexFieldStack.push(unknownComplexField);
            currentInstrText = [];
        } else if (type === "end") {
            complexFieldStack.pop();
        } else if (type === "separate") {
            var hyperlinkOptions = parseHyperlinkFieldCode(currentInstrText.join(''));
            var complexField = hyperlinkOptions === null ? unknownComplexField : {type: "hyperlink", options: hyperlinkOptions};
            complexFieldStack.pop();
            complexFieldStack.push(complexField);
        }
        return emptyResult();
    }

    function currentHyperlinkOptions() {
        var topHyperlink = _.last(complexFieldStack.filter(function(complexField) {
            return complexField.type === "hyperlink";
        }));
        return topHyperlink ? topHyperlink.options : null;
    }

    function parseHyperlinkFieldCode(code) {
        var externalLinkResult = /\s*HYPERLINK "(.*)"/.exec(code);
        if (externalLinkResult) {
            return {href: externalLinkResult[1]};
        }

        var internalLinkResult = /\s*HYPERLINK\s+\\l\s+"(.*)"/.exec(code);
        if (internalLinkResult) {
            return {anchor: internalLinkResult[1]};
        }

        return null;
    }

    function readInstrText(element) {
        currentInstrText.push(element.text());
        return emptyResult();
    }

    function readSymbol(element) {
        // See 17.3.3.30 sym (Symbol Character) of ECMA-376 4th edition Part 1
        var font = element.attributes["w:font"];
        var char = element.attributes["w:char"];
        var unicodeCharacter = dingbatToUnicode.hex(font, char);
        if (unicodeCharacter == null && /^F0..$/.test(char)) {
            unicodeCharacter = dingbatToUnicode.hex(font, char.substring(2));
        }

        if (unicodeCharacter == null) {
            return emptyResultWithMessages([warning(
                "A w:sym element with an unsupported character was ignored: char " +  char + " in font " + font
            )]);
        } else {
            return elementResult(new documents.Text(unicodeCharacter.string));
        }
    }

    function noteReferenceReader(noteType) {
        return function(element) {
            var noteId = element.attributes["w:id"];
            return elementResult(new documents.NoteReference({
                noteType: noteType,
                noteId: noteId
            }));
        };
    }

    function readCommentReference(element) {
        return elementResult(documents.commentReference({
            commentId: element.attributes["w:id"]
        }));
    }

    function readChildElements(element) {
        return readXmlElements(element.children);
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

    // 新增获取行间距和首行缩进方法
    function readSpacing(element, type) {
        var spacing = {
            line: parseToNumber(element.attributes['w:val']),
            lineRule: parseToNumber(element.attributes['w:lineRule']),
            before: parseToNumber(element.attributes['w:beforeAutospacing']),
            after: parseToNumber(element.attributes['w:afterAutospacing'])
        };
 
        return filterObject(spacing);
    }

    function readCharacterSpacing(element) {
        return parseToNumber(element.attributes['w:val']);
    }
  

    var xmlElementReaders = {
        "w:p": function(element) {
            return readXmlElements(element.children)
                .map(function(children) {
                    var properties = _.find(children, isParagraphProperties);
                    return new documents.Paragraph(
                        children.filter(negate(isParagraphProperties)),
                        properties
                    );
                })
                .insertExtra();
        },
        "w:pPr": function(element) {
            return readParagraphStyle(element).map(function(style) {
                var paragraphProp = {
                    type: "paragraphProperties",
                    styleId: style.styleId,
                    styleName: style.name,
                    heading: element.firstOrEmpty('w:pStyle').attributes['w:val'], // w:pStyle
                    numbering: readNumberingProperties(style.styleId, element.firstOrEmpty("w:numPr"), numbering),
                    keepNext: readBooleanElement(element.first("w:keepNext")),
                    keepLines: readBooleanElement(element.first("w:keepLines")),
                    pageBreakBefore: element.firstOrEmpty('w:pageBreakBefore').attributes['w:val'],
                    widowControl: readBooleanElement(element.first("w:widowControl")),
                    alignment: element.firstOrEmpty("w:jc").attributes["w:val"],
                    indent: readParagraphIndent(element.firstOrEmpty("w:ind")),
                    bgColor: readColor(element.firstOrEmpty('w:shd')), // 添加背景色的解析
                    spacing: readSpacing(element.firstOrEmpty('w:spacing'), 'w:p'), // 添加间距和缩进解析
                    border: readTableCellBorders(element.firstOrEmpty('w:pBdr')),
                    bidirectional: readBooleanElement(element.first("w:bidi")),
                    shading: readShading(element.firstOrEmpty("w:shd")),
                    suppressLineNumbers: readBooleanElement(element.first("w:suppressLineNumbers")),
                    wordWrap: readBooleanElement(element.first("w:suppressLineNumbers")),
                    scale: readSize(element.first("w:w")),
                    bullet: readNumberingProperties(style.styleId, element.firstOrEmpty("w:numPr"), numbering)
                    // TODO: 未知属性
                    // tabStops: '',
                    // frame: '',
                };
                return filterObject(paragraphProp);
            });
        },
        "w:r": function(element) {
            return readXmlElements(element.children)
                .map(function(children) {
                    var properties = _.find(children, isRunProperties) || {};
                    if (element.attributes['commentId']) {
                        properties['commentId'] = element.attributes['commentId'];
                    }
                    
                    children = children.filter(negate(isRunProperties));

                    var hyperlinkOptions = currentHyperlinkOptions();
                    if (hyperlinkOptions !== null) {
                        children = [new documents.Hyperlink(children, hyperlinkOptions)];
                    }

                    return new documents.Run(children, properties);
                });
        },
        "w:commentRangeStart": function(element) {
            return emptyResult();
        },
        "w:commentRangeEnd": function(element) {
            return emptyResult();
        },
        "w:sectPr": readBodyProperties,
        "w:rPr": readRunProperties,
        "w:fldChar": readFldChar,
        "w:instrText": readInstrText,
        "w:t": function(element) {
            return elementResult(new documents.Text(element.text()));
        },
        "w:tab": function(element) {
            return elementResult(new documents.Tab());
        },
        "w:noBreakHyphen": function() {
            return elementResult(new documents.Text("\u2011"));
        },
        "w:softHyphen": function(element) {
            return elementResult(new documents.Text("\u00AD"));
        },
        "w:sym": readSymbol,
        "w:hyperlink": function(element) {
            var relationshipId = element.attributes["r:id"];
            var anchor = element.attributes["w:anchor"];
            return readXmlElements(element.children).map(function(children) {
                function create(options) {
                    var targetFrame = element.attributes["w:tgtFrame"] || null;

                    return new documents.Hyperlink(
                        children,
                        _.extend({targetFrame: targetFrame}, options)
                    );
                }

                if (relationshipId) {
                    var href = relationships.findTargetByRelationshipId(relationshipId);
                    if (anchor) {
                        href = uris.replaceFragment(href, anchor);
                    }
                    return create({href: href});
                } else if (anchor) {
                    return create({anchor: anchor});
                } else {
                    return children;
                }
            });
        },
        "w:tbl": readTable,
        "w:tr": readTableRow,
        "w:tc": readTableCell,
        "w:footnoteReference": noteReferenceReader("footnote"),
        "w:endnoteReference": noteReferenceReader("endnote"),
        "w:commentReference": readCommentReference,
        "w:br": function(element) {
            var breakType = element.attributes["w:type"];
            if (breakType == null || breakType === "textWrapping") {
                return elementResult(documents.lineBreak);
            } else if (breakType === "page") {
                return elementResult(documents.pageBreak);
            } else if (breakType === "column") {
                return elementResult(documents.columnBreak);
            } else {
                return emptyResultWithMessages([warning("Unsupported break type: " + breakType)]);
            }
        },
        "w:bookmarkStart": function(element){
            var name = element.attributes["w:name"];
            if (name === "_GoBack") {
                return emptyResult();
            } else {
                return elementResult(new documents.BookmarkStart({name: name}));
            }
        },

        "mc:AlternateContent": function(element) {
            return readChildElements(element.first("mc:Fallback"));
        },

        "w:sdt": function(element) {
            return readXmlElements(element.firstOrEmpty("w:sdtContent").children);
        },

        "w:ins": readChildElements,
        "w:object": readChildElements,
        "w:smartTag": readChildElements,
        "w:drawing": function(element) {
            return readChildElements(element);
        },
        "w:pict": function(element) {
            return readChildElements(element).toExtra();
        },
        "v:roundrect": readChildElements,
        "v:shape": readChildElements,
        "v:textbox": readChildElements,
        "w:txbxContent": readChildElements,
        "wp:inline": readDrawingElement,
        "wp:anchor": readDrawingElement,
        "v:imagedata": readImageData,
        "v:group": readChildElements,
        "v:rect": readChildElements
    };

    return {
        readXmlElement: readXmlElement,
        readXmlElements: readXmlElements
    };


    function readTable(element) {
        var propertiesResult = readTableProperties(element.firstOrEmpty("w:tblPr"), element);
        return readXmlElements(element.children)
            .flatMap(calculateRowSpans)
            .flatMap(function(children) {
                return propertiesResult.map(function(properties) {
                    return documents.Table(children, properties);
                });
            });
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

    function readTableProperties(element, elements) {
        return readTableStyle(element).map(function(style) {
            return {
                styleId: style.styleId,
                styleName: style.name,
                float: readTableFloat(element.firstOrEmpty("w:tblpPr"), element),
                width: readTableWidthElement(element.firstOrEmpty("w:tblW")),
                indent: readTableWidthElement(element.firstOrEmpty("w:tblInd")),
                layout: element.firstOrEmpty("w:tblLayout").attributes["w:type"],
                cellMargin: readTableCellMargins(element.firstOrEmpty("w:tblCellMar")),
                columnWidths: tableGrid(elements.firstOrEmpty("w:tblGrid"))
            };
        });
    }

    function readTableRowProperties(element) {
        var isHeader = !!element.first("w:tblHeader");
        var heighProperties = element.firstOrEmpty("w:trHeight");
        
        var height = {
            value: parseToNumber(heighProperties.attributes["w:val"]),
            rule: heighProperties.attributes["w:hRule"]
        };
        // w:trHeight w:hRule w:val
        // TODO: w:cantSplit
        return {
            isHeader: isHeader,
            height: filterObject(height)
        };
    }

    function tableGrid(element) {
        var columnWidths = [];
        element.children.map(function(child) {
            columnWidths.push(parseToNumber(child.attributes["w:w"]));
        });
        return columnWidths;
    }

    function readTableRow(element) {
        var properties = element.firstOrEmpty("w:trPr");
        return readXmlElements(element.children).map(function(children) {
            return documents.TableRow(children, readTableRowProperties(properties));
        });
    }

    function readTableCell(element) {
        return readXmlElements(element.children).map(function(children) {
            var properties = element.firstOrEmpty("w:tcPr");
            var propertiesResult = readTableCellProperties(properties);

            var cell = documents.TableCell(children, propertiesResult);
            cell._vMerge = readVMerge(properties);
            return cell;
        });
    }

    function readTableCellProperties(element) {
        var gridSpan = element.firstOrEmpty("w:gridSpan").attributes["w:val"];
        var colSpan = gridSpan ? parseInt(gridSpan, 10) : 1;
        return {
            colSpan: colSpan,
            width: readTableWidthElement(element.firstOrEmpty("w:tcW")),
            verticalMerge: element.firstOrEmpty("w:vMerge").attributes['w:val'],
            borders: readTableCellBorders(element.firstOrEmpty("w:tcBorders")),
            shading: readShading(element.firstOrEmpty("w:shd")),
            verticalAlign: element.firstOrEmpty("w:vAlign").attributes['w:val'],
            margins: readTableCellMargins(element.firstOrEmpty("w:tcMar"))
        };
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

    function readShading(element) {
        var shading = {
            fill: element.attributes["w:fill"],
            color: element.attributes["w:color"],
            type: element.attributes["w:val"]
        };
        return filterObject(shading);
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

    function readVMerge(properties) {
        var element = properties.first("w:vMerge");
        if (element) {
            var val = element.attributes["w:val"];
            return val === "continue" || !val;
        } else {
            return null;
        }
    }

    function calculateRowSpans(rows) {
        var unexpectedNonRows = _.any(rows, function(row) {
            return row.type !== documents.types.tableRow;
        });
        if (unexpectedNonRows) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-row element in table, cell merging may be incorrect"
            )]);
        }
        var unexpectedNonCells = _.any(rows, function(row) {
            return _.any(row.children, function(cell) {
                return cell.type !== documents.types.tableCell;
            });
        });
        if (unexpectedNonCells) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-cell element in table row, cell merging may be incorrect"
            )]);
        }

        var columns = {};

        rows.forEach(function(row) {
            var cellIndex = 0;
            row.children.forEach(function(cell) {
                if (cell._vMerge && columns[cellIndex]) {
                    columns[cellIndex].rowSpan++;
                } else {
                    columns[cellIndex] = cell;
                    cell._vMerge = false;
                }
                cellIndex += cell.colSpan;
            });
        });

        rows.forEach(function(row) {
            row.children = row.children.filter(function(cell) {
                return !cell._vMerge;
            });
            row.children.forEach(function(cell) {
                delete cell._vMerge;
            });
        });

        return elementResult(rows);
    }

    function readDrawingElement(element) {
        var graphicData = element
            .getElementsByTagName("a:graphic")
            .getElementsByTagName("a:graphicData");
        var blips = graphicData
            .getElementsByTagName("pic:pic")
            .getElementsByTagName("pic:blipFill")
            .getElementsByTagName("a:blip");
        var chart = graphicData
            .getElementsByTagName("chart:chart");
        if (blips[0]) {
            return readDrawingImageElement(blips, element);
        }

        if (chart[0]) {
            return readDrawingChartElement(chart, element);
        }
    }

    /**
     * 渲染图片
     * */
    function readDrawingImageElement(blips, element) {
        return combineResults(blips.map(readBlip.bind(null, element)));
    }

    /**
     * 渲染图表
     * */
    function readDrawingChartElement(chart, element) {
        return combineResults(chart.map(readChart.bind(null, element)));
    }

    function readChart(element, chart) {
        var properties = element.first("wp:docPr").attributes;
        var altText = isBlank(properties.descr) ? properties.title : properties.descr;
        var blipImageFile = findChart(chart);
        if (element.attributes['commentId']) {
            blipImageFile['commentId'] = element.attributes['commentId'];
        }
        if (blipImageFile === null) {
            return emptyResultWithMessages([warning("Could not id for a:chart element")]);
        } else {
            return readChartEl(blipImageFile, altText);
        }
    }

    function readBlip(element, blip) {
        var properties = element.first("wp:docPr").attributes;
        var altText = isBlank(properties.descr) ? properties.title : properties.descr;
        var blipImageFile = findBlipImageFile(blip);
        if (blipImageFile === null) {
            return emptyResultWithMessages([warning("Could not find image file for a:blip element")]);
        } else {
            return readImage(blipImageFile, altText);
        }
    }

    function isBlank(value) {
        return value == null || /^\s*$/.test(value);
    }

    function findBlipImageFile(blip) {
        var embedRelationshipId = blip.attributes["r:embed"];
        var linkRelationshipId = blip.attributes["r:link"];
        if (embedRelationshipId) {
            return findEmbeddedImageFile(embedRelationshipId);
        } else if (linkRelationshipId) {
            var imagePath = relationships.findTargetByRelationshipId(linkRelationshipId);
            return {
                path: imagePath,
                read: files.read.bind(files, imagePath)
            };
        } else {
            return null;
        }
    }

    function findChart(chart) {
        var embedRelationshipId = chart.attributes["r:id"];
        if (embedRelationshipId) {
            return {
                id: embedRelationshipId
            };
        } else {
            return null;
        }
    }

    function readImageData(element) {
        var relationshipId = element.attributes['r:id'];

        if (relationshipId) {
            return readImage(
                findEmbeddedImageFile(relationshipId),
                element.attributes["o:title"]);
        } else {
            return emptyResultWithMessages([warning("A v:imagedata element without a relationship ID was ignored")]);
        }
    }

    function findEmbeddedImageFile(relationshipId) {
        var path = uris.uriToZipEntryName("word", relationships.findTargetByRelationshipId(relationshipId));
        return {
            path: path,
            read: docxFile.read.bind(docxFile, path)
        };
    }

    function readImage(imageFile, altText) {
        var contentType = contentTypes.findContentType(imageFile.path);
        var image = documents.Image({
            readImage: imageFile.read,
            altText: altText,
            contentType: contentType
        });
        var warnings = supportedImageTypes[contentType] ?
            [] : warning("Image of type " + contentType + " is unlikely to display in web browsers");
        return elementResultWithMessages(image, warnings);
    }

    function readChartEl(chart, altText) {
        var chartEl = documents.Chart({
            id: chart.id,
            altText: altText,
            commentId: chart.commentId
        });
        return elementResultWithMessages(chartEl, []);
    }

    // TODO: 新增document描述相关属性
    function readCustomDocDescEl(customDocDescProperties) {
        var customDocDescEl = documents.CustomDocDesc(customDocDescProperties);
        return elementResultWithMessages(customDocDescEl, []);
    }

    function undefinedStyleWarning(type, styleId) {
        return warning(
            type + " style with ID " + styleId + " was referenced but not defined in the document");
    }
}


function readNumberingProperties(styleId, element, numbering) {
    if (styleId != null) {
        var levelByStyleId = numbering.findLevelByParagraphStyleId(styleId);
        if (levelByStyleId != null) {
            return levelByStyleId;
        }
    }

    var level = element.firstOrEmpty("w:ilvl").attributes["w:val"];
    var numId = element.firstOrEmpty("w:numId").attributes["w:val"];
    if (level === undefined || numId === undefined) {
        return null;
    } else {
        return numbering.findLevel(numId, level);
        
    }
}

var supportedImageTypes = {
    "image/png": true,
    "image/gif": true,
    "image/jpeg": true,
    "image/svg+xml": true,
    "image/tiff": true
};

var ignoreElements = {
    "office-word:wrap": false,
    "v:shadow": false,
    "v:shapetype": false,
    "w:annotationRef": false,
    "w:bookmarkEnd": false,
    "w:sectPr": false,
    "w:proofErr": false,
    "w:lastRenderedPageBreak": false,
    "w:commentRangeStart": false,
    "w:commentRangeEnd": false,
    "w:del": false,
    "w:footnoteRef": false,
    "w:endnoteRef": false,
    "w:tblPr": false,
    "w:tblGrid": false,
    "w:trPr": false,
    "w:tcPr": false
};

function isParagraphProperties(element) {
    return element.type === "paragraphProperties";
}

function isRunProperties(element) {
    return element.type === "runProperties";
}

function negate(predicate) {
    return function(value) {
        return !predicate(value);
    };
}


function emptyResultWithMessages(messages) {
    return new ReadResult(null, null, messages);
}

function emptyResult() {
    return new ReadResult(null);
}

function elementResult(element) {
    return new ReadResult(element);
}

function elementResultWithMessages(element, messages) {
    return new ReadResult(element, null, messages);
}

function ReadResult(element, extra, messages) {
    this.value = element || [];
    this.extra = extra;
    this._result = new Result({
        element: this.value,
        extra: extra
    }, messages);
    this.messages = this._result.messages;
}

ReadResult.prototype.toExtra = function() {
    return new ReadResult(null, joinElements(this.extra, this.value), this.messages);
};

ReadResult.prototype.insertExtra = function() {
    var extra = this.extra;
    if (extra && extra.length) {
        return new ReadResult(joinElements(this.value, extra), null, this.messages);
    } else {
        return this;
    }
};

ReadResult.prototype.map = function(func) {
    var result = this._result.map(function(value) {
        return func(value.element);
    });
    return new ReadResult(result.value, this.extra, result.messages);
};

ReadResult.prototype.flatMap = function(func) {
    var result = this._result.flatMap(function(value) {
        return func(value.element)._result;
    });
    return new ReadResult(result.value.element, joinElements(this.extra, result.value.extra), result.messages);
};

function combineResults(results) {
    var result = Result.combine(_.pluck(results, "_result"));
    return new ReadResult(
        _.flatten(_.pluck(result.value, "element")),
        _.filter(_.flatten(_.pluck(result.value, "extra")), identity),
        result.messages
    );
}

function joinElements(first, second) {
    return _.flatten([first, second]);
}

function identity(value) {
    return value;
}
