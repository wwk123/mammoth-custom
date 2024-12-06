/* eslint-disable no-console */
var _ = require("underscore");

var promises = require("./promises");
var documents = require("./documents");
var htmlPaths = require("./styles/html-paths");
var results = require("./results");
var images = require("./images");
var Html = require("./html");
var writers = require("./writers");

exports.DocumentConverter = DocumentConverter;


function DocumentConverter(options) {
    return {
        convertToHtml: function(element) {
            var comments = _.indexBy(
                element.type === documents.types.document ? element.comments : [],
                "commentId"
            );
            var conversion = new DocumentConversion(options, comments);
            return conversion.convertToHtml(element);
        }
    };
}

function DocumentConversion(options, comments) {
    var noteNumber = 1;

    var noteReferences = [];

    var referencedComments = [];

    options = _.extend({ignoreEmptyParagraphs: true}, options);
    var idPrefix = options.idPrefix === undefined ? "" : options.idPrefix;
    var ignoreEmptyParagraphs = options.ignoreEmptyParagraphs;

    var defaultParagraphStyle = htmlPaths.topLevelElement("p");

    var styleMap = options.styleMap || [];

    function appendDocxAttributes(element, options) {
        var appendAttr = {};
        if (element['bullet']) {
            element['bullet']['level'] = Number(element['bullet']['level']);
        }
        for (var key in element) {
            if (Object.hasOwnProperty.call(element, key)) {
                var ele = element[key];
                if (key !== 'children' && ele) {
                    var value = typeof ele === 'object' ? JSON.stringify(ele) : ele;
                    appendAttr['data-' + key] = value;
                }
            }
        }
        if (!element['children']) {
            element['children'] = [];
        }
        
        return Object.assign({}, options, appendAttr);
    }

    function convertStyleForAttributes(element, options) {
        var styleMap = [];
        if (!options) {
            options = {};
        }
        var ptToEm = (element.fontSize / 12).toFixed(2);
        if (element.alignment && element.alignment !== 'both') {
            styleMap.push({key: 'text-align', value: element.alignment});
        }

        if (element.color) {
            styleMap.push({key: 'color', value: element.color});
        }
        if (element.bgColor) {
            styleMap.push({key: 'backgroundColor', value: element.bgColor});
        }
        if (element.font) {
            styleMap.push({key: 'fontFamily', value: element.font.cs || element.font.eastAsia || '仿宋'});
        }
        if (element.fontSize) {
            styleMap.push({key: 'fontSize', value: ptToEm + 'em'});
        }
        if (element.underline) {
            styleMap.push({key: 'text-decoration', value: 'underline'});
        }
        if (element.spacing) {
            styleMap.push({key: 'lineHeight', value: '1.5'});
        }
    
        var styleStr = styleMap
          .map(function(item) {
              return item.key + ':' + item.value;
          })
          .join(';');
        if (styleStr) {
            options.style = styleStr + ';' + (options.style || '');
        }
        return appendDocxAttributes(element, options);
    }
    
    function convertParagraphStyleForAttributes(element, options) {
        var styleMap = [];
        if (!options) {
            options = {};
        }
        var ptToEm = (element.fontSize / 12).toFixed(2);
        if (element.alignment && element.alignment !== 'both') {
            styleMap.push({key: 'text-align', value: element.alignment});
        } else {
            styleMap.push({key: 'text-align', value: 'justify'});
        }
        if (element.color) {
            styleMap.push({key: 'color', value: element.color});
        }
        if (element.bgColor) {
            styleMap.push({key: 'background-color', value: element.bgColor});
        }
        if (element.font) {
            styleMap.push({key: 'font-family', value: element.font});
        }
        if (element.fontSize) {
            styleMap.push({key: 'font-size', value: ptToEm + 'em'});
        }
        if (element.isUnderline) {
            styleMap.push({key: 'text-decoration', value: 'underline'});
        }
        if (element.isStrikethrough) {
            options["data-strike"] = element.isStrikethrough;
        }
        if (element.spacing) {
            styleMap.push({key: 'line-height', value: '1.5'});
        }
        if (element.indent) {
            if (element.indent['firstLine'] === '420') {
                styleMap.push({key: 'text-indent', value: '2em'});
            } else if (element.indent['firstLine'] === '640') {
                styleMap.push({key: 'text-indent', value: '2.66em'});
            } else if (element.indent['firstLine'] === '643') {
                styleMap.push({key: 'text-indent', value: '2.66em'});
            } else if (element.indent['firstLine'] === '880') {
                styleMap.push({key: 'text-indent', value: '3.1em'});
            }
        }
        var styleStr = styleMap
          .map(function(item) {
              return item.key + ':' + item.value;
          })
          .join(';');
        if (styleStr) {
            options.style = styleStr + ';' + (options.style || '');
        }
        return appendDocxAttributes(element, options);
    }

    function ptToEm(val) {
        return _.isNaN(val / 4) ? '1px' : Math.floor((val / 4)) + 'px';
    }

    function createBorderValue(options) {
        var borderStyleMap = {
            single: 'solid',
            nil: 'dashed'
        };
        var size = ptToEm(options.size);
        var style = borderStyleMap[options.style] || 'solid';
        var color = createColor(options.color);
        return size + ' ' + style + ' ' + color;
    }

    function createColor(color) {
        return color === 'auto' || !color ? '#92CDDC' : ('#' + color);
    }

    function convertBorderStyleForAttributes(element, options) {
        var styleMap = [];
        if (!options) {
            options = {};
        }
        if (options.borders) {
            if (options.borders.top) {
                styleMap.push({key: 'border-top', value: createBorderValue(options.borders.top)});
            }
            if (options.borders.bottom) {
                styleMap.push({key: 'border-bottom', value: createBorderValue(options.borders.bottom)});
            }
            if (options.borders.left) {
                styleMap.push({key: 'border-left', value: createBorderValue(options.borders.left)});
            }
            if (options.borders.right) {
                styleMap.push({key: 'border-right', value: createBorderValue(options.borders.right)});
            }
        } else {
            styleMap.push({key: 'border', value: '1px solid #92CDDC'});
        }

        if (element.shading) {
            if (element.shading.fill) {
                styleMap.push({key: 'background-color', value: element.shading.fill});
            }
        }

        delete options.borders;
        delete options.shading;
        
        var styleStr = styleMap
          .map(function(item) {
              return item.key + ':' + item.value;
          })
          .join(';');
        if (styleStr) {
            options.style = styleStr + ';' + (options.style || '');
        }
        return appendDocxAttributes(element, options);
    }

    function convertToHtml(document) {
        var messages = [];

        var html = elementToHtml(document, messages, {});

        var deferredNodes = [];
        walkHtml(html, function(node) {
            if (node.type === "deferred") {
                deferredNodes.push(node);
            }
        });
        var deferredValues = {};
        return promises.mapSeries(deferredNodes, function(deferred) {
            return deferred.value().then(function(value) {
                deferredValues[deferred.id] = value;
            });
        }).then(function() {
            function replaceDeferred(nodes) {
                return flatMap(nodes, function(node) {
                    if (node.type === "deferred") {
                        return deferredValues[node.id];
                    } else if (node.children) {
                        return [
                            _.extend({}, node, {
                                children: replaceDeferred(node.children)
                            })
                        ];
                    } else {
                        return [node];
                    }
                });
            }
            var writer = writers.writer({
                prettyPrint: options.prettyPrint,
                outputFormat: options.outputFormat
            });
            
            Html.write(writer, Html.simplify(replaceDeferred(html)));
            return new results.Result(writer.asString(), messages);
        });
    }

    function convertElements(elements, messages, options) {
        return flatMap(elements, function(element) {
            return elementToHtml(element, messages, options);
        });
    }

    function elementToHtml(element, messages, options) {
        if (!options) {
            throw new Error("options not set");
        }
        if (element.type === 'commentReference') {
            // console.log(element, 'commentReference');
        }
        var handler = elementConverters[element.type];
        if (handler) {
            return handler(element, messages, options);
        } else {
            return [];
        }
    }

    function convertParagraph(element, messages, options) {
        return htmlPathForParagraph(element, messages).wrap(function() {
            var content = convertElements(element.children, messages, options);
            if (ignoreEmptyParagraphs) {
                return content;
            } else {
                return [Html.forceWrite].concat(content);
            }
        });
    }

    function htmlPathForParagraph(element, messages) {
        var style = findStyle(element);

        if (style) {
            return style.to;
        } else {
            if (element.styleId) {
                messages.push(unrecognisedStyleWarning("paragraph", element));
            }

            return htmlPaths.topLevelElement('p', convertParagraphStyleForAttributes(element));
            // return defaultParagraphStyle;
        }
    }

    function convertRun(run, messages, options) {
        console.log(run, 'run ---------------------------------')
        var nodes = function() {
            return convertElements(run.children, messages, options);
        };
        var paths = [];
        // 添加了一个是否有额外的标签的判断
        var tagNumber = 0;
        // 提前生成属性，在生成标签的时候可以顺便传入
        var attributes = convertStyleForAttributes(run);
        if (run.isSmallCaps) {
            paths.push(findHtmlPathForRunProperty("smallCaps"));
        }
        if (run.isAllCaps) {
            paths.push(findHtmlPathForRunProperty("allCaps"));
        }
        if (run.isStrikethrough) {
            paths.push(findHtmlPathForRunProperty("strikethrough", "s", attributes));
            tagNumber++;
        }
        if (run.underline) {
            paths.push(findHtmlPathForRunProperty("underline"));
        }
        if (run.verticalAlignment === documents.verticalAlignment.subscript) {
            paths.push(htmlPaths.element("sub", attributes, {fresh: false}));
            tagNumber++;
        }
        if (run.verticalAlignment === documents.verticalAlignment.superscript) {
            paths.push(htmlPaths.element("sup", attributes, {fresh: false}));
            tagNumber++;
        }
        if (run.italics) {
            paths.push(findHtmlPathForRunProperty("italic", "em", attributes));
            tagNumber++;
        }
        if (run.bold) {
            paths.push(findHtmlPathForRunProperty("bold", "strong", attributes));
            tagNumber++;
        }

        var stylePath = htmlPaths.empty;
        var style = findStyle(run);
        if (style) {
            stylePath = style.to;
        } else if (run.styleId) {
            messages.push(unrecognisedStyleWarning("run", run));
        }
        // 最后兜底，如果上面没生成标签，并且也有样式的话，就生成一个span标签插入
        if (attributes && attributes.style && tagNumber == 0) {
            paths.push(htmlPaths.element('span', attributes, {fresh: true}));
        }
        paths.push(stylePath);

        paths.forEach(function(path) {
            nodes = path.wrap.bind(path, nodes);
        });

        return nodes();
    }

    function findHtmlPathForRunProperty(elementType, defaultTagName, attributes) {
        var path = findHtmlPath({type: elementType});
        if (path) {
            return path;
        } else if (defaultTagName) {
            return htmlPaths.element(defaultTagName, attributes, {fresh: false});
        } else {
            return htmlPaths.empty;
        }
    }

    function findHtmlPath(element, defaultPath) {
        var style = findStyle(element);
        return style ? style.to : defaultPath;
    }

    function findStyle(element) {
        for (var i = 0; i < styleMap.length; i++) {
            if (styleMap[i].from.matches(element)) {
                return styleMap[i];
            }
        }
    }

    function recoveringConvertImage(convertImage) {
        return function(image, messages) {
            return promises.attempt(function() {
                return convertImage(image, messages);
            }).caught(function(error) {
                messages.push(results.error(error));
                return [];
            });
        };
    }

    function noteHtmlId(note) {
        return referentHtmlId(note.noteType, note.noteId);
    }

    function noteRefHtmlId(note) {
        return referenceHtmlId(note.noteType, note.noteId);
    }

    function referentHtmlId(referenceType, referenceId) {
        return htmlId(referenceType + "-" + referenceId);
    }

    function referenceHtmlId(referenceType, referenceId) {
        return htmlId(referenceType + "-ref-" + referenceId);
    }

    function htmlId(suffix) {
        return idPrefix + suffix;
    }

    var defaultTablePath = htmlPaths.elements([
        htmlPaths.element("table", {style: "border-collapse: collapse;width: 100%;"}, {fresh: true})
    ]);

    function convertTable(element, messages, options) {
        var optionResult = {
            style: "border-collapse: collapse;width: 100%;"
        };
        if (element.columnWidths) {
            optionResult['data-column-widths'] = JSON.stringify(element.columnWidths);
        }
        if (element.float) {
            optionResult['data-float']  = JSON.stringify(element.float);
        }
        if (element.indent) {
            optionResult['data-indent'] = JSON.stringify(element.indent);
        }
        if (element.layout) {
            optionResult['data-layout']  = JSON.stringify(element.layout);
        }
        if (element.cellMargin) {
            optionResult['data-cell-margin'] = JSON.stringify(element.cellMargin);
        }
        if (element.overlap) {
            optionResult['data-overlap'] = element.overlap;
        }
        if (element.width) {
            optionResult['data-width'] = JSON.stringify(element.width);
        }
        if (element.styleId) {
            optionResult.styleId = element.styleId;
            optionResult['style-id'] = element.styleId;
        }
        if (element.visuallyRightToLeft) {
            optionResult['visually-right-to-left'] = element.visuallyRightToLeft;
        }
        if (element.borders) {
            optionResult['borders'] = JSON.stringify(element.borders);
        }
        if (element.shading) {
            optionResult['shading'] = JSON.stringify(element.shading);
        }
        if (element.alignment) {
            optionResult['alignment'] = element.alignment;
        }
        if (element.styleName) {
            optionResult.styleName = element.styleName;
        }
        
        var tablePath = htmlPaths.elements([
            htmlPaths.element("table", optionResult, {fresh: true})
        ]);

        return findHtmlPath(element, tablePath).wrap(function() {
            return convertTableChildren(element, messages, options);
        });
    }

    function convertTableChildren(element, messages, options) {
        var bodyIndex = _.findIndex(element.children, function(child) {
            return !child.type === documents.types.tableRow || !child.isHeader;
        });
        if (bodyIndex === -1) {
            bodyIndex = element.children.length;
        }
        var children;
        if (bodyIndex === 0) {
            children = convertElements(
                element.children,
                messages,
                _.extend({}, options, {isTableHeader: false})
            );
        } else {
            var headRows = convertElements(
                element.children.slice(0, bodyIndex),
                messages,
                _.extend({}, options, {isTableHeader: true})
            );
            var bodyRows = convertElements(
                element.children.slice(bodyIndex),
                messages,
                _.extend({}, options, {isTableHeader: false})
            );
            children = [
                Html.freshElement("thead", {}, headRows),
                Html.freshElement("tbody", {}, bodyRows)
            ];
        }
        return [Html.forceWrite].concat(children);
    }

    function convertTableRow(element, messages, options) {
        var rowOptions = {};
        if (element.height) {
            rowOptions['data-height'] = JSON.stringify(element.height);
        }
        if (element.isHeader) {
            rowOptions.isHeader = element.isHeader;
        }
        var children = convertElements(element.children, messages, options);
        return [
            Html.freshElement("tr", rowOptions, [Html.forceWrite].concat(children))
        ];
    }

    function convertTableCell(element, messages, options) {
        var tagName = options.isTableHeader ? "th" : "td";
        var children = convertElements(element.children, messages, options);
        var attributes = {};
        if (element.colSpan !== 1) {
            attributes.colspan = element.colSpan.toString();
        }
        if (element.rowSpan !== 1) {
            attributes.rowspan = element.rowSpan.toString();
        }
        if (element.borders) {
            attributes['data-borders'] = JSON.stringify(element.borders);
        }
        if (element.shading) {
            attributes['data-shading'] = JSON.stringify(element.shading);
        }
        if (element.width) {
            attributes['data-width'] = JSON.stringify(element.width);
        }
        if (element.verticalMerge) {
            attributes['data-vertical-merge'] = JSON.stringify(element.verticalMerge);
        }
        if (element.verticalAlign) {
            attributes['data-vertical-align'] = element.verticalAlign;
        }
        if (element.margins) {
            attributes['data-margins'] = JSON.stringify(element.margins);
        }
        if (element.textDirection) {
            attributes['data-text-direction'] = JSON.stringify(element.textDirection);
        }

        return [
            Html.freshElement(tagName, convertBorderStyleForAttributes(element, attributes), [Html.forceWrite].concat(children))
        ];
    }

    function convertCommentReference(reference, messages, options) {
        return findHtmlPath(reference, htmlPaths.ignore).wrap(function() {
            var comment = comments[reference.commentId];
            var count = referencedComments.length + 1;
            var label = "[" + commentAuthorLabel(comment) + count + "]";
            referencedComments.push({label: label, comment: comment});
            // TODO: remove duplication with note references
            return [
                Html.freshElement("a", {
                    href: "#" + referentHtmlId("comment", reference.commentId),
                    id: referenceHtmlId("comment", reference.commentId)
                }, [Html.text(label)])
            ];
        });
    }

    function convertComment(referencedComment, messages, options) {
        // TODO: remove duplication with note references

        var label = referencedComment.label;
        var comment = referencedComment.comment;
        var body = convertElements(comment.body, messages, options).concat([
            Html.nonFreshElement("p", {}, [
                Html.text(" "),
                Html.freshElement("a", {"href": "#" + referenceHtmlId("comment", comment.commentId)}, [
                    Html.text("↑")
                ])
            ])
        ]);

        return [
            Html.freshElement(
                "dt",
                {"id": referentHtmlId("comment", comment.commentId)},
                [Html.text("Comment " + label)]
            ),
            Html.freshElement("dd", {}, body)
        ];
    }

    function convertBreak(element, messages, options) {
        return htmlPathForBreak(element).wrap(function() {
            return [];
        });
    }

    function convertCustomDocDesc(element, messages, options) {
        return [];
    }

    function htmlPathForBreak(element) {
        var style = findStyle(element);
        if (style) {
            return style.to;
        } else if (element.breakType === "line") {
            return htmlPaths.topLevelElement("br");
        } else {
            return htmlPaths.empty;
        }
    }

    var elementConverters = {
        "document": function(document, messages, options) {
            var children = convertElements(document.children, messages, options);
            var notes = noteReferences.map(function(noteReference) {
                return document.notes.resolve(noteReference);
            });
            var notesNodes = convertElements(notes, messages, options);
            return children.concat([
                Html.freshElement("ol", {}, notesNodes),
                Html.freshElement("dl", {}, flatMap(referencedComments, function(referencedComment) {
                    return convertComment(referencedComment, messages, options);
                }))
            ]);
        },
        "paragraph": convertParagraph,
        "run": convertRun,
        "text": function(element, messages, options) {
            return [Html.text(element.value)];
        },
        "tab": function(element, messages, options) {
            return [Html.text("\t")];
        },
        "hyperlink": function(element, messages, options) {
            var href = element.anchor ? "#" + htmlId(element.anchor) : element.href;
            var attributes = {href: href};
            if (element.targetFrame != null) {
                attributes.target = element.targetFrame;
            }

            var children = convertElements(element.children, messages, options);
            return [Html.nonFreshElement("a", attributes, children)];
        },
        "bookmarkStart": function(element, messages, options) {
            var anchor = Html.freshElement("a", {
                id: htmlId(element.name)
            }, [Html.forceWrite]);
            return [anchor];
        },
        "noteReference": function(element, messages, options) {
            noteReferences.push(element);
            var anchor = Html.freshElement("a", {
                href: "#" + noteHtmlId(element),
                id: noteRefHtmlId(element)
            }, [Html.text("[" + (noteNumber++) + "]")]);

            return [Html.freshElement("sup", {}, [anchor])];
        },
        "note": function(element, messages, options) {
            var children = convertElements(element.body, messages, options);
            var backLink = Html.elementWithTag(htmlPaths.element("p", {}, {fresh: false}), [
                Html.text(" "),
                Html.freshElement("a", {href: "#" + noteRefHtmlId(element)}, [Html.text("↑")])
            ]);
            var body = children.concat([backLink]);

            return Html.freshElement("li", {id: noteHtmlId(element)}, body);
        },
        "commentReference": convertCommentReference,
        "comment": convertComment,
        "image": deferredConversion(recoveringConvertImage(options.convertImage || images.dataUri)),
        "chart": function(element, messages, options) {
            var anchor = Html.freshElement("p", {
                id: element.id,
                'data-commentId': element.commentId
            }, [Html.forceWrite]);
            return [anchor];
        },
        'customDocDesc': convertCustomDocDesc,
        "table": convertTable,
        "tableRow": convertTableRow,
        "tableCell": convertTableCell,
        "break": convertBreak
    };
    return {
        convertToHtml: convertToHtml
    };
}

var deferredId = 1;

function deferredConversion(func) {
    return function(element, messages, options) {
        return [
            {
                type: "deferred",
                id: deferredId++,
                value: function() {
                    return func(element, messages, options);
                }
            }
        ];
    };
}

function unrecognisedStyleWarning(type, element) {
    if (type === 'paragraph') {
        // console.log(type, element, 'unrecognisedStyleWarning');
    }
    return results.warning(
        "Unrecognised " + type + " style: '" + element.styleName + "'" +
        " (Style ID: " + element.styleId + ")"
    );
}

function flatMap(values, func) {
    return _.flatten(values? values.map(func) : [], true);
}

function walkHtml(nodes, callback) {
    nodes.forEach(function(node) {
        callback(node);
        if (node.children) {
            walkHtml(node.children, callback);
        }
    });
}

var commentAuthorLabel = exports.commentAuthorLabel = function commentAuthorLabel(comment) {
    return comment.authorInitials || "";
};
