/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-16 00:27:07
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-25 21:55:18
 * @FilePath: \mammoth.js\lib\xml\reader.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
/* eslint-disable no-console */
var promises = require("../promises");
var _ = require("underscore");

var xmldom = require("./xmldom");
var nodes = require("./nodes");
var Element = nodes.Element;

exports.readString = readString;

var Node = xmldom.Node;

function readString(xmlString, namespaceMap) {
    namespaceMap = namespaceMap || {};
    var attribute = null;

    try {
        var document = xmldom.parseFromString(xmlString, "text/xml");
    } catch (error) {
        return promises.reject(error);
    }

    if (document.documentElement.tagName === "parsererror") {
        return promises.resolve(new Error(document.documentElement.textContent));
    }

    function convertNode(node) {
        switch (node.nodeType) {
        case Node.ELEMENT_NODE:
            if (node.tagName === 'w:commentRangeStart') {
                attribute = {};
                attribute['commentId'] = 'comment-' + node.attributes[0]['nodeValue'];
            }
            if (node.tagName === 'w:commentRangeEnd') {
                attribute = null;
            }
            
            return convertElement(node, attribute);
        case Node.TEXT_NODE:
            return nodes.text(node.nodeValue);
        }
    }

    function convertElement(element, otherAttribute) {
        var convertedName = convertName(element);
        var convertedChildren = [];
        _.forEach(element.childNodes, function(childNode) {
            var convertedNode = convertNode(childNode);
            if (convertedNode) {
                convertedChildren.push(convertedNode);
            }
        });

        var convertedAttributes = {};
        _.forEach(element.attributes, function(attribute) {
            convertedAttributes[convertName(attribute)] = attribute.value;
        });
        return new Element(convertedName, Object.assign({}, convertedAttributes, otherAttribute), convertedChildren);
    }

    function convertName(node) {
        if (node.namespaceURI) {
            var mappedPrefix = namespaceMap[node.namespaceURI];
            var prefix;
            if (mappedPrefix) {
                prefix = mappedPrefix + ":";
            } else {
                prefix = "{" + node.namespaceURI + "}";
            }
            return prefix + node.localName;
        } else {
            return node.localName;
        }
    }

    return promises.resolve(convertNode(document.documentElement));
}
