/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-16 00:26:57
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-16 15:15:52
 * @FilePath: \mammoth.js\lib\docx\office-xml-reader.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
var _ = require("underscore");

var promises = require("../promises");
var xml = require("../xml");


exports.read = read;
exports.readXmlFromZipFile = readXmlFromZipFile;

var xmlNamespaceMap = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main": "w",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing": "wp",
    "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
    "http://schemas.openxmlformats.org/drawingml/2006/picture": "pic",
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "chart",
    "http://schemas.openxmlformats.org/package/2006/content-types": "content-types",
    "urn:schemas-microsoft-com:vml": "v",
    "http://schemas.openxmlformats.org/markup-compatibility/2006": "mc",
    "urn:schemas-microsoft-com:office:word": "office-word"
};


function read(xmlString) {
    return xml.readString(xmlString, xmlNamespaceMap)
        .then(function(document) {
            return collapseAlternateContent(document)[0];
        });
}


function readXmlFromZipFile(docxFile, path) {
    if (docxFile.exists(path)) {
        return docxFile.read(path, "utf-8")
            .then(stripUtf8Bom)
            .then(read);
    } else {
        return promises.resolve(null);
    }
}


function stripUtf8Bom(xmlString) {
    return xmlString.replace(/^\uFEFF/g, '');
}


function collapseAlternateContent(node) {
    if (node.type === "element") {
        if (node.name === "mc:AlternateContent") {
            return node.first("mc:Fallback").children;
        } else {
            node.children = _.flatten(node.children.map(collapseAlternateContent, true));
            return [node];
        }
    } else {
        return [node];
    }
}
