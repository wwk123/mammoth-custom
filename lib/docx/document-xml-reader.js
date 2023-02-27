/* eslint-disable no-console */
/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-14 22:48:05
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-25 17:29:51
 * @FilePath: \mammoth.js\lib\docx\document-xml-reader.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
exports.DocumentXmlReader = DocumentXmlReader;

var documents = require("../documents");
var Result = require("../results").Result;

function DocumentXmlReader(options) {
    var bodyReader = options.bodyReader;
    
    function convertXmlToDocument(element) {
        var body = element.first("w:body");
        var sectPr = body.first("w:sectPr");
        console.log(element, 'element');
        console.log(sectPr, 'sectPr');
        
        var result = bodyReader.readXmlElements(body.children)
            .map(function(children) {
                return new documents.Document(children, {
                    notes: options.notes,
                    comments: options.comments
                });
            });
        return new Result(result.value, result.messages);
    }
    
    return {
        convertXmlToDocument: convertXmlToDocument
    };
}
