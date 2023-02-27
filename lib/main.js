/* eslint-disable no-console */
/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-14 22:48:05
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-17 19:15:54
 * @FilePath: \mammoth.js\lib\main.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
/* global process */

var fs = require("fs");
var path = require("path");

var mammoth = require("./");
var promises = require("./promises");
var images = require("./images");

function main(argv) {
    var docxPath = argv["docx-path"];
    var outputPath = argv["output-path"];
    var outputDir = argv.output_dir;
    var outputFormat = argv.output_format;
    var styleMapPath = argv.style_map;
    
    readStyleMap(styleMapPath).then(function(styleMap) {
        var options = {
            styleMap: styleMap,
            outputFormat: outputFormat
        };
        
        if (outputDir) {
            var basename = path.basename(docxPath, ".docx");
            outputPath = path.join(outputDir, basename + ".html");
            var imageIndex = 0;
            options.convertImage = images.imgElement(function(element) {
                imageIndex++;
                var extension = element.contentType.split("/")[1];
                var filename = imageIndex + "." + extension;
                
                return element.read().then(function(imageBuffer) {
                    var imagePath = path.join(outputDir, filename);
                    return promises.nfcall(fs.writeFile, imagePath, imageBuffer);
                }).then(function() {
                    return {src: filename};
                });
            });
        }
        
        return mammoth.convert({path: docxPath}, options)
            .then(function(result) {
                result.messages.forEach(function(message) {
                    process.stderr.write(message.message);
                    process.stderr.write("\n");
                });
                
                var outputStream = outputPath ? fs.createWriteStream(outputPath) : process.stdout;
                
                outputStream.write(result.value);
            });
    }).done();
}

function readStyleMap(styleMapPath) {
    if (styleMapPath) {
        return promises.nfcall(fs.readFile, styleMapPath, "utf8");
    } else {
        return promises.resolve(null);
    }
}

module.exports = main;
