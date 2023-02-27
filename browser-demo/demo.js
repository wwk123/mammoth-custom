/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-14 22:48:05
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-23 08:47:46
 * @FilePath: \mammoth.js\browser-demo\demo.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
import { parse } from './ast/index.js';
(function() {
    document.getElementById("document")
        .addEventListener("change", handleFileSelect, false);
    const styleMap = {}
    function handleFileSelect(event) {
        readFileInputEventAsArrayBuffer(event, function(arrayBuffer) {
            /**
             * styleMap:控制Word样式到HTML的映射。如果选项。styleMap是一个字符串，
             * 每一行都被视为一个单独的样式映射，忽略空行和以#:If选项开头的行。
             * styleMap是一个数组，每个元素都是一个表示单个样式映射的字符串。
             * 有关样式映射语法的参考，请参见“编写样式映射”。
             * 
             * 
             * */ 
            // 
            // includeEmbeddedStyleMap:默认情况下，如果文档包含嵌入式样式映射，那么它将与默认样式映射结合。若要忽略任何嵌入式样式映射，请设置选项。includeEmbeddedStyleMap为false。

            // includeDefaultStyleMap:默认情况下，在styleMap中传递的样式映射与默认样式映射相结合。若要完全停止使用默认样式映射，请设置选项。includeDefaultStyleMap为false。

            // convertImage:默认情况下，图像被转换为<img>元素，源包含在src属性中。将此选项设置为图像转换器以覆盖默认行为。

            // ignoreEmptyParagraphs:默认情况下，空段落将被忽略。将此选项设置为false可在输出中保留空段落。

            // idPrefix:用于添加任何生成id的字符串，例如书签、脚注和尾注所使用的id。默认为空字符串。

            // transformDocument:如果设置了，该函数将应用于转换为HTML之前从docx文件读取的文档。文档转换的API应该被认为是不稳定的。参见文档转换。
            const monospaceFonts = ["consolas", "courier", "courier new"];
            
            function transformRun(run) {
                // var runs = mammoth.transforms.getDescendantsOfType(paragraph, 'run');
                if (run['font'] && run['fontSize']) {
                    styleMap[`${run['font']}-${run['fontSize']}`] = {
                        verticalAlignment: run['verticalAlignment'],
                        isBold: run['isBold'],
                        isItalic: run['isItalic'],
                    }
                    return {
                        ...run,
                        styleId: `${run['font']}-${run['fontSize']}`,
                        styleName: `${run['font']}-${run['fontSize']}`
                    };
                }else {
                    return run;
                }
            }
            const options = {
                styleMap: [
                    "p[style-name='Section Title'] => h1:fresh",
                    "p[style-name='Subsection Title'] => h2:fresh",
                    "comment-reference => sup"
                ],
                includeDefaultStyleMap: false,
                convertImage: mammoth.images.imgElement(function(image) {
                    return image.read("base64").then(function(imageBuffer) {
                        return {
                            src: "data:" + image.contentType + ";base64," + imageBuffer
                        };
                    });
                })
            }

            mammoth.convertToHtml({arrayBuffer: arrayBuffer}, options)
                .then(displayResult, function(error) {
                    console.error(error);
                });
        });
    }

    function displayResult(result) {
        console.log(result)
        console.log(styleMap)
        document.getElementById("output").innerHTML = result.value;

        var messageHtml = result.messages.map(function(message) {
            return '<li class="' + message.type + '">' + escapeHtml(message.message) + "</li>";
        }).join("");

        document.getElementById("messages").innerHTML = "<ul>" + messageHtml + "</ul>";
    }

    function readFileInputEventAsArrayBuffer(event, callback) {
        var file = event.target.files[0];

        var reader = new FileReader();

        reader.onload = function(loadEvent) {
            var arrayBuffer = loadEvent.target.result;
            callback(arrayBuffer);
        };

        reader.readAsArrayBuffer(file);
    }

    function escapeHtml(value) {
        return value
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
    }
    document.getElementById("btn").addEventListener('click', () => {
     // 设置每个节点标签属性
     // let attrRE = /\s([^'"/\s><]+?)[\s/>]|([^\s=]+)=\s?(".*?"|'.*?')/g;
      // function parseTag(tag) {
      //   let res = {
      //     type: "tag",
      //     name: "",
      //     voidElement: false,
      //     attrs: {},
      //     children: [],
      //   };
      //   let tagMatch = tag.match(/<\/?([^\s]+?)[/\s>]/);
      //   if (tagMatch) {
      //     // 标签名称为正则匹配的第2项
      //     res.name = tagMatch[1];
      //     if (tag.charAt(tag.length - 2) === "/") {
      //       // 判断tag字符串倒数第二项是不是 / 设置为空标签。 例子：<img/>
      //       res.voidElement = true;
      //     }
      //   }
      //   // 匹配所有的标签正则
      //   let classList = tag.match(/\s([^'"/\s><]+?)\s*?=\s*?(".*?"|'.*?')/g);

      //   if (classList && classList.length) {
      //     for (let i = 0; i < classList.length; i++) {
      //       // 去空格再以= 分隔字符串  得到['属性名称','属性值']
      //       let c = classList[i].replace(/\s*/g, "").split("=");
      //       // 循环设置属性
      //       if (c[1]) res.attrs[c[0]] = c[1].substring(1, c[1].length - 1);
      //     }
      //   }
      //   return res;
      // }

      // function parse(html) {
      //   let result = [];
      //   let current;
      //   let level = -1;
      //   let arr = [];
      //   let tagRE = /<[a-zA-Z\-\!\/](?:"[^"]*"['"]*|'[^']*'['"]*|[^'">])*>/g;

      //   html.replace(tagRE, function (tag, index) {
      //     // 判断第二个字符是不是'/'来判断是否open
      //     let isOpen = tag.charAt(1) !== "/";
      //     // 获取标签末尾的索引
      //     let start = index + tag.length;
      //     // 标签之前的文本信息
      //     let text = html.slice(start, html.indexOf("<", start));

      //     let parent;
      //     if (isOpen) {
      //       level++;
      //       // 设置标签属性
      //       current = parseTag(tag);
      //       // 判断是否为文本信息，是就push一个text children  不等于'  '
      //       if (!current.voidElement && text.trim()) {
      //         current.children.push({
      //           type: "text",
      //           content: text,
      //         });
      //       }
      //       // 如果我们是根用户，则推送新的基本节点
      //       if (level === 0) {
      //         result.push(current);
      //       }
      //       // 判断有没有上层，有就push当前标签
      //       parent = arr[level - 1];
      //       if (parent) {
      //         parent.children.push(current);
      //       }
      //       arr[level] = current;
      //     }
      //     // 如果不是开标签，或者是空元素：</div><img>
      //     if (!isOpen || current.voidElement) {
      //       // level--
      //       level--;
      //     }
      //   });
      //   return result;
      // }
        const innerHTML = document.getElementById('output').innerHTML
        // const parseHtml = new ParseHtml(innerHTML);
        // const astObj = parseHtml.parse();
        let ast = parse(innerHTML);
        console.log(ast);
    })
})();
