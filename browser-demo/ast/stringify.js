/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-19 10:11:53
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-21 10:33:13
 * @FilePath: \mammoth.js\browser-demo\ast\stringify.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */

function attrString(attrs) {
  const buff = [];
  for (let key in attrs) {
    buff.push(key + '="' + attrs[key] + '"');
  }
  if (!buff.length) {
    return '';
  }
  return ' ' + buff.join(' ');
}

function _stringify(buff, doc) {
  switch (doc.type) {
    case 'text':
      return buff + doc.content;
    case 'tag':
      buff +=
        '<' +
        doc.name +
        (doc.attrs ? attrString(doc.attrs) : '') +
        (doc.voidElement ? '/>' : '>');
      if (doc.voidElement) {
        return buff;
      }
      return buff + doc.children.reduce(_stringify, '') + '</' + doc.name + '>';
    case 'comment':
      buff += '<!--' + doc.comment + '-->';
      return buff;
    default:
      return '';
  }
}

export const stringify = (docx) => {
  return docx.reduce(function (token, rootEl) {
    return token + _stringify('', rootEl);
  }, '');
};