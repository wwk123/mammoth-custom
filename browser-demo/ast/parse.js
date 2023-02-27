/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-19 10:06:07
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-19 10:18:43
 * @FilePath: \mammoth.js\browser-demo\ast\parse.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
import { parseTag } from './parseTag.js';
const tagRE = /<[a-zA-Z0-9\-\!\/](?:"[^"]*"|'[^']*'|[^'">])*>/g;
const whitespaceRE = /^\s*$/;

// re-used obj for quick lookups of components
const empty = Object.create(null);

export const parse = (html, options = {}) => {
  options || (options = {});
  options.components || (options.components = empty);
  const result = [];
  const arr = [];
  let current;
  let level = -1;
  let inComponent = false;

  // handle text at top level
  if (html.indexOf('<') !== 0) {
    var end = html.indexOf('<');
    result.push({
      type: 'text',
      content: end === -1 ? html : html.substring(0, end),
    });
  }
  // @ts-ignore
  html.replace(tagRE, function (tag, index) {
    if (inComponent) {
      if (tag !== '</' + current.name + '>') {
        return '';
      } else {
        inComponent = false;
      }
    }
    const isOpen = tag.charAt(1) !== '/';
    const isComment = tag.startsWith('<!--');
    const start = index + tag.length;
    const nextChar = html.charAt(start);

    let parent;

    if (isComment) {
      const comment = parseTag(tag);

      // if we're at root, push new base node
      if (level < 0) {
        result.push(comment);
        return result
      }
      parent = arr[level];
      if (parent && parent.children && Array.isArray(parent.children)) {
        parent.children.push(comment);
      }
      return result
    }

    if (isOpen) {
      level++;

      current = parseTag(tag);
      if (
        current.type === 'tag' &&
        current.name &&
        options.components &&
        options.components[current.name]
      ) {
        current.type = 'component';
        inComponent = true;
      }

      if (
        !current.voidElement &&
        !inComponent &&
        nextChar &&
        nextChar !== '<' &&
        Array.isArray(current.children)
      ) {
        current.children.push({
          type: 'text',
          content: html.slice(start, html.indexOf('<', start)),
        });
      }

      // if we're at root, push new base node
      if (level === 0) {
        result.push(current);
      }

      parent = arr[level - 1];

      if (parent && parent.children) {
        parent.children.push(current);
      }

      arr[level] = current;
    }

    if (!isOpen || current.voidElement) {
      if (
        level > -1 &&
        (current.voidElement || current.name === tag.slice(2, -1))
      ) {
        level--;
        // move current up a level to match the end tag
        current = level === -1 ? result : arr[level];
      }
      if (!inComponent && nextChar !== '<' && nextChar) {
        // trailing text node
        // if we're at the root, push a base text node. otherwise add as
        // a child to the current node.
        parent = level === -1 ? result : arr[level].children;

        // calculate correct end of the content slice in case there's
        // no tag after the text node.
        const end = html.indexOf('<', start);
        let content = html.slice(start, end === -1 ? undefined : end);
        // if a node is nothing but whitespace, collapse it as the spec states:
        // https://www.w3.org/TR/html4/struct/text.html#h-9.1
        if (whitespaceRE.test(content)) {
          content = ' ';
        }
        // don't add whitespace-only text nodes if they would be trailing text nodes
        // or if they would be leading whitespace-only text nodes:
        //  * end > -1 indicates this is not a trailing text node
        //  * leading node is when level is -1 and parent has length 0
        if ((end > -1 && level + parent.length >= 0) || content !== ' ') {
          if (parent && Array.isArray(parent)) {
            parent.push({
              type: 'text',
              content: content,
            });
          }
        }
      }
    }
  });

  return result;
};