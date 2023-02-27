/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-14 22:48:05
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-20 15:29:48
 * @FilePath: \mammoth.js\lib\html\simplify.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
var _ = require("underscore");

var ast = require("./ast");

function simplify(nodes) {
    return collapse(removeEmpty(nodes));
}

function collapse(nodes) {
    var children = [];
    
    nodes.map(collapseNode).forEach(function(child) {
        appendChild(children, child);
    });
    return children;
}

function collapseNode(node) {
    return collapsers[node.type](node);
}

var collapsers = {
    element: collapseElement,
    text: identity,
    forceWrite: identity
};

function collapseElement(node) {
    return ast.elementWithTag(node.tag, collapse(node.children));
}

function identity(value) {
    return value;
}
// TODO: 元素复用的逻辑
function appendChild(children, child) {
    var lastChild = children[children.length - 1];
    if (child.type === "element" && !child.tag.fresh && lastChild && lastChild.type === "element" && child.tag.matchesElement(lastChild.tag)) {
        if (child.tag.separator) {
            appendChild(lastChild.children, ast.text(child.tag.separator));
        }
        child.children.forEach(function(grandChild) {
            // Mutation is fine since simplifying elements create a copy of the children.
            appendChild(lastChild.children, grandChild);
        });
    } else {
        children.push(child);
    }
}

function removeEmpty(nodes) {
    return flatMap(nodes, function(node) {
        return emptiers[node.type](node);
    });
}

function flatMap(values, func) {
    return _.flatten(_.map(values, func), true);
}

var emptiers = {
    element: elementEmptier,
    text: textEmptier,
    forceWrite: neverEmpty
};

function neverEmpty(node) {
    return [node];
}

function elementEmptier(element) {
    var children = removeEmpty(element.children);
    if (children.length === 0 && !ast.isVoidElement(element)) {
        return [];
    } else {
        return [ast.elementWithTag(element.tag, children)];
    }
}

function textEmptier(node) {
    if (node.value.length === 0) {
        return [];
    } else {
        return [node];
    }
}

module.exports = simplify;
