/*
 * @Author: wwk123 m17600463015@163.com
 * @Date: 2023-02-16 00:26:47
 * @LastEditors: wwk123 m17600463015@163.com
 * @LastEditTime: 2023-02-17 19:03:20
 * @FilePath: \mammoth.js\lib\results.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
var _ = require("underscore");


exports.Result = Result;
exports.success = success;
exports.warning = warning;
exports.error = error;


function Result(value, messages) {
    this.value = value;
    this.messages = messages || [];
}

Result.prototype.map = function(func) {
    return new Result(func(this.value), this.messages);
};

Result.prototype.flatMap = function(func) {
    var funcResult = func(this.value);
    return new Result(funcResult.value, combineMessages([this, funcResult]));
};

Result.prototype.flatMapThen = function(func) {
    var that = this;
    return func(this.value).then(function(otherResult) {
        return new Result(otherResult.value, combineMessages([that, otherResult]));
    });
};

Result.combine = function(results) {
    var values = _.flatten(_.pluck(results, "value"));
    var messages = combineMessages(results);
    return new Result(values, messages);
};

function success(value) {
    return new Result(value, []);
}

function warning(message) {
    return {
        type: "warning",
        message: message
    };
}

function error(exception) {
    return {
        type: "error",
        message: exception.message,
        error: exception
    };
}

function combineMessages(results) {
    var messages = [];
    _.flatten(_.pluck(results, "messages"), true).forEach(function(message) {
        if (!containsMessage(messages, message)) {
            messages.push(message);
        }
    });
    return messages;
}

function containsMessage(messages, message) {
    return _.find(messages, isSameMessage.bind(null, message)) !== undefined;
}

function isSameMessage(first, second) {
    return first.type === second.type && first.message === second.message;
}
