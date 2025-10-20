"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.UnitConverter = void 0;
exports.rgbToHex = rgbToHex;
function rgbToHex(a) {
    //RGB(204,204,024)
    //十六进制颜色值的正则表达式
    var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
    var that = a;
    if (/^(rgb|RGB)/.test(that)) {
        var aColor = that.replace(/(?:\(|\)|rgb|RGB)*/g, '').split(',');
        var strHex = '#';
        for (var i = 0; i < aColor.length; i++) {
            var hex = Number(aColor[i]).toString(16);
            if (hex === '0') {
                hex += hex;
            }
            strHex += hex;
        }
        if (strHex.length !== 7) {
            strHex = that;
        }
        return strHex;
    }
    if (reg.test(that)) {
        var aNum = that.replace(/#/, '').split('');
        if (aNum.length === 6) {
            return that;
        }
        else if (aNum.length === 3) {
            var numHex = '#';
            for (var i = 0; i < aNum.length; i += 1) {
                numHex += aNum[i] + aNum[i];
            }
            return numHex;
        }
    }
    return that;
}
// 单位转换工具函数
exports.UnitConverter = {
    // px转twips (1px ≈ 15 twips，基于1pt=1/72英寸，1px=1/96英寸)
    pxToTwips(px) {
        return Math.round(px * 15);
    },
    // 解析line-height值（支持数字、百分比、px单位）
    parseLineHeight(lineHeightValue) {
        // 默认字体大小16px
        if (!lineHeightValue)
            return null;
        // 处理倍数（如1.5）
        if (!isNaN(parseFloat(lineHeightValue)) && isFinite(lineHeightValue)) {
            return {
                type: 'multiple',
                value: parseFloat(lineHeightValue),
            };
        }
        // 处理百分比（如150%）
        if (lineHeightValue.endsWith('%')) {
            return {
                type: 'multiple',
                value: parseFloat(lineHeightValue) / 100,
            };
        }
        // 处理px单位（如24px）
        if (lineHeightValue.endsWith('px')) {
            const pxValue = parseFloat(lineHeightValue);
            return {
                type: 'exact',
                value: this.pxToTwips(pxValue),
            };
        }
        // 处理normal关键字（通常为1.2倍）
        if (lineHeightValue === 'normal') {
            return {
                type: 'multiple',
                value: 1.2,
            };
        }
        return null;
    },
};
