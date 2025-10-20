"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const docx_1 = require("docx");
const utils_1 = require("./utils");
/**
 * 将HTML转换为Word文档
 * @param {string} html - 输入的HTML内容
 * @returns {Promise<Blob>} - 返回Word文档的Blob对象
 */
async function htmlToDocx(html) {
    // 创建DOM解析器
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    // 创建Word文档
    // 遍历HTML节点并转换为Word元素
    const body = doc.body;
    const children = [];
    for (let i = 0; i < body.childNodes.length; i++) {
        const node = body.childNodes[i];
        if (node.nodeName === 'P') {
            children.push(await createParagraphNode(node));
        }
        if (node.nodeName.match(/^H[1-6]$/)) {
            children.push(await createHeadingNode(node));
        }
        if (node.nodeName === 'UL' || node.nodeName === 'OL') {
            const list = await createListNode(node);
            children.push(...list);
        }
        if (node.nodeName === 'TABLE') {
            const table = await createTableNode(node);
            children.push(table);
        }
    }
    const docx = new docx_1.Document({
        sections: [{ children: children }],
        numbering: {
            config: [
                {
                    reference: 'bullet-points',
                    levels: [{ level: 0, format: 'bullet', text: '•', style: { paragraph: { indent: { left: 0, firstLine: 0 } } } }],
                },
                {
                    reference: 'numbered-list',
                    levels: [{ level: 0, format: 'decimal', text: '%1.', style: { paragraph: { indent: { left: 0, firstLine: 0 } } } }],
                },
            ],
        },
    });
    // 生成Word文档
    return await docx_1.Packer.toBlob(docx);
}
/**
 * 获取文本样式
 * @param {HTMLElement} node - 输入的HTML元素
 * @param {HTMLElement} [childNode] - 子元素
 * @returns {{[key:string]:any}} - 返回文本样式对象
 */
function getTextStyle(node, childNode) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t;
    const indent = (_a = node.style) === null || _a === void 0 ? void 0 : _a.textIndent; // 默认应该只存在 p 标签上
    const color = ((_b = node.style) === null || _b === void 0 ? void 0 : _b.color) || ((_c = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _c === void 0 ? void 0 : _c.color);
    const fontSize = ((_d = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _d === void 0 ? void 0 : _d.fontSize) || ((_e = node.style) === null || _e === void 0 ? void 0 : _e.fontSize);
    // 解析行高
    const lineHeight = ((_f = node.style) === null || _f === void 0 ? void 0 : _f.lineHeight) || ((_g = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _g === void 0 ? void 0 : _g.lineHeight) ? utils_1.UnitConverter.parseLineHeight(((_h = node.style) === null || _h === void 0 ? void 0 : _h.lineHeight) || ((_j = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _j === void 0 ? void 0 : _j.lineHeight)) : null;
    const style = {
        color: color ? (0, utils_1.rgbToHex)(color) : undefined,
        size: fontSize ? parseInt(fontSize) * 1.6 : 16,
        bold: ((_k = node.style) === null || _k === void 0 ? void 0 : _k.fontWeight) === 'bold' || ((_l = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _l === void 0 ? void 0 : _l.fontWeight) === 'bold',
        italics: ((_m = node.style) === null || _m === void 0 ? void 0 : _m.fontStyle) === 'italic' || ((_o = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _o === void 0 ? void 0 : _o.fontStyle) === 'italic',
        underline: ((_p = node.style) === null || _p === void 0 ? void 0 : _p.textDecoration) === 'underline' || ((_q = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _q === void 0 ? void 0 : _q.textDecoration) === 'underline',
        font: ((_r = node.style) === null || _r === void 0 ? void 0 : _r.fontFamily) || ((_s = childNode === null || childNode === void 0 ? void 0 : childNode.style) === null || _s === void 0 ? void 0 : _s.fontFamily),
        alignment: (_t = node.style) === null || _t === void 0 ? void 0 : _t.textAlign,
        indent: { firstLine: parseInt(indent || '0'), start: 0, left: 0 },
        spacing: lineHeight
            ? lineHeight.type === 'multiple'
                ? { line: lineHeight.value * 240 } // 240 twips = 12pt，作为基准值
                : { line: lineHeight.value, lineRule: docx_1.LineRuleType.EXACT }
            : undefined,
    };
    return style;
}
/**
 * 创建段落元素
 * @param {HTMLElement} node - 输入的HTML元素
 * @param {Paragraph} [paragraph] - 段落对象
 * @returns {Promise<Paragraph>} - 返回创建的段落对象
 */
async function createParagraphNode(node, paragraph) {
    const style = getTextStyle(node);
    if (!paragraph) {
        paragraph = new docx_1.Paragraph({ alignment: style.alignment, indent: style.indent, spacing: style.spacing });
    }
    for (const child of Array.from(node.childNodes)) {
        switch (child.nodeName) {
            case 'SPAN':
                const childStyle = getTextStyle(child, node);
                paragraph.addChildElement(await createChildNode(child, { ...style, ...childStyle }));
                break;
            case 'TEXT':
                if (child.textContent !== '') {
                    paragraph.addChildElement(await createTextNode(child, style));
                }
                break;
            case 'STRONG':
                paragraph.addChildElement(await createTextNode(child, { ...style, bold: true }));
                break;
            case 'IMG':
                paragraph.addChildElement(await createImageNode(child));
                break;
            case 'BR':
                paragraph.addChildElement(new docx_1.TextRun({ text: '\n' }));
                break;
            default:
                paragraph.addChildElement(await createTextNode(child, style));
                break;
        }
    }
    return paragraph;
}
/**
 *
 * @param node
 * @param style
 * @returns
 */
async function createChildNode(node, style = {}) {
    let childNode = new docx_1.TextRun({ text: '' });
    for (const child of Array.from(node.childNodes)) {
        switch (child.nodeName) {
            case 'STRONG':
                childNode = await createChildNode(child, { ...style, bold: true });
                break;
            case 'IMG':
                childNode = await createImageNode(child);
                break;
            case 'BR':
                childNode = new docx_1.TextRun({ text: '\n' });
                break;
            default:
                childNode = await createTextNode(child, style);
        }
    }
    return childNode;
}
/**
 * 创建标题元素
 * @param {HTMLElement} node - 输入的HTML元素
 * @returns {Promise<Table>} - 创建的标题对象
 */
async function createTableNode(node) {
    const rows = [];
    for (const tr of Array.from(node.querySelectorAll('tr'))) {
        const cells = [];
        for (const td of Array.from(tr.querySelectorAll('td, th'))) {
            const cellChildren = [];
            for (const child of Array.from(td.childNodes)) {
                cellChildren.push(new docx_1.Paragraph({ spacing: { line: 20 * 20 }, indent: { firstLine: 20 }, children: [new docx_1.TextRun({ text: child.textContent || '', size: 18 })] }));
            }
            const size = td.getAttribute('width') && td.getAttribute('width') !== 'auto' ? parseInt(td.getAttribute('width')) : 100;
            // 处理单元格边框
            const border = { size: 1, color: '#000000', style: 'single' };
            const cell = new docx_1.TableCell({
                children: cellChildren,
                columnSpan: td.getAttribute('colspan') ? parseInt(td.getAttribute('colspan')) : undefined,
                rowSpan: td.getAttribute('rowspan') ? parseInt(td.getAttribute('rowspan')) : undefined,
                width: { size, type: 'auto' },
                borders: { top: border, bottom: border, left: border, right: border },
            });
            cells.push(cell);
        }
        rows.push(new docx_1.TableRow({ children: cells }));
    }
    return new docx_1.Table({ rows, width: { size: 100, type: 'pct' } });
}
/**
 * 创建列表节点
 * @param node - HTML列表元素(ul或ol)
 * @returns 由段落组成的数组，每个段落代表一个列表项
 */
async function createListNode(node) {
    const list = [];
    for (const li of Array.from(node.childNodes)) {
        if (li.textContent !== '') {
            const numbering = li.textContent === '\n' ? undefined : { reference: node.nodeName === 'UL' ? 'bullet-points' : 'numbered-list', level: 0 };
            const listItem = await createParagraphNode(li, new docx_1.Paragraph({ numbering }));
            list.push(listItem);
        }
    }
    return list;
}
/**
 * 创建文本元素
 * @param {HTMLElement} imgNode - 输入的HTML元素
 * @returns {Promise<ImageRun>} - 创建的文本对象
 */
async function createImageNode(imgNode) {
    var _a, _b, _c;
    const imagePath = (_a = imgNode.getAttribute('src')) !== null && _a !== void 0 ? _a : '';
    let imageData;
    if (imagePath.startsWith('data:image')) {
        // 处理base64图片
        const base64Data = imagePath.split(',')[1];
        // 将base64字符串转换为Uint8Array
        imageData = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));
    }
    else {
        // 处理普通URL图片
        imageData = await fetch(imagePath)
            .then(res => res.blob())
            .then(blob => blob.arrayBuffer())
            .then(buffer => new Uint8Array(buffer))
            .catch(e => {
            console.error(e);
            return;
        });
    }
    // 获取图片宽度和高度
    const width = parseInt((_b = imgNode.getAttribute('width')) !== null && _b !== void 0 ? _b : '650');
    const height = parseInt((_c = imgNode.getAttribute('height')) !== null && _c !== void 0 ? _c : '280');
    // 处理图片，使用docx库中的ImageRun组件，需要将图片转换为Uint8Array或Blob对象，这里使用base64字符串作为示例，实际使用时需要根据实际情况进行处理
    return new docx_1.ImageRun({ data: imageData, transformation: { width, height } });
}
/**
 * 创建标题元素
 * @param {HTMLElement} node - 输入的HTML元素
 * @returns {Promise<Paragraph>} - 创建的标题对象
 */
async function createHeadingNode(node) {
    const style = getTextStyle(node);
    const level = parseInt(node.nodeName.substring(1));
    return new docx_1.Paragraph({
        alignment: style.alignment,
        heading: (docx_1.HeadingLevel.HEADING_1 + (level - 1)),
        children: [new docx_1.TextRun({ text: node.textContent, bold: true })],
    });
}
/**
 * 创建文本元素
 * @param {HTMLElement} node - 输入的HTML元素
 * @param {TextStyle} [style] - 文本样式
 * @returns {Promise<TextRun>} - 创建的文本对象
 */
async function createTextNode(node, style = {}) {
    return new docx_1.TextRun({
        text: node.textContent || '',
        font: style.font,
        size: style.size,
        color: style.color,
        bold: style.bold,
        italics: style.italics,
        underline: style.underline,
    });
}
exports.default = htmlToDocx;
