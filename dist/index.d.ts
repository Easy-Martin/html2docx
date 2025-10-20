/**
 * 将HTML转换为Word文档
 * @param {string} html - 输入的HTML内容
 * @returns {Promise<Blob>} - 返回Word文档的Blob对象
 */
declare function htmlToDocx(html: string): Promise<Blob>;
export default htmlToDocx;
