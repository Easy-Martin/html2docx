# html-convert-docx

一个将 HTML 字符串转换为 Word (`.docx`) 文件的轻量库，基于 [`docx`](https://www.npmjs.com/package/docx)。适用于浏览器环境，将常见 HTML 标签与部分样式映射为 Word 文档内容。

## 特性
- 支持标签：`p`、`h1`~`h6`、`ul/ol/li`、`table/tr/td/th`、`img`、`span`、`strong`、`br`、文本节点
- 支持样式：颜色（`rgb(...)` 或 `#hex`）、字体大小（`px`）、加粗、斜体、下划线、字体、对齐、首行缩进、行高（数字/百分比/px/`normal`）
- 列表：圆点符号 `•` 的无序列表与十进制的有序列表
- 表格：支持 `colspan`、`rowspan`、`width` 属性，统一 1px 单线边框
- 图片：支持 `data:image/*;base64,...` 与网络图片（`fetch`）

## 安装
```bash
npm i html-convert-docx
# 或者
pnpm add html-convert-docx
```

> 依赖 `docx` 已在本包内声明，无需单独安装。

## 快速开始（浏览器）
```ts
import htmlToDocx from 'html-convert-docx';

const html = `
  <h1 style="text-align:center">示例文档</h1>
  <p style="text-indent:32px; line-height:1.6; color:rgb(34,34,34)">
    这是一段包含 <strong>加粗</strong>、<span style="font-style:italic">斜体</span>、
    <span style="text-decoration:underline">下划线</span> 的文本。
    <br/>
    下面是一个图片：
    <img src="data:image/png;base64,...." width="320" height="180" />
  </p>
  <ul>
    <li>项目一</li>
    <li>项目二</li>
  </ul>
  <table>
    <tr>
      <th width="120">表头A</th>
      <th width="180">表头B</th>
    </tr>
    <tr>
      <td colspan="1">单元格1</td>
      <td>单元格2</td>
    </tr>
  </table>
`;

// 生成 Blob 并触发下载
const blob = await htmlToDocx(html);
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'document.docx';
a.click();
URL.revokeObjectURL(url);
```

## API
- `htmlToDocx(html: string): Promise<Blob>`
  - 传入 HTML 字符串，返回生成的 `.docx` 文件 Blob（浏览器）。

## 实现说明（简要）
- 使用 `DOMParser` 解析 HTML 字符串，遍历 `body.childNodes`，将特定标签映射为 `docx` 的 `Paragraph`、`Table`、`ImageRun` 等对象。
- 文本样式来自元素的 `style`：
  - `color`：通过 `rgbToHex` 处理 `rgb(...)` 与 `#hex`
  - `size`：解析 `font-size` 的 `px` 值，内部按 docx 需求进行换算
  - `bold`/`italics`/`underline`/`font`/`alignment`
  - `indent.firstLine`：来自 `text-indent`
  - `spacing`（行高）：支持数字倍数、百分比、`px`、`normal`（按 1.2 倍解析），底层转换为 twips
- 列表：为 `ul/ol` 配置 `numbering`，无序列表用 `•`，有序列表用十进制；每个 `li` 转为一个段落。
- 表格：遍历 `tr`、`td/th`，支持 `colspan/rowspan/width`，统一设置四边 1px 单线边框，整体宽度 100%。
- 图片：
  - Base64：解析 `data:image/*;base64,...` 为 `Uint8Array`
  - URL：通过 `fetch` 获取并转 `Uint8Array`，默认尺寸为 `width=650`、`height=280`，可通过属性覆盖

## 兼容性与限制
- 首选浏览器环境：实现依赖 `DOMParser`、`fetch`，返回类型为 `Blob`。
  - Node 环境需要自行提供 `DOMParser` 与 `fetch` 的 polyfill，并将打包方式改为 `Packer.toBuffer` 等。
- 支持的标签与样式为常见子集，复杂嵌套/高级 CSS（如嵌套列表的层级样式、链接、行内复杂样式）可能无法完整映射。
- 远程图片受跨域限制；请求失败时图片可能无法插入。

## 在你的项目中使用
- 直接传入来自富文本编辑器（如 TipTap）或自定义模板生成的 HTML 字符串。
- 通过内联样式（`style="..."`）提供颜色、字体大小、行高、对齐、首行缩进等信息，以获得更接近预期的 Word 样式。

## 开发与构建
- 脚本：
  - `npm run dev`：TypeScript 监听编译
  - `npm run build`：TypeScript 构建到 `dist`
- TypeScript 配置：`lib` 目录为源码入口，输出到 `dist`

## 许可证
- [MIT](./LICENSE)

## 致谢
- 本项目使用了优秀的 [`docx`](https://www.npmjs.com/package/docx) 库来生成 Word 文档。