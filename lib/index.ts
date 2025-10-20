import { Document as DocumentDocx, Paragraph, TextRun, HeadingLevel, ImageRun, Packer, TableCell, TableRow, Table, LineRuleType } from 'docx';
import { type IBorderOptions } from 'docx';
import { rgbToHex, UnitConverter } from './utils';

/**
 * 将HTML转换为Word文档
 * @param {string} html - 输入的HTML内容
 * @returns {Promise<Blob>} - 返回Word文档的Blob对象
 */
async function htmlToDocx(html: string): Promise<Blob> {
  // 创建DOM解析器
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  // 创建Word文档

  // 遍历HTML节点并转换为Word元素
  const body = doc.body;
  const children = [];

  for (let i = 0; i < body.childNodes.length; i++) {
    const node = body.childNodes[i] as HTMLElement;
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

  const docx = new DocumentDocx({
    sections: [{ children: children as any }],
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
  return await Packer.toBlob(docx);
}

/**
 * 获取文本样式
 * @param {HTMLElement} node - 输入的HTML元素
 * @param {HTMLElement} [childNode] - 子元素
 * @returns {{[key:string]:any}} - 返回文本样式对象
 */
function getTextStyle(node: HTMLElement, childNode?: HTMLElement): { [key: string]: any } {
  const indent = node.style?.textIndent; // 默认应该只存在 p 标签上
  const color = node.style?.color || childNode?.style?.color;
  const fontSize = childNode?.style?.fontSize || node.style?.fontSize;

  // 解析行高
  const lineHeight = node.style?.lineHeight || childNode?.style?.lineHeight ? UnitConverter.parseLineHeight(node.style?.lineHeight || childNode?.style?.lineHeight) : null;

  const style = {
    color: color ? rgbToHex(color) : undefined,
    size: fontSize ? parseInt(fontSize) * 1.6 : 16,
    bold: node.style?.fontWeight === 'bold' || childNode?.style?.fontWeight === 'bold',
    italics: node.style?.fontStyle === 'italic' || childNode?.style?.fontStyle === 'italic',
    underline: node.style?.textDecoration === 'underline' || childNode?.style?.textDecoration === 'underline',
    font: node.style?.fontFamily || childNode?.style?.fontFamily,
    alignment: node.style?.textAlign,
    indent: { firstLine: parseInt(indent || '0'), start: 0, left: 0 },
    spacing: lineHeight
      ? lineHeight.type === 'multiple'
        ? { line: lineHeight.value * 240 } // 240 twips = 12pt，作为基准值
        : { line: lineHeight.value, lineRule: LineRuleType.EXACT }
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
async function createParagraphNode(node: HTMLElement, paragraph?: Paragraph): Promise<Paragraph> {
  const style = getTextStyle(node);
  if (!paragraph) {
    paragraph = new Paragraph({ alignment: style.alignment as any, indent: style.indent as any, spacing: style.spacing });
  }

  for (const child of Array.from(node.childNodes)) {
    switch (child.nodeName) {
      case 'SPAN':
        const childStyle = getTextStyle(child as HTMLElement, node);
        paragraph.addChildElement(await createChildNode(child as HTMLElement, { ...style, ...childStyle }));
        break;
      case 'TEXT':
        if (child.textContent !== '') {
          paragraph.addChildElement(await createTextNode(child as HTMLElement, style));
        }
        break;
      case 'STRONG':
        paragraph.addChildElement(await createTextNode(child as HTMLElement, { ...style, bold: true }));
        break;
      case 'IMG':
        paragraph.addChildElement(await createImageNode(child as HTMLImageElement));
        break;
      case 'BR':
        paragraph.addChildElement(new TextRun({ text: '\n' }));
        break;
      default:
        paragraph.addChildElement(await createTextNode(child as HTMLElement, style));
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
async function createChildNode(node: HTMLElement, style: { [key: string]: any } = {}): Promise<TextRun | ImageRun> {
  let childNode: TextRun | ImageRun = new TextRun({ text: '' });
  for (const child of Array.from(node.childNodes)) {
    switch (child.nodeName) {
      case 'STRONG':
        childNode = await createChildNode(child as HTMLElement, { ...style, bold: true });
        break;
      case 'IMG':
        childNode = await createImageNode(child as HTMLImageElement);
        break;
      case 'BR':
        childNode = new TextRun({ text: '\n' });
        break;
      default:
        childNode = await createTextNode(child as HTMLElement, style);
    }
  }
  return childNode;
}

/**
 * 创建标题元素
 * @param {HTMLElement} node - 输入的HTML元素
 * @returns {Promise<Table>} - 创建的标题对象
 */
async function createTableNode(node: HTMLElement): Promise<Table> {
  const rows = [];
  for (const tr of Array.from(node.querySelectorAll('tr'))) {
    const cells = [] as TableCell[];
    for (const td of Array.from(tr.querySelectorAll('td, th'))) {
      const cellChildren = [];
      for (const child of Array.from(td.childNodes)) {
        cellChildren.push(new Paragraph({ spacing: { line: 20 * 20 }, indent: { firstLine: 20 }, children: [new TextRun({ text: child.textContent || '', size: 18 })] }));
      }
      const size = td.getAttribute('width') && td.getAttribute('width') !== 'auto' ? parseInt(td.getAttribute('width') as string) : 100;
      // 处理单元格边框
      const border = { size: 1, color: '#000000', style: 'single' } as IBorderOptions;
      const cell = new TableCell({
        children: cellChildren as any,
        columnSpan: td.getAttribute('colspan') ? parseInt(td.getAttribute('colspan') as string) : undefined,
        rowSpan: td.getAttribute('rowspan') ? parseInt(td.getAttribute('rowspan') as string) : undefined,
        width: { size, type: 'auto' },
        borders: { top: border, bottom: border, left: border, right: border },
      });
      cells.push(cell);
    }
    rows.push(new TableRow({ children: cells }));
  }
  return new Table({ rows, width: { size: 100, type: 'pct' } });
}

/**
 * 创建列表节点
 * @param node - HTML列表元素(ul或ol)
 * @returns 由段落组成的数组，每个段落代表一个列表项
 */
async function createListNode(node: HTMLElement) {
  const list = [];
  for (const li of Array.from(node.childNodes)) {
    if (li.textContent !== '') {
      const numbering = li.textContent === '\n' ? undefined : { reference: node.nodeName === 'UL' ? 'bullet-points' : 'numbered-list', level: 0 };
      const listItem = await createParagraphNode(li as HTMLElement, new Paragraph({ numbering }));
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
async function createImageNode(imgNode: HTMLImageElement): Promise<ImageRun> {
  const imagePath = imgNode.getAttribute('src') ?? '';
  let imageData: any;
  if (imagePath.startsWith('data:image')) {
    // 处理base64图片
    const base64Data = imagePath.split(',')[1];
    // 将base64字符串转换为Uint8Array
    imageData = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));
  } else {
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
  const width = parseInt(imgNode.getAttribute('width') ?? '650');
  const height = parseInt(imgNode.getAttribute('height') ?? '280');
  // 处理图片，使用docx库中的ImageRun组件，需要将图片转换为Uint8Array或Blob对象，这里使用base64字符串作为示例，实际使用时需要根据实际情况进行处理
  return new ImageRun({ data: imageData, transformation: { width, height } } as any);
}

/**
 * 创建标题元素
 * @param {HTMLElement} node - 输入的HTML元素
 * @returns {Promise<Paragraph>} - 创建的标题对象
 */
async function createHeadingNode(node: HTMLElement): Promise<Paragraph> {
  const style = getTextStyle(node);
  const level = parseInt(node.nodeName.substring(1));
  return new Paragraph({
    alignment: style.alignment as any,
    heading: (HeadingLevel.HEADING_1 + (level - 1)) as any,
    children: [new TextRun({ text: node.textContent as string, bold: true })],
  });
}

/**
 * 创建文本元素
 * @param {HTMLElement} node - 输入的HTML元素
 * @param {TextStyle} [style] - 文本样式
 * @returns {Promise<TextRun>} - 创建的文本对象
 */
async function createTextNode(node: HTMLElement, style: ReturnType<typeof getTextStyle> = {} as ReturnType<typeof getTextStyle>): Promise<TextRun> {
  return new TextRun({
    text: node.textContent || '',
    font: style.font,
    size: style.size,
    color: style.color,
    bold: style.bold,
    italics: style.italics,
    underline: style.underline as any,
  });
}

export default htmlToDocx;
