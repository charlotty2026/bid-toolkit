#!/usr/bin/env python3
"""
给Word文档添加斜向文字水印。
适用于标书/投标文件的"仅供参考""机密""样本"等水印标记。

用法:
    from bid_toolkit.scripts.watermark import add_watermark
    add_watermark('输入.docx', '输出.docx', text='仅供参考')
"""

from docx import Document
from lxml import etree

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_V = 'urn:schemas-microsoft-com:vml'
NS_O = 'urn:schemas-microsoft-com:office:office'
NS_W10 = 'urn:schemas-microsoft-com:office:word'


def _qn(tag, nsp=NS_W):
    return f'{{{nsp}}}{tag}'


def add_watermark(input_path: str, output_path: str, text: str = '仅供参考',
                  font: str = '微软雅黑', color: str = '#808080',
                  opacity: str = '0.5') -> bool:
    """
    给Word文档所有页面添加斜向文字水印。

    Word水印 ≈ 藏在页眉里的VML旋转文字图形。
    通过lxml操作底层OpenXML，在每个section的header中插入水印SDT结构。

    Args:
        input_path:  输入docx路径
        output_path: 输出docx路径
        text:        水印文字，如"仅供参考""机密""样本"
        font:        字体名称，默认"微软雅黑"
        color:       颜色十六进制，默认"#808080"灰色
        opacity:     透明度0-1，默认"0.5"

    Returns:
        True表示成功

    Raises:
        FileNotFoundError: 输入文件不存在
        docx.opc.exceptions.PackageNotFoundError: 输入不是有效docx
    """
    doc = Document(input_path)

    for si, section in enumerate(doc.sections):
        header = section.header
        header_elem = header._element

        # 清空页眉原有内容
        for p in list(header_elem.findall(_qn('p'))):
            header_elem.remove(p)

        p_elem = etree.SubElement(header_elem, _qn('p'))

        # SDT — 结构化文档标签，告诉Word这是水印
        sdt = etree.SubElement(p_elem, _qn('sdt'))

        # SDT属性
        sdtPr = etree.SubElement(sdt, _qn('sdtPr'))
        etree.SubElement(sdtPr, _qn('id')).set(_qn('val'), str(2147483647 - si))
        etree.SubElement(sdtPr, _qn('lock')).set(_qn('val'), 'sdtLocked')

        docPartObj = etree.SubElement(sdtPr, _qn('docPartObj'))
        etree.SubElement(docPartObj, _qn('docPartGallery')).set(_qn('val'), 'Watermarks')
        etree.SubElement(docPartObj, _qn('docPartUnique')).set(_qn('val'), 'true')

        # SDT结束属性
        sdtEndPr = etree.SubElement(sdt, _qn('sdtEndPr'))
        rPr_end = etree.SubElement(sdtEndPr, _qn('rPr'))
        etree.SubElement(rPr_end, _qn('rStyle')).set(_qn('val'), 'Normal')

        # SDT内容
        sdtContent = etree.SubElement(sdt, _qn('sdtContent'))
        r = etree.SubElement(sdtContent, _qn('r'))

        # 禁用拼写检查
        rPr = etree.SubElement(r, _qn('rPr'))
        etree.SubElement(rPr, _qn('noProof'))

        # VML图形容器
        pict = etree.SubElement(r, _qn('pict'))

        # ShapeType — VML图形类型定义
        st = etree.SubElement(pict, _qn('shapetype', NS_V))
        st.set(_qn('id', NS_V), f'_x0000_t136_{si}')
        st.set(_qn('coordsize', NS_V), '1600,21600')
        st.set(_qn('spt', NS_O), '136')
        st.set(_qn('adj', NS_O), '10800')

        path_v = etree.SubElement(st, _qn('path', NS_V))
        path_v.set(_qn('textpathok', NS_V), 't')

        tp = etree.SubElement(st, _qn('textpath', NS_V))
        tp.set(_qn('on', NS_V), 't')
        tp.set(_qn('fitshape', NS_V), 't')

        lock_v = etree.SubElement(st, _qn('lock', NS_O))
        lock_v.set(_qn('ext', NS_V), 'edit')
        lock_v.set(_qn('shapetype', NS_O), 't')

        # Shape — 实际水印图形
        uid = 357476642 + si
        shape = etree.SubElement(pict, _qn('shape', NS_V))
        shape.set(_qn('id', NS_V), f'PowerPlusWaterMarkObject{uid}')
        shape.set(_qn('type', NS_V), f'#_x0000_t136_{si}')
        shape.set(_qn('style', NS_V),
                  'position:absolute;left:0;text-align:center;'
                  'margin-left:0;margin-top:0;'
                  'width:420pt;height:297pt;'
                  'rotation:-315;z-index:-251656192;mso-wrap-edited:f')
        shape.set(_qn('fillcolor', NS_V), color)

        fill = etree.SubElement(shape, _qn('fill', NS_V))
        fill.set(_qn('opacity', NS_V), opacity)

        # 水印文字
        textPath = etree.SubElement(shape, _qn('textpath', NS_V))
        textPath.set(_qn('style', NS_V), f'font-family:"{font}";font-size:1pt')
        textPath.set(_qn('string', NS_V), text)

        # 文本框锚定
        tw = etree.SubElement(shape, _qn('textwrap', NS_W10))
        tw.set(_qn('anchorx', NS_W10), 'margin')
        tw.set(_qn('anchory', NS_W10), 'margin')

    doc.save(output_path)
    return True
