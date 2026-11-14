#
# Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

"""Ground truth: inline pictures (the height of the lines which have them)."""
import struct
import sys
import zlib
from docx import build, p, r

DRAWING_NS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
              'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')


def png(width, height, rgb):
    """Returns a PNG picture filled with the color."""
    def chunk(kind, data):
        body = kind + data
        return struct.pack('>I', len(data)) + body + struct.pack('>I', zlib.crc32(body) & 0xffffffff)
    row = b'\x00' + bytes(rgb) * width
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(row * height)) + chunk(b'IEND', b''))


def picture(rid, width, height, index):
    """Returns a run with an inline picture of the size in points."""
    size = 'cx="%d" cy="%d"' % (width * 12700, height * 12700)
    return ('<w:r><w:drawing><wp:inline %s distT="0" distB="0" distL="0" distR="0"><wp:extent %s/>'
            '<wp:docPr id="%d" name="Picture %d"/><a:graphic>'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="%d" name="Picture %d"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext %s/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
            '</a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
            % (DRAWING_NS, size, index, index, index, index, rid, size))


def body():
    b = []
    left = '<w:jc w:val="left"/>'
    counter = [0]

    def pic(width, height):
        counter[0] += 1
        return picture('rId50', width, height, counter[0])

    b.append(p(r('G01 行内画像の前') + pic(60, 30) + r('行内画像の後'), left))
    b.append(p(r('G02 次の段落'), left))
    b.append(p(pic(100, 50), left))
    b.append(p(r('G03 画像だけの段落の次'), left))
    b.append(p(r('G04 小さい画像') + pic(20, 10) + r('の後'), left))
    b.append(p(r('G05 次の段落'), left))
    b.append(p(r('G06 グリッドなし') + pic(60, 30) + r('の後'), '<w:snapToGrid w:val="0"/>' + left))
    b.append(p(r('G07 次の段落'), left))
    b.append(p(r('G08 中央') + pic(60, 40) + r('揃え'), '<w:jc w:val="center"/>'))
    b.append(p(r('G09 次の段落'), left))
    borders = ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>' % e
                      for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    b.append('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>%s</w:tblBorders></w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>'
             '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>%s</w:tc>'
             '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'
             % (borders, p(r('G10 表の中') + pic(40, 30)), p(r('G10 隣のセル'))))
    b.append(p(r('G11 表の次の段落'), left))
    b.append(p(r('G12 固定行間20pt') + pic(60, 30) + r('の後'), '<w:spacing w:line="400" w:lineRule="exact"/>' + left))
    b.append(p(r('G13 末尾の段落'), left))
    return ''.join(b)


build(sys.argv[1], body(), media={'rId50': png(40, 20, (0x44, 0x72, 0xC4))})
