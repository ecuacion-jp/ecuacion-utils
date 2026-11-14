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

"""Ground truth: page break control (keep with next, keep lines, widow / orphan, spacing at page top).

The page has 36 lines of the 18pt grid. Each case starts on a new page, fills some lines, and puts
the paragraphs under test at the bottom of the page.
"""
import sys
from docx import build, p, r

WIDOW = '<w:widowControl/>'
NO_WIDOW = '<w:widowControl w:val="0"/>'
PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>'


def text(label, lines):
    """Returns text which takes the number of lines (40 characters fit in a line)."""
    return (label + 'あ' * 200)[:40 * (lines - 1) + 20]


def case(b, name, fillers):
    for i in range(fillers):
        b.append(p(r('%s行%02d' % (name, i + 1)), '<w:pageBreakBefore/>' if i == 0 else ''))


def cell(text_, width):
    return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>' % (width, p(r(text_)))


def table(name, rows, header=False, cant_split=False, tall_row=None):
    borders = ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>' % e
                      for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    out = ['<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>%s</w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="6000"/></w:tblGrid>' % borders]
    for i in range(rows):
        trpr = ''
        if header and i == 0:
            trpr = '<w:trPr><w:tblHeader/></w:trPr>'
        elif cant_split and i == tall_row:
            trpr = '<w:trPr><w:cantSplit/></w:trPr>'
        content = '%s表%02d' % (name, i + 1)
        second = content + '内容'
        if i == tall_row:
            second = (content + 'い' * 200)[:28 * 3 + 10]
        out.append('<w:tr>%s%s%s</w:tr>' % (trpr, cell(content, 2000), cell(second, 6000)))
    out.append('</w:tbl>')
    return ''.join(out)


def body():
    b = []
    case(b, 'A', 35)
    b.append(p(r('A見出し次と分離しない'), '<w:keepNext/>'))
    b.append(p(r(text('A本文', 2)), NO_WIDOW))

    case(b, 'B', 35)
    b.append(p(r(text('B孤立行制御3行', 3)), WIDOW))

    case(b, 'C', 34)
    b.append(p(r(text('C孤立行制御3行', 3)), WIDOW))

    case(b, 'D', 34)
    b.append(p(r(text('D孤立行制御4行', 4)), WIDOW))

    case(b, 'E', 33)
    b.append(p(r(text('E孤立行制御4行', 4)), WIDOW))

    case(b, 'F', 34)
    b.append(p(r(text('F段落を分割しない', 4)), NO_WIDOW + '<w:keepLines/>'))

    case(b, 'G', 35)
    b.append(p(r(text('G制御なし3行', 3)), NO_WIDOW))

    case(b, 'H', 35)
    b.append(p(r('H段落前1行'), '<w:spacing w:beforeLines="100" w:before="360"/>'))
    b.append(p(r('H次の段落')))

    case(b, 'I', 3)
    b.append(p(r('I改ページ前の段落前間隔'), '<w:pageBreakBefore/><w:spacing w:beforeLines="100" w:before="360"/>'))
    b.append(p(r('I次の段落')))

    case(b, 'J', 33)
    b.append(p(r('J見出し1'), '<w:keepNext/>'))
    b.append(p(r('J見出し2'), '<w:keepNext/>'))
    b.append(p(r(text('J本文', 2)), NO_WIDOW))

    case(b, 'K', 3)
    b.append(p(r('K改ページ前') + PAGE_BREAK + r('K改ページ後')))
    b.append(p(r('K次の段落')))
    b.append(p(r('K段落末尾の改ページ') + PAGE_BREAK))
    b.append(p(r('K改ページ直後の段落'), '<w:spacing w:beforeLines="100" w:before="360"/>'))

    case(b, 'L', 34)
    b.append(table('L', 5))
    b.append(p(r('L表の後')))

    case(b, 'M', 33)
    b.append(p(r('M見出し次と分離しない'), '<w:keepNext/>'))
    b.append(table('M', 4))
    b.append(p(r('M表の後')))

    case(b, 'N', 30)
    b.append(table('N', 12, header=True))
    b.append(p(r('N表の後')))

    case(b, 'O', 33)
    b.append(table('O', 3, tall_row=1))
    b.append(p(r('O表の後')))

    case(b, 'P', 33)
    b.append(table('P', 3, cant_split=True, tall_row=1))
    b.append(p(r('P表の後')))
    return ''.join(b)


build(sys.argv[1], body())
