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

"""Ground truth: tables (widths, position, margins, borders, merged cells, alignment, heights, styles)."""
import sys
from docx import build, p, r, default_styles

LONG = '問い合わせ対応（月20件まで。超過分は別途見積りとし、甲乙協議の上で決定する。）'
EDGES = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')


def borders(val='single', sz=4, edges=EDGES, tag='tblBorders'):
    return '<w:%s>%s</w:%s>' % (tag, ''.join(
        '<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="auto"/>' % (e, val, sz) for e in edges), tag)


def tc(content, width, tcpr='', ppr='', rpr=''):
    if not content.startswith('<w:'):
        content = p(r(content, rpr), ppr)
    return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>' % (width, tcpr, content)


def tr(cells, trpr=''):
    return '<w:tr>%s%s</w:tr>' % ('<w:trPr>%s</w:trPr>' % trpr if trpr else '', ''.join(cells))


def tbl(grid, rows, tblpr=None):
    if tblpr is None:
        tblpr = '<w:tblW w:w="0" w:type="auto"/>' + borders()
    return '<w:tbl><w:tblPr>%s</w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (
        tblpr, ''.join('<w:gridCol w:w="%d"/>' % g for g in grid), ''.join(rows))


STYLES = default_styles(extra=(
    '<w:style w:type="table" w:styleId="Fancy"><w:name w:val="Fancy"/><w:basedOn w:val="a1"/><w:tblPr>'
    '<w:tblStyleRowBandSize w:val="1"/>' + borders() + '</w:tblPr>'
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr>'
    '<w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>'
    '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr>'
    '</w:tblStylePr></w:style>'))


def body():
    b = []

    def label(text):
        b.append(p(r(text)))

    label('T01 基本の表と段落配置')
    b.append(tbl([2000, 3000, 3494], [
        tr([tc('左', 2000), tc('中央', 3000, ppr='<w:jc w:val="center"/>'), tc('右', 3494, ppr='<w:jc w:val="right"/>')]),
        tr([tc('一行', 2000), tc(LONG, 3000), tc('末尾', 3494)]),
    ]))
    label('T02 中央配置の表')
    b.append(tbl([2000, 2000], [tr([tc('甲', 2000), tc('乙', 2000)])],
                 '<w:tblW w:w="4000" w:type="dxa"/><w:jc w:val="center"/>' + borders()))
    label('T03 右配置の表')
    b.append(tbl([2000, 2000], [tr([tc('甲', 2000), tc('乙', 2000)])],
                 '<w:tblW w:w="4000" w:type="dxa"/><w:jc w:val="right"/>' + borders()))
    label('T04 インデントのある表')
    b.append(tbl([2000, 2000], [tr([tc('甲', 2000), tc('乙', 2000)])],
                 '<w:tblW w:w="4000" w:type="dxa"/><w:tblInd w:w="840" w:type="dxa"/>' + borders()))
    label('T05 セル余白（表全体とセル個別）')
    b.append(tbl([3000, 3000], [
        tr([tc('余白広い', 3000), tc('個別余白', 3000, '<w:tcMar><w:top w:w="200" w:type="dxa"/>'
                                 '<w:left w:w="600" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/></w:tcMar>')]),
        tr([tc('次の行', 3000), tc('次の行', 3000)]),
    ], '<w:tblW w:w="0" w:type="auto"/>' + borders() + '<w:tblCellMar><w:left w:w="300" w:type="dxa"/>'
       '<w:right w:w="300" w:type="dxa"/></w:tblCellMar>'))
    label('T06 行の高さと縦位置')
    b.append(tbl([2000, 2000, 2000], [
        tr([tc('上', 2000), tc('中央', 2000, '<w:vAlign w:val="center"/>'), tc('下', 2000, '<w:vAlign w:val="bottom"/>')],
           '<w:trHeight w:val="1000"/>'),
        tr([tc('固定', 2000), tc('固定中央', 2000, '<w:vAlign w:val="center"/>'), tc('固定下', 2000, '<w:vAlign w:val="bottom"/>')],
           '<w:trHeight w:val="800" w:hRule="exact"/>'),
        tr([tc('通常', 2000), tc('通常', 2000), tc('通常', 2000)]),
    ]))
    label('T07 結合セル')
    b.append(tbl([2000, 2000, 2000], [
        tr([tc('横結合', 4000, '<w:gridSpan w:val="2"/>'), tc('右上', 2000)]),
        tr([tc('縦結合', 2000, '<w:vMerge w:val="restart"/><w:vAlign w:val="center"/>'), tc('中1', 2000), tc('右1', 2000)]),
        tr([tc('', 2000, '<w:vMerge/>'), tc('中2', 2000), tc('右2', 2000)]),
        tr([tc('', 2000, '<w:vMerge/>'), tc('中3', 2000), tc('右3', 2000)]),
    ]))
    label('T08 罫線なしの表')
    b.append(tbl([2000, 2000], [tr([tc('甲', 2000), tc('乙', 2000)]), tr([tc('丙', 2000), tc('丁', 2000)])],
                 '<w:tblW w:w="0" w:type="auto"/>'))
    label('T09 太い罫線と二重線')
    b.append(tbl([2000, 2000], [
        tr([tc('太線', 2000), tc('太線', 2000)]),
        tr([tc('太線2', 2000), tc('太線2', 2000)]),
    ], '<w:tblW w:w="0" w:type="auto"/>' + borders(sz=24)))
    b.append(p(''))
    b.append(tbl([2000, 2000], [
        tr([tc('二重線', 2000), tc('セル個別', 2000, borders('single', 18, ('top', 'left', 'bottom', 'right'), 'tcBorders'))]),
        tr([tc('二重線2', 2000), tc('二重線2', 2000)]),
    ], '<w:tblW w:w="0" w:type="auto"/>' + borders('double', 6)))
    label('T10 入れ子の表')
    inner = tbl([1500, 1500], [tr([tc('内1', 1500), tc('内2', 1500)]), tr([tc('内3', 1500), tc('内4', 1500)])])
    b.append(tbl([2000, 4000], [tr([tc('外', 2000), tc(p(r('外の段落')) + inner + p(r('表の後の段落')), 4000)])]))
    label('T11 セル内の複数段落と段落間隔')
    b.append(tbl([3000, 3000], [tr([
        tc(p(r('段落1')) + p(r('段落2'), '<w:spacing w:before="120" w:after="120"/>') + p(r('段落3')), 3000),
        tc(p(r('字下げ'), '<w:ind w:left="420"/>') + p(r('9ポイント', '<w:sz w:val="18"/>'))
           + p(r('14ポイント', '<w:sz w:val="28"/>')), 3000)])]))
    label('T12 表スタイル（タイトル行と縞模様）')
    b.append(tbl([3000, 3000], [
        tr([tc('見出し1', 3000), tc('見出し2', 3000)]),
        tr([tc('縞1', 3000), tc('値1', 3000)]),
        tr([tc('縞2', 3000), tc('値2', 3000)]),
        tr([tc('縞3', 3000), tc('値3', 3000)]),
    ], '<w:tblStyle w:val="Fancy"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" '
       'w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'))
    label('T13 連続する表')
    b.append(tbl([2000, 2000], [tr([tc('表A', 2000), tc('表A', 2000)])]))
    b.append(p(''))
    b.append(tbl([3000, 3000], [tr([tc('表B', 3000), tc('表B', 3000)])]))
    label('T14 幅が本文より広い表と長い英単語')
    b.append(tbl([1200, 7304], [tr([tc('Supercalifragilistic', 1200), tc('A quick brown fox jumps over the lazy dog, and runs away. ' * 3, 7304)])]))
    label('末尾の段落')
    return ''.join(b)


build(sys.argv[1], body(), styles=STYLES)
