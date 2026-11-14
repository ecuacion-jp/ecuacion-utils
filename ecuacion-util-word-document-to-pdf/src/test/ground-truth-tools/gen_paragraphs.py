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

"""Ground truth: indents, alignment, tabs, numbering, paragraph spacing, borders, East Asian spacing."""
import sys
from docx import build, p, r

LONG = '甲及び乙は、本契約の履行に関連して知り得た相手方の技術上又は営業上の情報（以下「秘密情報」という。）を、相手方の事前の書面による承諾なく第三者に開示又は漏洩してはならない。'
MIXED = '本サービス（Service）の利用料金は月額1,200円（税込）とし、API経由のアクセスは1日あたり10,000回までとする。Userは24時間365日利用できるが、SLAの対象は平日9時から18時までとする。'
TAB = '<w:r><w:tab/></w:r>'
BR = '<w:r><w:br/></w:r>'


def tabs(*stops):
    return '<w:tabs>%s</w:tabs>' % ''.join(
        '<w:tab w:val="%s"%s w:pos="%d"/>' % (v, ' w:leader="%s"' % l if l else '', pos) for v, l, pos in stops)


def num(num_id, ilvl=0):
    return '<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>' % (ilvl, num_id)


def lvl(ilvl, fmt, text, left, hanging, suff=None, jc='left', rpr=''):
    return ('<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="%s"/>%s<w:lvlText w:val="%s"/>'
            '<w:lvlJc w:val="%s"/><w:pPr><w:ind w:left="%d" w:hanging="%d"/></w:pPr>%s</w:lvl>'
            % (ilvl, fmt, '<w:suff w:val="%s"/>' % suff if suff else '', text, jc, left, hanging,
               '<w:rPr>%s</w:rPr>' % rpr if rpr else ''))


SINGLE = [  # numId 2.. : (format, text, suffix, justification)
    ('decimal', '%1.', None, 'left'),
    ('decimalFullWidth', '%1', None, 'left'),
    ('decimalEnclosedCircle', '%1', None, 'left'),
    ('upperRoman', '%1.', None, 'right'),
    ('lowerRoman', '%1)', None, 'left'),
    ('upperLetter', '%1.', None, 'left'),
    ('lowerLetter', '(%1)', None, 'left'),
    ('japaneseCounting', '第%1章', 'space', 'left'),
    ('aiueoFullWidth', '%1', None, 'left'),
    ('irohaFullWidth', '%1', None, 'left'),
    ('bullet', '●', None, 'left'),
    ('decimal', '%1)', 'nothing', 'left'),
    ('decimalZero', '%1', 'space', 'left'),
]


def numbering():
    out = ['<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>'
           + lvl(0, 'decimal', '%1', 420, 420) + lvl(1, 'decimal', '%1.%2', 840, 630)
           + lvl(2, 'decimal', '%1.%2.%3', 1470, 840) + '</w:abstractNum>']
    for i, (fmt, text, suff, jc) in enumerate(SINGLE):
        out.append('<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="hybridMultilevel"/>%s</w:abstractNum>'
                   % (i + 1, lvl(0, fmt, text, 630, 630, suff, jc)))
    for i in range(len(SINGLE) + 1):
        out.append('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>' % (i + 1, i))
    return ''.join(out)


def body():
    b = []

    def add(label, text, ppr='', inner=None):
        b.append(p(inner if inner is not None else r(label + ' ' + text), ppr))

    left = '<w:jc w:val="left"/>'
    add('I01', '左インデント', left + '') if False else None
    add('I01', '左インデント420', '<w:ind w:left="420"/>' + left)
    add('I02', LONG, '<w:ind w:left="420" w:right="840"/>')
    add('I03', LONG, '<w:ind w:firstLine="210"/>')
    add('I04', LONG, '<w:ind w:left="420" w:hanging="420"/>')
    add('I05', LONG, '<w:ind w:leftChars="200" w:left="420"/>')
    add('I06', LONG, '<w:ind w:firstLineChars="100" w:firstLine="210"/>')
    add('I07', LONG, '<w:ind w:leftChars="100" w:left="210" w:hangingChars="100" w:hanging="210"/>')
    add('A01', '中央揃え', '<w:jc w:val="center"/>')
    add('A02', '右揃え', '<w:jc w:val="right"/>')
    add('A03', LONG, '<w:jc w:val="both"/>')
    add('A04', '均等割り付け', '<w:jc w:val="distribute"/>')
    add('A05', '', '<w:jc w:val="both"/>', r('A05 ' + LONG[:30]) + BR + r(LONG))
    add('A06', LONG, left)
    add('T01', '', left, r('T01') + TAB + r('既定') + TAB + r('タブ') + TAB + r('abc'))
    add('T02', '', tabs(('left', None, 2000), ('left', None, 5000)) + left,
        r('T02') + TAB + r('左2000') + TAB + r('左5000'))
    add('T03', '', tabs(('center', None, 4252), ('right', None, 8504)) + left,
        r('T03') + TAB + r('中央') + TAB + r('右端'))
    add('T04', '', tabs(('right', 'dot', 8504)) + left, r('T04 点線リーダー') + TAB + r('12'))
    add('T05', '', tabs(('right', 'hyphen', 6000), ('right', 'underscore', 8504)) + left,
        r('T05') + TAB + r('ハイフン') + TAB + r('下線'))
    add('T06', '', tabs(('left', None, 3000)) + '<w:ind w:left="840" w:hanging="840"/>' + left,
        r('T06') + TAB + r('ぶら下げ位置') + TAB + r('左3000'))
    for level, text in [(0, '第一'), (1, '第一の一'), (1, '第一の二'), (2, '第一の二の一'), (0, '第二'), (1, '第二の一')]:
        add('', '', num(1, level) + left, r('N多階層' + text))
    for i, (fmt, _, _, _) in enumerate(SINGLE):
        for k in range(3):
            add('', '', num(i + 2) + left, r('N%s項目%d' % (fmt, k + 1)))
    add('', '', num(2) + left, r('N長い項目' + LONG))
    add('S01', '同じスタイル間隔なし1', '<w:spacing w:before="240" w:after="240"/><w:contextualSpacing/>' + left)
    add('S02', '同じスタイル間隔なし2', '<w:spacing w:before="240" w:after="240"/><w:contextualSpacing/>' + left)
    add('S03', '通常の間隔', '<w:spacing w:before="240" w:after="240"/>' + left)
    add('S04', '次の段落', left)
    bdr = '<w:%s w:val="single" w:sz="8" w:space="1" w:color="000000"/>'
    add('B01', '下罫線', '<w:pBdr>%s</w:pBdr>' % (bdr % 'bottom') + left)
    add('B02', '次の段落', left)
    add('B03', '囲み罫線と網かけ', '<w:pBdr>%s</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/>'
        % ''.join(bdr % e for e in ('top', 'left', 'bottom', 'right')) + left)
    add('B04', '次の段落', left)
    add('E01', MIXED)
    add('E02', MIXED, '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>')
    add('E03', MIXED, '<w:wordWrap w:val="0"/>')
    add('E04', MIXED, '<w:overflowPunct w:val="0"/>')
    add('E05', '「」（）、。「括弧」（括弧）「「二重」」。、。）」（「' * 4)
    add('E06', 'The quick brown fox jumps over the lazy dog. ' * 5)
    add('E07', 'Supercalifragilisticexpialidocious' * 4, left)
    return ''.join(b)


build(sys.argv[1], body(), numbering=numbering())
