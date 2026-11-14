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

"""Ground truth: sections, page settings, headers / footers and page numbers."""
import sys
from docx import build, p, r, fld, SETTINGS

BREAK = '<w:r><w:br w:type="page"/></w:r>'
CENTER = '<w:jc w:val="center"/>'


def refs(**kinds):
    out = []
    for key, rid in kinds.items():
        tag, kind = key.split('_')
        out.append('<w:%sReference w:type="%s" r:id="%s"/>' % (tag, kind, rid))
    return ''.join(out)


def sect(refs_xml, size, margins, extra=''):
    return '%s%s%s<w:cols w:space="425"/>%s' % (refs_xml, size, margins, extra)


A4 = '<w:pgSz w:w="11906" w:h="16838"/>'
A4L = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
B5 = '<w:pgSz w:w="10318" w:h="14570"/>'
MAR = '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>'
MAR_WIDE = '<w:pgMar w:top="1440" w:right="2880" w:bottom="1440" w:left="2880" w:header="720" w:footer="720" w:gutter="0"/>'
MAR_NARROW = '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'
GRID = '<w:docGrid w:type="lines" w:linePitch="360"/>'


def page_footer(label):
    return p(r(label + ' ') + fld('PAGE', '1') + r(' / ') + fld('NUMPAGES', '1') + r(' セクション内')
             + fld('SECTIONPAGES', '1'), CENTER)


stories = {
    'rId20': ('header', p(r('S1 既定ヘッダー'))),
    'rId21': ('header', p(r('S1 先頭ページヘッダー'))),
    'rId22': ('header', p(r('S1 偶数ページヘッダー'), '<w:jc w:val="right"/>')),
    'rId23': ('footer', page_footer('S1既定')),
    'rId24': ('footer', page_footer('S1先頭')),
    'rId25': ('footer', page_footer('S1偶数')),
    'rId30': ('header', p(r('S2 横向きヘッダー'))),
    'rId31': ('footer', page_footer('S2')),
    'rId40': ('header', ''.join(p(r('S4 高いヘッダー%d行目' % (i + 1))) for i in range(6))),
    'rId41': ('footer', p(r('S4 フッター1行目'), CENTER) + page_footer('S4') + p(r('S4 フッター3行目'), CENTER)),
}


def body():
    b = []
    b.append(p(r('第1セクション 1ページ目（先頭ページ）')))
    b.append(p(r('本文') + BREAK))
    b.append(p(r('第1セクション 2ページ目（偶数）')))
    b.append(p(r('本文') + BREAK))
    b.append(p(r('第1セクション 3ページ目（奇数）')))
    b.append(p(r('本文') + BREAK))
    b.append(p(r('第1セクション 4ページ目（偶数）')))
    b.append(p(r('セクション最後の段落'), '<w:sectPr>%s</w:sectPr>' % sect(
        refs(header_default='rId20', header_first='rId21', header_even='rId22',
             footer_default='rId23', footer_first='rId24', footer_even='rId25'),
        A4, MAR, '<w:pgNumType w:fmt="lowerRoman"/><w:titlePg/>' + GRID)))
    b.append(p(r('第2セクション 横向き1ページ目')))
    for i in range(30):
        b.append(p(r('第2セクション本文%02d行目' % (i + 1))))
    b.append(p(r('セクション最後の段落'), '<w:sectPr>%s</w:sectPr>' % sect(
        refs(header_default='rId30', footer_default='rId31'), A4L, MAR_WIDE,
        '<w:pgNumType w:start="1"/>')))
    b.append(p(r('第3セクション 連続（ヘッダーは前セクションを継承）')))
    b.append(p(r('本文' * 40)))
    b.append(p(r('セクション最後の段落'), '<w:sectPr>%s</w:sectPr>' % sect(
        '', A4L, MAR_WIDE, '<w:type w:val="continuous"/>')))
    b.append(p(r('第4セクション B5 高いヘッダー')))
    for i in range(45):
        b.append(p(r('第4セクション本文%02d行目' % (i + 1))))
    b.append(p(r('文書の最後の段落')))
    return ''.join(b)


final = sect(refs(header_default='rId40', footer_default='rId41'), B5, MAR_NARROW,
             '<w:pgNumType w:fmt="upperRoman" w:start="5"/>')
build(sys.argv[1], body(), sect=final, stories=stories, settings=SETTINGS + '<w:evenAndOddHeaders/>')
