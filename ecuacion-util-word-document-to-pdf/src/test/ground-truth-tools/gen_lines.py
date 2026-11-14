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

"""Ground truth: line heights and baselines for fonts, sizes, line spacing rules and the grid."""
import sys
from docx import build, p, r, fld, fonts, SECT_A4_GRID, SECT_A4_NOGRID

JA = 'あいう漢字'
cases = [
    ('Latin Century abc', '', ''),
    (JA + '游明朝', '', ''),
    ('Latin Arial abc', '', fonts('Arial', '游明朝')),
    ('Latin Times abc', '', fonts('Times New Roman', '游明朝') + '<w:sz w:val="24"/>'),
    ('mixed abc' + JA, '', ''),
    (JA + '12pt', '', '<w:sz w:val="24"/>'),
    (JA + '14pt', '', '<w:sz w:val="28"/>'),
    (JA + '9pt', '', '<w:sz w:val="18"/>'),
    (JA + 'MS明朝', '', fonts('Century', 'ＭＳ 明朝')),
    (JA + 'MS明朝14pt', '', fonts('Century', 'ＭＳ 明朝') + '<w:sz w:val="28"/>'),
    (JA + 'メイリオ', '', fonts('Century', 'メイリオ')),
    (JA + '游ゴシック', '', fonts('Arial', '游ゴシック')),
    (JA + 'grid off', '<w:snapToGrid w:val="0"/>', ''),
    ('Latin grid off abc', '<w:snapToGrid w:val="0"/>', ''),
    (JA + '1.5行', '<w:spacing w:line="360" w:lineRule="auto"/>', ''),
    (JA + '2行', '<w:spacing w:line="480" w:lineRule="auto"/>', ''),
    (JA + '固定20pt', '<w:spacing w:line="400" w:lineRule="exact"/>', ''),
    (JA + '固定12pt', '<w:spacing w:line="240" w:lineRule="exact"/>', ''),
    (JA + '最小24pt', '<w:spacing w:line="480" w:lineRule="atLeast"/>', ''),
    (JA + '最小10pt', '<w:spacing w:line="200" w:lineRule="atLeast"/>', ''),
    (JA + '1.5行grid off', '<w:snapToGrid w:val="0"/><w:spacing w:line="360" w:lineRule="auto"/>', ''),
    (JA + '前6後6', '<w:spacing w:before="120" w:after="120"/>', ''),
    (JA + '前1行後0.5行', '<w:spacing w:beforeLines="100" w:before="360" w:afterLines="50" w:after="180"/>', ''),
    (JA + '末尾', '', ''),
]


def body():
    out = []
    for i, (text, ppr, rpr) in enumerate(cases):
        out.append(p(r('P%02d ' % (i + 1) + text, rpr), ppr))
        out.append(p(r('次の行%02d' % (i + 1))))
    return ''.join(out)


footer = p(r('フッター ') + fld('PAGE', '1'), '<w:jc w:val="center"/>') + p(r('footer latin only'), '<w:jc w:val="center"/>')
header = p(r('header latin only')) + p(r('ヘッダー和文'))
build(sys.argv[1], body(), sect=SECT_A4_GRID if sys.argv[2] == 'grid' else SECT_A4_NOGRID, header=header, footer=footer)
