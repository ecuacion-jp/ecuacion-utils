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

"""Ground truth: character formatting, fonts per script, and the precedence of styles."""
import sys
from docx import build, p, r, fonts, default_styles

LONG = '甲及び乙は、本契約の履行に関連して知り得た相手方の技術上又は営業上の情報を、相手方の事前の書面による承諾なく第三者に開示してはならない。'

STYLES = default_styles(extra=(
    '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:basedOn w:val="a"/>'
    '<w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/>'
    '<w:rPr><w:b w:val="0"/><w:i/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="Small"><w:name w:val="Small"/>'
    '<w:rPr><w:color w:val="FF0000"/><w:sz w:val="18"/></w:rPr></w:style>'))


def body():
    b = []
    left = '<w:jc w:val="left"/>'

    def add(inner, ppr=left):
        b.append(p(inner, ppr))

    add(r('C01 サイズ') + r('8pt', '<w:sz w:val="16"/>') + r('10.5pt') + r('14pt', '<w:sz w:val="28"/>')
        + r('20pt', '<w:sz w:val="40"/>') + r('末尾'))
    add(r('C02 ') + r('太字Bold', '<w:b/>') + r('斜体Italic', '<w:i/>') + r('太字斜体BoldItalic', '<w:b/><w:i/>') + r('標準'))
    add(r('C03 ') + r('下線', '<w:u w:val="single"/>') + r('二重下線', '<w:u w:val="double"/>')
        + r('取り消し線', '<w:strike/>') + r('二重取り消し線', '<w:dstrike/>') + r('標準'))
    add(r('C04 ') + r('赤字', '<w:color w:val="FF0000"/>') + r('蛍光ペン', '<w:highlight w:val="yellow"/>')
        + r('網かけ', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>') + r('標準'))
    add(r('C05 上付きx') + r('2', '<w:vertAlign w:val="superscript"/>') + r(' 下付きH')
        + r('2', '<w:vertAlign w:val="subscript"/>') + r('O 注') + r('※1', '<w:vertAlign w:val="superscript"/>')
        + r('標準'))
    add(r('C06 ') + r('文字間隔広く', '<w:spacing w:val="20"/>') + r('文字間隔狭く', '<w:spacing w:val="-10"/>')
        + r('Spacing', '<w:spacing w:val="40"/>') + r('標準'))
    add(r('C07 ') + r('all caps text', '<w:caps/>') + r(' 標準'))
    add(r('C08 ') + r('Arial漢字MS明朝abc', fonts('Arial', 'ＭＳ 明朝'))
        + r('Timesメイリオxyz', fonts('Times New Roman', 'メイリオ'))
        + r('Century游ゴシック', fonts('Century', '游ゴシック')) + r('標準'))
    add(r('C09 ヒントなし“引用”‘a’…―') + r('ヒントあり“引用”‘a’…―', '<w:rFonts w:hint="eastAsia"/>') + r('標準'))
    add(r('C10 段落スタイル太字12pt'), '<w:pStyle w:val="Base"/>')
    add(r('C11 派生スタイル斜体') + r('文字スタイル9pt赤', '<w:rStyle w:val="Small"/>')
        + r('直接16pt', '<w:rStyle w:val="Small"/><w:sz w:val="32"/>') + r('直接太字', '<w:b/>'),
        '<w:pStyle w:val="Derived"/>')
    add(r('C12 ') + r(LONG[:20]) + r(LONG[20:40], '<w:b/>') + r(LONG[40:55], '<w:sz w:val="28"/>') + r(LONG[55:])
        + r(LONG, '<w:i/>') + r(LONG[:30], '<w:sz w:val="16"/>'), '<w:jc w:val="both"/>')
    add(r('C13 下線付きタブ') + '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:tab/></w:r>' + r('下線付き空白')
        + r('　　　　　', '<w:u w:val="single"/>') + r('末尾'))
    add(r('C14 半角ｶﾀｶﾅ全角ＡＢＣ１２３記号①②㈱〒※→←■□●○◆△▲☆★'))
    add(r('C15 改行しない空白A B C ハイフンA') + '<w:r><w:noBreakHyphen/></w:r>' + r('B 標準'))
    add(r('C16 太字の長文'))
    add(r(LONG + LONG, '<w:b/>'), '<w:jc w:val="both"/>')
    add(r('C17 游ゴシック太字とArial太字'))
    add(r(LONG + ' The quick brown fox jumps over the lazy dog. ' * 2, fonts('Arial', '游ゴシック') + '<w:b/>'),
        '<w:jc w:val="both"/>')
    add(r('C18 MS明朝14ptの長文'))
    add(r(LONG, fonts('Century', 'ＭＳ 明朝') + '<w:sz w:val="28"/>'), '<w:jc w:val="both"/>')
    add(r('C19 メイリオの長文'))
    add(r(LONG, fonts('Century', 'メイリオ')), '<w:jc w:val="both"/>')
    add(r('末尾の段落'))
    return ''.join(b)


build(sys.argv[1], body(), styles=STYLES)
