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

"""Builds minimal .docx files Word accepts (element order must follow the schema)."""
import zipfile

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
CT_MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'

SECT_A4_GRID = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" '
                'w:left="1701" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/>'
                '<w:docGrid w:type="lines" w:linePitch="360"/>')
SECT_A4_NOGRID = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" '
                  'w:left="1701" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/>')


def fonts(ascii_font, ea_font):
    return '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>' % (ascii_font, ea_font, ascii_font)


def default_styles(ascii_font='Century', ea_font='游明朝', size=21, extra=''):
    return ('<w:docDefaults><w:rPrDefault><w:rPr>%s<w:kern w:val="2"/><w:sz w:val="%d"/>'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
            '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
            '<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:tblPr>'
            '<w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
            '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
            '</w:tblCellMar></w:tblPr></w:style>%s') % (fonts(ascii_font, ea_font), size, extra)


SETTINGS = ('<w:defaultTabStop w:val="840"/><w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>')


def r(text, rpr=''):
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % ('<w:rPr>' + rpr + '</w:rPr>' if rpr else '', text)


def p(inner, ppr=''):
    return '<w:p>%s%s</w:p>' % ('<w:pPr>' + ppr + '</w:pPr>' if ppr else '', inner)


def fld(instr, cached):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>%s</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>') % (instr, cached)


def build(path, body, styles=None, sect=SECT_A4_GRID, header=None, footer=None, numbering=None,
          settings=SETTINGS, stories=None, media=None):
    """stories: {relationship id: ('header' | 'footer', blocks)} for the parts which section
    properties in the body (or sect) refer to by themselves.
    media: {relationship id: PNG bytes} for the pictures which drawings refer to."""
    parts = {}
    rels = []
    overrides = []

    def part(name, kind, xml, rid, ctype):
        parts['word/' + name] = X + xml
        rels.append('<Relationship Id="%s" Type="%s/%s" Target="%s"/>' % (rid, R, kind, name))
        overrides.append('<Override PartName="/word/%s" ContentType="%s%s+xml"/>' % (name, CT_MAIN, ctype))

    part('styles.xml', 'styles', '<w:styles %s>%s</w:styles>' % (W, styles or default_styles()), 'rId1', 'styles')
    part('settings.xml', 'settings', '<w:settings %s>%s</w:settings>' % (W, settings), 'rId2', 'settings')
    if numbering:
        part('numbering.xml', 'numbering', '<w:numbering %s>%s</w:numbering>' % (W, numbering), 'rId3', 'numbering')
    refs = ''
    if header:
        part('header1.xml', 'header', '<w:hdr %s>%s</w:hdr>' % (W, header), 'rId10', 'header')
        refs += '<w:headerReference w:type="default" r:id="rId10"/>'
    if footer:
        part('footer1.xml', 'footer', '<w:ftr %s>%s</w:ftr>' % (W, footer), 'rId11', 'footer')
        refs += '<w:footerReference w:type="default" r:id="rId11"/>'
    for index, (rid, (kind, blocks)) in enumerate(sorted((stories or {}).items())):
        tag = 'hdr' if kind == 'header' else 'ftr'
        part('%s_%d.xml' % (kind, index), kind, '<w:%s %s>%s</w:%s>' % (tag, W, blocks, tag), rid, kind)
    for rid in sorted(media or {}):
        rels.append('<Relationship Id="%s" Type="%s/image" Target="media/%s.png"/>' % (rid, R, rid))
    document = X + '<w:document %s><w:body>%s<w:sectPr>%s%s</w:sectPr></w:body></w:document>' % (W, body, refs, sect)
    ct = (X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="png" ContentType="image/png"/>'
          '<Override PartName="/word/document.xml" ContentType="%sdocument.main+xml"/>%s</Types>'
          % (CT_MAIN, ''.join(overrides)))
    prels = (X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/></Relationships>' % R)
    drels = (X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
             % ''.join(rels))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', prels)
        z.writestr('word/_rels/document.xml.rels', drels)
        z.writestr('word/document.xml', document)
        for name, xml in parts.items():
            z.writestr(name, xml)
        for rid, data in (media or {}).items():
            z.writestr('word/media/%s.png' % rid, data)
